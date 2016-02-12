(ns excel-to-json.converter
  (:require [flatland.ordered.map :refer [ordered-map]]
            [clj-excel.core :as ce]
            [clojure.core.match :refer [match]])
  (:import [org.apache.poi.ss.usermodel DataFormatter Cell Workbook CreationHelper]))

(def ^:dynamic *evaluator*)

(def data-formatter (DataFormatter. java.util.Locale/US))

(defn conj*
  [s x]
  (conj (vec s) x))

(defn split-keys [k]
  (map keyword (clojure.string/split (name k) #"\.")))

(defn safe-keyword [k]
  (keyword (str (if (instance? Number k) (long k) k))))

(defn apply-format [cell]
  (.formatCellValue ^DataFormatter data-formatter cell *evaluator*))

(defn convert-value [value]
  (if-let [result (re-find #"^([-+]?[0-9]+)(\.[0-9]+)?$" value)]
    (if (last result)
      (BigDecimal. value)
      (. Integer parseInt value))
    (case (clojure.string/lower-case value)
                  "true" true
                  "false" false
                  value)))

(defn convert-cell [cell]
  (convert-value (apply-format cell)))

(defn convert-header [cell]
  (keyword (convert-cell cell)))

(defn is-blank? [cell]
  (or (= (.getCellType ^Cell cell) Cell/CELL_TYPE_BLANK)) (= (convert-cell cell) ""))

(defn with-index [cells]
  (into {} (map (fn [c] [(.getColumnIndex ^Cell c) c]) cells)))

(defn split-array [array-char cell]
  (->>
    (clojure.string/split (apply-format cell) (re-pattern array-char))
    (map clojure.string/trim)
    (remove empty?)
    (map convert-value)))

(defn parse-column-normal [header cell]
  [(split-keys (convert-header header)) (convert-cell cell)])

(defn parse-column-array [key-path array-char cell]
  [(split-keys key-path) (split-array (or array-char ",") cell)])

(defn parse-column [header cell]
  (match (re-find #"(.*)@(.)?" (apply-format header))
    nil (parse-column-normal header cell)
    [_ key-path array-char] (parse-column-array key-path array-char cell)))

(defn unpack-keys [header row]
  (let [indexed-header (with-index header)
        indexed-row (with-index row)]
    (reduce (fn [acc [i header]]
              (let [cell (get indexed-row i)]
                (if (or (is-blank? header) (nil? cell) (is-blank? cell))
                  acc
                  (let [[key-path value] (parse-column header cell)]
                    (assoc-in acc key-path value)))))
      (ordered-map) indexed-header)))

(defn non-empty-rows [rows]
  (filter
    (fn [row]
      (let [cell (first row)]
        (and
          (= (.getColumnIndex ^Cell cell) 0)
          (not (is-blank? cell)))))
    rows))

(defn headers-and-rows [sheet]
  (let [rows (non-empty-rows sheet)]
    [(first rows) (rest rows)]))

(defn ensure-ordered [m k]
  (if (nil? (k m)) (assoc m k (ordered-map)) m))

(defn blank? [value]
  (cond
   (integer? value) false
   :else (clojure.string/blank? value)))

(defn add-sheet-config [primary-key current-key sheets config]
  (reduce (fn [acc0 sheet]
            (let [[headers rows] (headers-and-rows sheet)
                  secondary-key (convert-header (second headers))
                  unpacked-rows (map #(unpack-keys headers %) rows)
                  grouped-rows (group-by primary-key unpacked-rows)
                  secondary-config (get grouped-rows (name current-key))]
              ;; TODO remove either primary or current key
              (reduce (fn [acc row]
                        (let [nested-key (get row secondary-key)
                              safe-nested-key (safe-keyword nested-key)
                              sub (dissoc row primary-key secondary-key)]
                          (if (empty? sub)
                            acc
                            (if (blank? nested-key)
                              (update-in acc [secondary-key] conj* sub)
                              (assoc-in (ensure-ordered acc secondary-key)
                                        [secondary-key safe-nested-key] sub)))))
                      acc0 secondary-config)))
          config sheets))

(defn name-and-type-from-sheet [sheet]
  (let [sheet-name (.getSheetName sheet)
        result (rest (re-find #"^(.*)\.json(?:(#|@)?.*)$" sheet-name))
        json-name (first result)
        tag (last result)]
    (if json-name [json-name (case tag "@" 'single-object 'normal)] nil)))

(defn group-sheets [workbook]
  (seq (reduce (fn [acc sheet]
                 (if-let [[filename type] (name-and-type-from-sheet sheet)]
                   (update-in acc [[filename type]] (fnil conj []) sheet) acc))
               {} workbook)))

(defn parse-single-object [sheets]
  (let [sheet (first sheets)
        [_headers rows] (headers-and-rows sheet)]
    (reduce (fn [acc row]
              (let [[key-path value] (parse-column (first row) (second row))]
                (assoc-in acc key-path value)))
            {} rows)))

(defn parse-normal [sheets]
  (let [[headers rows] (headers-and-rows (first sheets))
        primary-key (convert-header (first headers))]
    (doall (for [row rows]
             (let [config (unpack-keys headers row)
                   current-key (keyword (get config primary-key))]
               (add-sheet-config primary-key current-key (rest sheets) config))))))

(defn parse-workbook [workbook]
  (binding [*evaluator* (.createFormulaEvaluator
                         (.getCreationHelper ^Workbook workbook))]
    (doall (for [[[json-name type] sheets] (group-sheets workbook)]
             [json-name (case type
                          single-object (parse-single-object sheets)
                          normal (parse-normal sheets))]))))

(defn convert [file-path]
  (parse-workbook (ce/workbook-xssf file-path)))
