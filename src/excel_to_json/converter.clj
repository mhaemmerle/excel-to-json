(ns excel-to-json.converter
  (:require [flatland.ordered.map :refer [ordered-map]]
            [clj-excel.core :as ce])
  (:import [org.apache.poi.ss.usermodel DataFormatter Cell]))

(def ^:dynamic *evaluator*)

(defn split-keys [k]
  (map keyword (clojure.string/split (name k) #"\.")))

(defn safe-keyword [k]
  (keyword (str (if (instance? Number k) (long k) k))))

(defn apply-format [cell]
  (.formatCellValue (DataFormatter.) cell *evaluator*))

(defn safe-value [cell]
  (let [value (apply-format cell)]
    (try
      (. Integer parseInt value)
      (catch Exception e
        (try
          (. Float parseFloat value)
          (catch Exception e
            (case (clojure.string/lower-case value)
              "true" true
              "false" false
              value)))))))

(defn safe-key [cell]
  (keyword (safe-value cell)))

(defn is-blank? [cell]
  (or (= (.getCellType cell) Cell/CELL_TYPE_BLANK)) (= (safe-value cell) ""))

(defn with-index [cells]
  (into {} (map (fn [c] [(.getColumnIndex c) c]) cells)))

(defn unpack-keys [header row]
  (let [indexed-header (with-index header)
        indexed-row (with-index row)]
    (reduce (fn [acc [i header]]
              (let [cell (get indexed-row i)]
                (if (or (is-blank? header) (nil? cell) (is-blank? cell))
                  acc
                  (assoc-in acc (split-keys (safe-key header)) (safe-value cell)))))
      (ordered-map) indexed-header)))

(defn non-empty-rows [rows]
  (filter
    (fn [row]
      (let [cell (first row)]
        (and
          (= (.getColumnIndex cell) 0)
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
                  secondary-key (safe-key (second headers))
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
                              (update-in acc [secondary-key] conj sub)
                              (assoc-in (ensure-ordered acc secondary-key)
                                        [secondary-key safe-nested-key] sub)))))
                      acc0 secondary-config)))
          config sheets))

(defn filename-from-sheet [sheet]
  (nth (re-find #"^(.*)\.json(#.*)?$" (.getSheetName sheet)) 1))

(defn group-sheets [workbook]
  (seq (reduce (fn [acc sheet]
                 (if-let [filename (filename-from-sheet sheet)]
                   (update-in acc [filename] (fnil conj []) sheet) acc))
               {} workbook)))

(defn parse-sheets [sheets]
  (let [[headers rows] (headers-and-rows (first sheets))
        primary-key (safe-key (first headers))]
    (doall (for [row rows]
             (let [config (unpack-keys headers row)
                   current-key (keyword (get config primary-key))]
               (add-sheet-config primary-key current-key (rest sheets) config))))))

(defn parse-workbook [workbook]
  (binding [*evaluator* (.createFormulaEvaluator (.getCreationHelper workbook))]
    (doall (for [[name sheets] (group-sheets workbook)]
             [name (parse-sheets sheets)]))))

(defn convert [file-path]
  (parse-workbook (ce/workbook-xssf file-path)))
