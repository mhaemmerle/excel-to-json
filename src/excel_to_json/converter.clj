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
            value))))))

(defn unpack-keys [column-names row]
  (let [cv (vec column-names)
        row-map (map (fn [cell]
                       [(get cv (.getColumnIndex cell)) cell]) row)]
    (reduce (fn [acc [header cell]]
              (if (= (.getCellType cell) Cell/CELL_TYPE_BLANK)
                acc
                (assoc-in acc (split-keys header) (safe-value cell))))
            (ordered-map) row-map)))

(defn column-names-and-rows [sheet]
  (let [header-row (first sheet)
        column-names (map #(keyword (safe-value %)) header-row)]
    [column-names (rest sheet)]))

(defn ensure-ordered [m k]
  (if (nil? (k m)) (assoc m k (ordered-map)) m))

(defn blank? [value]
  (cond
   (integer? value) false
   :else (clojure.string/blank? value)))

(defn add-sheet-config [primary-key current-key sheets config]
  (reduce (fn [acc0 sheet]
            (let [[column-names rows] (column-names-and-rows sheet)
                  secondary-key (second column-names)
                  unpacked-rows (map #(unpack-keys column-names %) rows)
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
  (let [[column-names rows] (column-names-and-rows (first sheets))
        primary-key (first column-names)]
    (doall (for [row rows]
             (let [config (unpack-keys column-names row)
                   current-key (keyword (get config primary-key))]
               (add-sheet-config primary-key current-key (rest sheets) config))))))

(defn parse-workbook [workbook]
  (binding [*evaluator* (.createFormulaEvaluator (.getCreationHelper workbook))]
    (doall (for [[name sheets] (group-sheets workbook)]
             [name (parse-sheets sheets)]))))

(defn convert [file-path]
  (parse-workbook (ce/workbook-xssf file-path)))
