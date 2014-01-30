(ns excel-to-json.core
  (:gen-class)
  (:require [cheshire.core :refer [generate-string]]
            [clojure-watch.core :refer [start-watch]]
            [clansi.core :refer [style]]
            [flatland.ordered.map :refer [ordered-map]]
            [clj-excel.core :as ce])
  (:import org.apache.poi.ss.usermodel.DataFormatter))

;; 'watching' taken from https://github.com/ibdknox/cljs-watch/

(def ^:dynamic *evaluator*)

(defn text-timestamp []
  (let [c (java.util.Calendar/getInstance)
        f (java.text.SimpleDateFormat. "HH:mm:ss")]
    (.format f (.getTime c))))

(defn watcher-print [& text]
  (apply println (style (str (text-timestamp) " :: watcher :: ") :magenta) text))

(defn error-print [& text]
  (apply println (style "error :: " :red) text))

(defn status-print [text]
  (println "    " (style text :green)))

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
  (reduce (fn [acc [header cell]]
            (assoc-in acc (split-keys header) (safe-value cell)))
          (ordered-map)
          (zipmap column-names row)))

(defn column-names-and-rows [sheet]
  (let [header-row (first sheet)
        column-names (map #(keyword (safe-value %)) header-row)]
    [column-names (rest sheet)]))

(defn ensure-ordered [m k]
  (if (nil? (k m)) (assoc m k (ordered-map)) m))

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
                            (if (nil? nested-key)
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

(defn is-xlsx? [file]
  (re-matches #"^((?!~\$).)*.xlsx$" (.getName file)))

(defn get-filename [file]
  (first (clojure.string/split (.getName file) #"\.")))

(defn open-workbook [file-path]
  (ce/workbook-xssf file-path))

(defn convert-and-save [file target-dir]
  (try
    (let [file-path (.getPath file)
          workbook (open-workbook file-path)]
      (let [parsed-configs (parse-workbook workbook)]
        (doseq [[filename config] parsed-configs]
          (let [output-file (str target-dir "/" filename ".json")
                json-string (generate-string config {:pretty true})]
            (spit output-file json-string)
            (watcher-print "Converted" file-path "->" output-file)))))
    (catch Exception e
      (error-print (str "Converting" file "failed with: " e "\n"))
      (clojure.pprint/pprint (.getStackTrace e)))))

(defn watch-callback [target-dir event filename]
  (let [file (clojure.java.io/file filename)]
    (when (is-xlsx? file)
      (watcher-print "Updating changed file...")
      (convert-and-save file target-dir)
      (status-print "[done]"))))

(defn start [source-dir target-dir]
  (watcher-print "Converting files in" source-dir "with output to" target-dir)
  (let [directory (clojure.java.io/file source-dir)
        xlsx-files (reduce (fn [acc f]
                             (if (and (.isFile f) (is-xlsx? f))
                               (conj acc f)
                               acc)) [] (.listFiles directory))]
    (doseq [file xlsx-files]
      (convert-and-save file target-dir))
    (status-print "[done]")
    (start-watch [{:path source-dir
                   :event-types [:create :modify]
                   :bootstrap (fn [path] (watcher-print "Starting to watch" path))
                   :callback (partial watch-callback target-dir)
                   :options {:recursive false}}])))

(defn -main [& args]
  (if (seq args)
    (let [source-dir (first args) target-dir (second args)]
      (start source-dir (or target-dir source-dir)))
    (watcher-print "Usage: excel-to-json SOURCEDIR [TARGETDIR]\n")))
