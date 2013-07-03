(ns excel-to-json.core
  (:gen-class)
  (:require [cheshire.core :refer [generate-string]]
            [incanter.excel :refer [read-xls]]))

(defn split-keys
  [k]
  (map keyword (clojure.string/split (name k) #"\.")))

(defn unpack-keys
  [m]
  (reduce (fn [acc [k v]]
            (if (nil? v)
              acc
              (let [safe-value (if (instance? String v)
                                 (read-string v)
                                 v)]
                (assoc-in acc (split-keys k) safe-value)))) {} m))

(defn column-names-and-rows
  [sheet]
  (let [[_ column-names] (first sheet)
        [_ rows] (first (rest sheet))]
    [column-names rows]))

(defn safe-keyword
  [kw]
  (keyword (str (if (instance? Number kw) (long kw) kw))))

(defn add-sheet-config
  [primary-key current-key sheets config]
  (for [sheet sheets]
    (let [[column-names rows] (column-names-and-rows sheet)
          secondary-key (second column-names)
          secondary-config (get (group-by primary-key rows) (name current-key))]
      ;; TODO remove either primary or current key
      ;; (println primary-key current-key secondary-key)
      (reduce (fn [acc row]
                (let [nested-key (safe-keyword (get row secondary-key))]
                  (assoc-in acc [current-key secondary-key nested-key]
                            (unpack-keys (dissoc row primary-key secondary-key)))))
              config secondary-config))))

(defn parse-document
  [document primary-key]
  (let [global-sheet (first document)
        [column-names rows] (column-names-and-rows global-sheet)]
    (for [row rows]
      (let [current-config (unpack-keys row)
            current-key (keyword (get row primary-key))]
        (add-sheet-config primary-key current-key (rest document) current-config)))))

(defn is-xlsx?
  [file]
  (re-matches #".*\.xlsx" (.getName file)))

(defn get-filename
  [file]
  (first (clojure.string/split (.getName file) #"\.")))

(defn run
  [source-dir]
  (let [directory (clojure.java.io/file source-dir)
        files (file-seq directory)
        xlsx-files (reduce (fn [acc f]
                             (if (and (.isFile f)
                                      (is-xlsx? f))
                               (conj acc f)
                               acc)) [] files)]
    (doseq [file xlsx-files]
      (let [file-path (.getPath file)
            output-file (str (.getParent file) "/" (get-filename file) ".json")
            document (read-xls file-path :header-keywords true :all-sheets? true)
            config (parse-document document :id)]
        (println "converted" file-path "->" output-file)
        (spit output-file (generate-string config {:pretty true}))))))

(defn -main
  [& args]
  (if (seq args)
    (run (first args))
    (println "path to Excel files is required")))
