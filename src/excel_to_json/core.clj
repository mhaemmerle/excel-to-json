(ns excel-to-json.core
  (:gen-class)
  (:require [cheshire.core :refer [generate-string]]
            [incanter.excel :refer [read-xls]]
            [clojure-watch.core :refer [start-watch]]
            [clansi.core :refer [style]]))

;; 'watching' taken from https://github.com/ibdknox/cljs-watch/

(def ^:dynamic *primary-key* :id)

(defn text-timestamp
  []
  (let [c (java.util.Calendar/getInstance)
        f (java.text.SimpleDateFormat. "HH:mm:ss")]
    (.format f (.getTime c))))

(defn watcher-print
  [& text]
  (apply println (style (str (text-timestamp) " :: watcher :: ") :magenta) text))

(defn error-print
  [& text]
  (apply println (style "error :: " :red) text))

(defn status-print
  [text]
  (println "    " (style text :green)))

(defn split-keys
  [k]
  (map keyword (clojure.string/split (name k) #"\.")))

(defn safe-keyword
  [k]
  (keyword (str (if (instance? Number k) (long k) k))))

(defn safe-value
  [v]
  (let [r (if (instance? String v)
            (read-string v)
            v)]
    (if (and (number? r) (= 0.0 (mod r 1)))
      (long r)
      r)))

(defn unpack-keys
  [m]
  (reduce (fn [acc [k v]]
            (if (or (nil? v) (and (coll? v) (empty? v)))
              acc
              (assoc-in acc (split-keys k) (safe-value v)))) {} m))

(defn column-names-and-rows
  [sheet]
  (let [[_ column-names] (first sheet)
        [_ rows] (first (rest sheet))]
    [column-names rows]))

(defn add-sheet-config
  [primary-key current-key sheets config]
  (reduce (fn [acc0 sheet]
            (let [[column-names rows] (column-names-and-rows sheet)
                  secondary-key (second column-names)
                  secondary-config (get (group-by primary-key rows) (name current-key))]
                  ;; TODO remove either primary or current key
                  ;; (println primary-key current-key secondary-key)
              (reduce (fn [acc row]
                        (let [nested-key (get row secondary-key)
                              sub (unpack-keys (dissoc row primary-key secondary-key))]
                          (if (empty? sub)
                            acc
                            (if (nil? nested-key)
                              (update-in acc [secondary-key] conj sub)
                              (assoc-in acc [secondary-key
                                             (safe-keyword nested-key)] sub)))))
                      acc0 secondary-config))) config sheets))

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
  (re-matches #"^((?!~\$).)*.xlsx$" (.getName file)))

(defn get-filename
  [file]
  (first (clojure.string/split (.getName file) #"\.")))

(defn convert-and-save
  [file]
  (let [file-path (.getPath file)
        output-file (str (.getParent file) "/" (get-filename file) ".json")
        document (read-xls file-path :header-keywords true :all-sheets? true)]
    (try
      (let [config (flatten (parse-document document *primary-key*))]
        (spit output-file (generate-string config {:pretty true}))
        (watcher-print "Converted" file-path "->" output-file))
      (catch Exception e
        (error-print "Conversion failed with: " e "\n")))))

(defn watch-callback
  [event filename]
  (let [file (clojure.java.io/file filename)]
    (when (is-xlsx? file)
      (watcher-print "Updating changed file...")
      (convert-and-save file)
      (status-print "[done]"))))

(defn start
  [source-dir]
  (watcher-print "Converting files in" source-dir)
  (let [directory (clojure.java.io/file source-dir)
        xlsx-files (reduce (fn [acc f]
                             (if (and (.isFile f) (is-xlsx? f))
                               (conj acc f)
                               acc)) [] (file-seq directory))]
    (doseq [file xlsx-files]
      (convert-and-save file))
    (status-print "[done]")
    (start-watch [{:path source-dir
                   :event-types [:create :modify]
                   :bootstrap (fn [path] (watcher-print "Starting to watch" path))
                   :callback watch-callback
                   :options {:recursive false}}])))

(defn -main
  [& args]
  (if (seq args)
    (start (first args))
    (watcher-print "Source directory path is required\n")))
