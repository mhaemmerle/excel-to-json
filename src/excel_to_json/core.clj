(ns excel-to-json.core
  (:gen-class)
  (:require [cheshire.core :refer [generate-string]]
            [incanter.excel :refer [read-xls]]
            [clojure-watch.core :refer [start-watch]]))

;; watching/ANSI printing taken from https://github.com/ibdknox/cljs-watch/

(def pk :id)

(def ANSI-CODES
  {:reset "[0m"
   :default "[39m"
   :white "[37m"
   :black "[30m"
   :red "[31m"
   :green "[32m"
   :blue "[34m"
   :yellow "[33m"
   :magenta "[35m"
   :cyan "[36m"})

(defn ansi
  [code]
  (str \u001b (get ANSI-CODES code (:reset ANSI-CODES))))

(defn style
  [s & codes]
  (str (apply str (map ansi codes)) s (ansi :reset)))

(defn watcher-print
  [& text]
  (print (style (str "watcher :: ") :magenta))
  (apply print text)
  (flush))

(defn error-print
  [& text]
  (print (style (str "error :: ") :red))
  (apply print text)
  (flush))

(defn status-print
  [text]
  (print "    " (style text :green) "\n")
  (flush))

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

(defn convert-and-save
  [file]
  (let [file-path (.getPath file)
        output-file (str (.getParent file) "/" (get-filename file) ".json")
        document (read-xls file-path :header-keywords true :all-sheets? true)]
    (try
      (let [config (flatten (parse-document document pk))]
        (spit output-file (generate-string config {:pretty true}))
        (watcher-print "Converted" file-path "->" output-file "\n"))
      (catch Exception e
        (error-print "Conversion failed with: " e "\n")))))

(defn watch-callback
  [event filename]
  (let [file (clojure.java.io/file filename)]
    (when (is-xlsx? file)
      (watcher-print "Updating changed file...\n")
      (convert-and-save file)
      (status-print "[done]"))))

(defn start
  [source-dir]
  (watcher-print "Converting files in" source-dir "\n")
  (let [directory (clojure.java.io/file source-dir)
        xlsx-files (reduce (fn [acc f]
                             (if (and (.isFile f)
                                      (is-xlsx? f))
                               (conj acc f)
                               acc)) [] (file-seq directory))]
    (doseq [file xlsx-files]
      (convert-and-save file))
    (status-print "[done]")
    (start-watch [{:path source-dir
                   :event-types [:create :modify]
                   :bootstrap (fn [path] (watcher-print "Starting to watch" path "\n"))
                   :callback watch-callback
                   :options {:recursive false}}])))

(defn -main
  [& args]
  (if (seq args)
    (start (first args))
    (watcher-print "Source directory path is required\n")))
