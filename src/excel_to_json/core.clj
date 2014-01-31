(ns excel-to-json.core
  (:gen-class)
  (:require [cheshire.core :refer [generate-string]]
            [clojure-watch.core :refer [start-watch]]
            [clansi.core :refer [style]]
            [clojure.tools.cli :as cli]
            [excel-to-json.converter :as converter])
  (:import java.io.File))

(set! *warn-on-reflection* true)

(defn text-timestamp []
  (let [c (java.util.Calendar/getInstance)
        f (java.text.SimpleDateFormat. "HH:mm:ss")]
    (.format f (.getTime c))))

;; 'watching' taken from https://github.com/ibdknox/cljs-watch/

(defn watcher-print [& text]
  (apply println (style (str (text-timestamp) " :: watcher :: ") :magenta) text))

(defn error-print [& text]
  (apply println (style "error :: " :red) text))

(defn status-print [text]
  (println "    " (style text :green)))

(defn is-xlsx? [^File file]
  (re-matches #"^((?!~\$).)*.xlsx$" (.getName file)))

(defn get-filename [^File file]
  (first (clojure.string/split (.getName file) #"\.")))

(defn convert-and-save [^File file target-dir]
  (try
    (let [file-path (.getPath file)]
      (doseq [[filename config] (converter/convert file-path)]
        (let [output-file (str target-dir "/" filename ".json")
              json-string (generate-string config {:pretty true})]
          (spit output-file json-string)
          (watcher-print "Converted" file-path "->" output-file))))
    (catch Exception e
      (error-print (str "Converting" file "failed with: " e "\n"))
      (clojure.pprint/pprint (.getStackTrace e)))))

(defn watch-callback [target-dir event filename]
  (let [file (clojure.java.io/file filename)]
    (when (is-xlsx? file)
      (watcher-print "Updating changed file...")
      (convert-and-save file target-dir)
      (status-print "[done]"))))

(defn start [source-dir target-dir & {:keys [disable-watching]}]
  (watcher-print "Converting files in" source-dir "with output to" target-dir)
  (let [directory (clojure.java.io/file source-dir)
        xlsx-files (reduce (fn [acc ^File f]
                             (if (and (.isFile f) (is-xlsx? f))
                               (conj acc f)
                               acc)) [] (.listFiles directory))]
    (doseq [file xlsx-files]
      (convert-and-save file target-dir))
    (status-print "[done]")
    (when (not disable-watching)
      (start-watch [{:path source-dir
                     :event-types [:create :modify]
                     :bootstrap (fn [path] (watcher-print "Starting to watch" path))
                     :callback (partial watch-callback target-dir)
                     :options {:recursive false}}]))))

(def option-specs
  [[nil "--disable-watching" "Disable watching" :default false :flag true]
   ["-h" "--help" "Show help" :default false :flag true]])

(defn -main [& args]
  (let [p (cli/parse-opts args option-specs)]
    (when (:help (:options p))
      (println (:summary p))
      (System/exit 0))
    (let [arguments (:arguments p)]
      (if (> (count arguments) 1)
        (let [source-dir (first arguments) target-dir (second arguments)
              disable-watching (:disable-watching (:options p))]
          (start source-dir (or target-dir source-dir)
                 :disable-watching disable-watching))
        (println "Usage: excel-to-json SOURCEDIR [TARGETDIR]")))))
