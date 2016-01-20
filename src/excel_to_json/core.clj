(ns excel-to-json.core
  (:gen-class)
  (:require [clojure.core.async :refer [go chan <! >! put!]]
            [cheshire.core :refer [generate-string]]
            [fswatch.core :as fs]
            [clojure.tools.cli :as cli]
            [excel-to-json.converter :as converter]
            [excel-to-json.logger :as log])
  (:import java.io.File [excel_to_json.logger PrintLogger]
           java.nio.file.Path))

(set! *warn-on-reflection* true)

(def ^:dynamic *logger* (PrintLogger.))

;; 'watching' taken from https://github.com/ibdknox/cljs-watch/

(defn is-xlsx? [^File file]
  (re-matches #"^((?!~\$).)*.xlsx$" (.getName file)))

(defn get-filename [^File file]
  (first (clojure.string/split (.getName file) #"\.")))

(defn convert-and-save [^File file target-path ext]
  (try
    (let [file-path (.getPath file)]
      (doseq [[filename config] (converter/convert file-path)]
        (let [output-file (str target-path "/" filename "." ext)
              json-string (generate-string config {:pretty true})]
          (spit output-file json-string)
          (log/info *logger* (str "Converted " file-path "->" output-file)))))
    (catch Exception e
      (log/error *logger* (str "Converting" file "failed with: " e "\n"))
      (clojure.pprint/pprint (.getStackTrace e)))))

(defn watch-callback [source-path source-file target-path ext ^Path file-path]
  (let [f (.getPath ^File (.toFile file-path))
        file (try
               (clojure.java.io/file source-path f)
               (catch Exception e
                 (log/error *logger* e)))
        a1 (if (nil? source-file) nil
               (.getAbsolutePath ^File source-file))
        a2 (.getAbsolutePath ^File file)]
    (when (is-xlsx? file)
      (when (or (nil? source-file)
                (= a1 a2))
        (log/info *logger* "Updating changed file...")
        (convert-and-save file target-path ext)
        (log/status *logger* "[done]")))))

(defn run [{:keys [source-path source-file target-path ext] :as state}]
  (log/info *logger* (format "Converting files from '%s' to '%s'"
                             source-path target-path))
  (let [directory (clojure.java.io/file source-path)
        xlsx-files (if (nil? source-file)
                     (reduce (fn [acc ^File f]
                               (if (and (.isFile f) (is-xlsx? f))
                                 (conj acc f)
                                 acc)) [] (.listFiles directory))
                     [source-file])]
    (doseq [file xlsx-files]
      (convert-and-save file target-path ext))
    (log/status *logger* "[done]")
    state))

(defn stop-watching [state]
  (if-let [path (:watched-path state)]
    (do
      (fs/unwatch-path path)
      (dissoc state :watched-path))
    state))

(defn start-watching [{:keys [source-path source-file
                              target-path watched-path ext] :as state}]
  (let [callback #(watch-callback source-path source-file target-path ext %)
        new-state (if (not (= watched-path source-path))
                    (stop-watching state)
                    state)]
    (fs/watch-path source-path :create callback :modify callback)
    (log/info *logger* (format "Starting to watch '%s'" source-path))
    (assoc new-state :watched-path source-path)))

(def option-specs
  [[nil "--disable-watching" "Disable watching" :default false :flag true]
   ["-h" "--help" "Show help" :default false :flag true]
   ["-e" "--ext EXT" "Use ext instead on json" :default "json"]]
  )

;; re-run on directory-change

(defn switch-watching! [state enabled?]
  (if enabled?
    (if (every? #(not (nil? %)) (map state [:source-path :target-path]))
      (start-watching state)
      state)
    (stop-watching state)))

(defmulti handle-event (fn [state [event-type payload]] event-type))

(defmethod handle-event :path-change [state [event-type payload]]
  (let [path (.getPath ^File (:file payload))]
    (case (:type payload)
      :source (assoc state :source-path path)
      :target (assoc state :target-path path))))

(defmethod handle-event :run [state _]
  (run state))

(defmethod handle-event :watching [state [event-type payload]]
  (switch-watching! state payload))

(defmethod handle-event :default [state [event-type payload]]
  (log/error *logger* (format "Unknown event-type '%s'" event-type))
  state)

(defn get-absolute-path [^java.io.File f]
  (let [absolute-path (.getAbsolutePath f)]
    (.substring absolute-path 0 (.lastIndexOf absolute-path File/separator))))

(defn -main [& args]
  (let [parsed-options (cli/parse-opts args option-specs)]
    (when (:help (:options parsed-options))
      (println (:summary parsed-options))
      (System/exit 0))
    (let [arguments (:arguments parsed-options)]
      (if (> (count arguments) 0)
        (let [source-arg (first arguments) target-path-arg (second arguments)
              source-is-file (.isFile (clojure.java.io/file source-arg))
              source-file (if source-is-file
                            (clojure.java.io/file source-arg)
                            nil)
              source-path (if source-is-file
                            (get-absolute-path source-file)
                            source-arg)
              target-path (if target-path-arg target-path-arg source-path)
              state {:source-path source-path
                     :source-file source-file
                     :target-path target-path
                     :watched-path source-path
                     :ext (:ext (:options parsed-options))}]
          (run state)
          (when-not (:disable-watching (:options parsed-options))
            (start-watching state)
            nil))
        (do
          (println "Usage: excel-to-json SOURCE [TARGETDIR]")
          (println "       SOURCE can be either a directory or a single file."))))))
