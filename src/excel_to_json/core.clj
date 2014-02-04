(ns excel-to-json.core
  (:gen-class)
  (:require [cheshire.core :refer [generate-string]]
            [fswatch.core :as fs]
            [clansi.core :refer [style]]
            [clojure.tools.cli :as cli]
            [excel-to-json.converter :as converter]
            [excel-to-json.gui :as gui]
            [clojure.core.async :refer [go chan <! >! put!]])
  (:import java.io.File
           sun.nio.fs.UnixPath))

(set! *warn-on-reflection* true)

(def ^:dynamic *prn-fn* println)

(defn text-timestamp []
  (let [calendar (java.util.Calendar/getInstance)
        date-format (java.text.SimpleDateFormat. "HH:mm:ss")]
    (.format date-format (.getTime calendar))))

;; 'watching' taken from https://github.com/ibdknox/cljs-watch/

(defn watcher-print [& text]
  (apply *prn-fn* (style (str (text-timestamp) " :: watcher :: ") :magenta) text))

(defn error-print [& text]
  (apply *prn-fn* (style "error :: " :red) text))

(defn status-print [text]
  (*prn-fn* "    " (style text :green)))

(defn is-xlsx? [^File file]
  (re-matches #"^((?!~\$).)*.xlsx$" (.getName file)))

(defn get-filename [^File file]
  (first (clojure.string/split (.getName file) #"\.")))

(defn convert-and-save [^File file target-path]
  (try
    (let [file-path (.getPath file)]
      (doseq [[filename config] (converter/convert file-path)]
        (let [output-file (str target-path "/" filename ".json")
              json-string (generate-string config {:pretty true})]
          (spit output-file json-string)
          (watcher-print "Converted" file-path "->" output-file))))
    (catch Exception e
      (error-print (str "Converting" file "failed with: " e "\n"))
      (clojure.pprint/pprint (.getStackTrace e)))))

(defn watch-callback [source-path target-path file-path]
  (let [file (clojure.java.io/file source-path (.toString ^UnixPath file-path))]
    (when (is-xlsx? file)
      (watcher-print "Updating changed file...")
      (convert-and-save file target-path)
      (status-print "[done]"))))

(defn run [{:keys [source-path target-path] :as state}]
  (watcher-print "Converting files in" source-path "with output to" target-path)
  (let [directory (clojure.java.io/file source-path)
        xlsx-files (reduce (fn [acc ^File f]
                             (if (and (.isFile f) (is-xlsx? f))
                               (conj acc f)
                               acc)) [] (.listFiles directory))]
    (doseq [file xlsx-files]
      (convert-and-save file target-path))
    (status-print "[done]")
    state))

(defn stop-watching [state]
  (if-let [path (:watched-path state)]
    (do
      (fs/unwatch-path path)
      (dissoc state :watched-path))
    state))

(defn start-watching [{:keys [source-path target-path watched-path] :as state}]
  (let [callback #(watch-callback source-path target-path %)
        new-state (if (not (= watched-path source-path))
                    (stop-watching state)
                    state)]
    (fs/watch-path source-path :create callback :modify callback)
    (watcher-print "Starting to watch" source-path)
    (assoc new-state :watched-path source-path)))

(def option-specs
  [[nil "--disable-watching" "Disable watching" :default false :flag true]
   ["-h" "--help" "Show help" :default false :flag true]])

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
  (*prn-fn* (format "Unknown event-type '%s'" event-type))
  state)

(defn -main [& args]
  (let [channel (chan)
        log (atom [])
        [_ source-path target-path] (gui/initialize channel log)
        m {:source-path source-path :target-path target-path :watched-path source-path}]
    (binding [*prn-fn* (fn [& args] (swap! log conj (apply str args)))]
      (let [initial-state (switch-watching! m true)]
        (go
         (loop [event (<! channel)
                state initial-state]
           (let [new-state (handle-event state event)]
             (recur (<! channel) new-state))))))))

;; (defn -main [& args]
;;   (let [p (cli/parse-opts args option-specs)]
;;     (when (:help (:options p))
;;       (println (:summary p))
;;       (System/exit 0))
;;     (let [arguments (:arguments p)]
;;       (if (> (count arguments) 1)
;;         (let [source-path (first arguments) target-path (second arguments)
;;               state {:source-path source-path :target-path
;;                      (or target-path source-path)
;;                      :watched-path source-path}]
;;           (run state)
;;           (when-not (:disable-watching (:options p))
;;             (start-watching state)))
;;         (println "Usage: excel-to-json SOURCEDIR [TARGETDIR]")))))
