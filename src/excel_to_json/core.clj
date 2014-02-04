(ns excel-to-json.core
  (:gen-class)
  (:require [cheshire.core :refer [generate-string]]
            [clojure-watch.core :refer [start-watch]]
            [clansi.core :refer [style]]
            [clojure.tools.cli :as cli]
            [excel-to-json.converter :as converter]
            [excel-to-json.gui :as gui]
            [clojure.core.async :refer [go chan <! >! put!]])
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

;; add-watcher
;; cancel-watcher

(defn watch-callback [target-dir event filename]
  (let [file (clojure.java.io/file filename)]
    (when (is-xlsx? file)
      (watcher-print "Updating changed file...")
      (convert-and-save file target-dir)
      (status-print "[done]"))))

;; (defn start [source-dir target-dir & {:keys [disable-watching]}]
(defn start [source-dir target-dir]
  ;; (watcher-print "Converting files in" source-dir "with output to" target-dir)
  (let [directory (clojure.java.io/file source-dir)
        xlsx-files (reduce (fn [acc ^File f]
                             (if (and (.isFile f) (is-xlsx? f))
                               (conj acc f)
                               acc)) [] (.listFiles directory))]
    (doseq [file xlsx-files]
      (convert-and-save file target-dir))
    ;; (status-print "[done]")
    ;; (when (not disable-watching)
    ;;   (start-watch [{:path source-dir
    ;;                  :event-types [:create :modify]
    ;;                  :bootstrap (fn [path] (watcher-print "Starting to watch" path))
    ;;                  :callback (partial watch-callback target-dir)
    ;;                  :options {:recursive false}}]))
    ))

(defn aw [source-dir target-dir]
  (println source-dir target-dir)
  (start-watch [{:path source-dir
                 :event-types [:create :modify]
                 :bootstrap (fn [path] (watcher-print "Starting to watch" path))
                 :callback (partial watch-callback target-dir)
                 :options {:recursive false}}]))

(defn rw [watcher]
  )

(def option-specs
  [[nil "--disable-watching" "Disable watching" :default false :flag true]
   ["-h" "--help" "Show help" :default false :flag true]])

;; be able to change watch target
;; re-run on directory-change

(defn switch-watching [state enabled?]
  (if enabled?
    (if (and (not (:watcher state))
             (every? #(not (nil? %)) (map state [:source-dir :target-dir])))
      (do
        (println "#1")
        (let [w (aw (:source-dir state) (:target-dir state))]
          (println "aw" aw)
          (assoc state :watcher w)))
      state)
    (if-let [watcher (:watcher state)]
      (do
        (println "#2")
        (rw watcher)
        (dissoc state :watcher))
      state)))

;; use a multimethod
(defn handle-event [state t payload]
  (case t
    :path-change (case (:type payload)
                   :source (assoc state :source-dir (.getPath (:file payload)))
                   :target (assoc state :target-dir (.getPath (:file payload))))
    ;; :run state
    :watching (switch-watching state payload)
    (do
      (println "unknown type:" t "with payload:" payload)
      state)))

(defn -main [& args]
  (let [channel (chan)
        log (atom [])
        [_ source-dir target-dir] (gui/initialize channel log)
        initial-state {:source-dir source-dir :target-dir target-dir}]
    (go
     (loop [[type payload] (<! channel)
            state initial-state]
       (println "state" state "type" type "payload" payload)
       (let [new-state (handle-event state type payload)]
         (println "new-state" new-state)
         (recur (<! channel)
                new-state)))))

  ;; (let [p (cli/parse-opts args option-specs)]
  ;;   (when (:help (:options p))
  ;;     (println (:summary p))
  ;;     (System/exit 0))
  ;;   (let [arguments (:arguments p)]
  ;;     (if (> (count arguments) 1)
  ;;       (let [source-dir (first arguments) target-dir (second arguments)
  ;;             disable-watching (:disable-watching (:options p))]
  ;;         (start source-dir (or target-dir source-dir)
  ;;                :disable-watching disable-watching))
  ;;       (println "Usage: excel-to-json SOURCEDIR [TARGETDIR]"))))

  )
