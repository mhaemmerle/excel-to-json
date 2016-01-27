(ns excel-to-json.cli
  (:gen-class)
  (:require [clojure.string :as string]
            [excel-to-json.core :as core]
            [clojure.tools.cli :as tools.cli])
  (:import java.io.File))

(def cli-options
  [[nil "--disable-watching" "Disable watching"
    :default false]
   ["-w" "--wrapper WRAPPER" "Wrap list in object"
    :default nil]
   ["-e" "--ext EXT" "Use ext instead of json"
    :default "json"]
   ["-h" "--help"]])
  ; ;; An option with a required argument
  ; [["-p" "--port PORT" "Port number"
  ;   :default 80
  ;   :parse-fn #(Integer/parseInt %)
  ;   :validate [#(< 0 % 0x10000) "Must be a number between 0 and 65536"]]
  ;  ;; A non-idempotent option
  ;  ["-v" nil "Verbosity level"
  ;   :id :verbosity
  ;   :default 0
  ;   :assoc-fn (fn [m k _] (update-in m [k] inc))]
  ;  ;; A boolean option defaulting to nil
  ;  ["-h" "--help"]])

(defn usage [options-summary]
  (->> ["Opinionated Excel to JSON converter."
        ""
        "Usage: excel-to-json [SOURCE] TARGET"
        ""
        "Options:"
        options-summary
        ""
        "Please refer to https://github.com/mhaemmerle/excel-to-json for more"
        "information."]
       (string/join \newline)))

(defn error-msg [errors]
  (str (string/join \newline errors)
       "\n\nUse -h or --help to show usage"))

(defn exit [status msg]
  (println msg)
  (System/exit status))

(defn get-absolute-path [^java.io.File f]
  (let [absolute-path (.getAbsolutePath f)]
    (.substring absolute-path 0 (.lastIndexOf absolute-path File/separator))))

(defn parse [args]
  (tools.cli/parse-opts args cli-options))

(defn -main [& args]
  (let [{:keys [options arguments errors summary]}
        (parse args)]
    ;; Handle help and error conditions
    (cond
      (:help options) (exit 0 (usage summary))
      (< (count arguments) 1) (exit 1 (usage summary))
      errors (exit 1 (error-msg errors)))
    ;; Execute program with options
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
                 :ext (:ext options)
                 :wrapper (:wrapper options)}]
      (core/run state)
      (when-not (:disable-watching options)
        (core/start-watching state)
        nil))))
