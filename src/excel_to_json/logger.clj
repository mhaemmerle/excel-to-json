(ns excel-to-json.logger
  (:require [clansi.core :refer [style]]))

(defn text-timestamp []
  (let [calendar (java.util.Calendar/getInstance)
        date-format (java.text.SimpleDateFormat. "HH:mm:ss")]
    (.format date-format (.getTime calendar))))

(defprotocol Logger
  (info [this msg])
  (error [this msg])
  (status [this msg]))

(deftype PrintLogger []
  Logger
  (info [this text]
    (println (style (str (text-timestamp) " :: watcher :: ") :magenta) text))
  (error [this text]
    (println (style "error :: " :red) text))
  (status [this text]
    (println "    " (style text :green))))

(deftype StoreLogger [store]
  Logger
  (info [this text]
    (swap! store conj (str (text-timestamp) text)))
  (error [this text]
    (swap! store conj (str "error :: " text)))
  (status [this text]
    (swap! store conj (str "    " text))))
