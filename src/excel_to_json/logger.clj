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

(defn add-line! [store line]
  (swap! store conj line))

(deftype StoreLogger [store]
  Logger
  (info [this text]
    (add-line! store (str (text-timestamp) " " text)))
  (error [this text]
    (add-line! store (str "error :: " text)))
  (status [this text]
    (add-line! store (str "    " text))))
