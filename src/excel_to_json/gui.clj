(ns excel-to-json.gui
  (:gen-class)
  (:require [clojure.core.async :refer [go chan <! >! put!]]
            [seesaw.core :as sc]
            [seesaw.bind :as sb]
            [seesaw.chooser :as sch]
            [seesaw.mig :as sm]
            [excel-to-json.core :as c])
  (:import java.util.prefs.Preferences
           [excel_to_json.logger StoreLogger]))

;; TODO button for applying source -> target

(sc/native!)

(defn preferences-node [path-name]
  (.node (Preferences/userRoot) path-name))

(def ^:dynamic *preferences* (preferences-node "excel-to-json"))

(defn get-preference
  ([key]
     (get-preference key nil))
  ([key default-value]
     (.get *preferences* (name key) default-value)))

(defn put-preference [key value]
  (.put *preferences* key value))

(defn create-border []
  (javax.swing.BorderFactory/createLineBorder java.awt.Color/BLACK))

(defn get-select-button [ch tag]
  (let [text (keyword (str "#" (name tag) "-text"))
        handler (fn [event]
                  (when-let [file (sch/choose-file :type "Select"
                                                   :selection-mode :dirs-only)]
                    (sc/text! (sc/select (sc/to-root event) [text]) (.getPath file))
                    (put-preference (str (name tag) "-directory") (.getPath file))
                    (put! ch [:path-change {:type tag :file file}])))]
    (sc/button :text "Choose" :action (sc/action :name "Open" :handler handler))))

(defn create-header [ch source-path target-path]
  (let [source-text (sc/text :id :source-text :text source-path)
        target-text (sc/text :id :target-text :text target-path)]
    (sm/mig-panel
     :constraints ["wrap 3, insets 0"
                   "[shrink 0]10[200, grow, fill]10[shrink 0]"
                   "[shrink 0]5[]"]
     :items [["Source directory:"]
             [source-text]
             [(get-select-button ch :source)]
             ["Target directory:"]
             [target-text]
             [(get-select-button ch :target)]])))

(defn item-renderer [renderer info]
  (sc/config! renderer :text (:value info)))

(defn create-log [model]
  (let [listbox (sc/listbox :renderer item-renderer)]
    (sb/bind model (sb/property listbox :model))
    (doto (sc/scrollable listbox)
      (.setBorder ,, (create-border)))))

(defn create-convert-button [ch]
  (let [items [(sc/checkbox :text "Watch directory"
                            :selected? true
                            :listen [:action #(put! ch [:watching (sc/value %)])])
               :fill-h
               (sc/button :text "Convert"
                          :listen [:action (fn [_] (put! ch [:run]))])]]
    (sc/horizontal-panel :items items)))

(defn create-panel [channel source-path target-path log-model]
  (sc/border-panel :border 5 :vgap 5 :hgap 5
                   :north (create-header channel source-path target-path)
                   :center (create-log log-model)
                   :south (create-convert-button channel)))

(defn initialize [channel log-model]
  (let [source-path (get-preference :source-directory)
        target-path (get-preference :target-directory)
        frame (sc/frame :title "Excel > JSON"
                        :width 1350
                        :height 650
                        :on-close :exit)
        panel (create-panel channel (or source-path "") (or target-path "") log-model)]
    (.add ^javax.swing.JFrame frame panel)
    (sc/invoke-later (sc/show! frame))
    [frame source-path target-path]))

(defn -main [& args]
  (let [channel (chan)
        log (atom [])
        [_ source-path target-path] (initialize channel log)
        m {:source-path source-path :target-path target-path :watched-path source-path}]
    (binding [c/*logger* (StoreLogger. log)]
      (let [initial-state (c/switch-watching! m true)]
        (go
         (loop [event (<! channel)
                state initial-state]
           (let [new-state (c/handle-event state event)]
             (recur (<! channel) new-state))))))))
