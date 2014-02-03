(ns excel-to-json.gui
  (:require [seesaw.core :as sc]
            [seesaw.bind :as sb]
            [seesaw.chooser :as sch]
            [seesaw.mig :as sm]
            [clojure.core.async :refer [put!]])
  (:import java.util.prefs.Preferences))

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

(defn create-border []
  (javax.swing.BorderFactory/createLineBorder java.awt.Color/BLACK))

(defn get-select-button [ch tag]
  (let [text (keyword (str "#" (name tag) "-text"))
        handler (fn [event]
                  (when-let [file (sch/choose-file :type "Select" :selection-mode :dirs-only)]
                    (sc/text! (sc/select (sc/to-root event) [text]) (.getPath file))
                    (put! ch [:path-change {:type tag :file file}])))]
    (sc/button :text "Choose" :action (sc/action :name "Open" :handler handler))))

(defn item-renderer [renderer info]
  (sc/config! renderer :text (:value info)))

(defn create-log [model]
  (let [listbox (sc/listbox :renderer item-renderer)]
    (sb/bind model (sb/property listbox :model))
    (doto (sc/scrollable listbox)
      (.setBorder ,, (create-border)))))

;; (get-preference :source-directory "")
(defn create-header [ch]
  (let [source-text (sc/text :id :source-text :editable? false)
        target-text (sc/text :id :target-text :editable? false)]
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

(defn create-convert-button [ch]
  (let [items [(sc/checkbox :text "Watch directory"
                            :selected? true
                            :listen [:action #(put! ch [:watching (sc/value %)])])
               :fill-h
               (sc/button :text "Convert"
                          :listen [:action (fn [_] (put! ch [:run]))])]]
    (sc/horizontal-panel :items items)))

(defn create-panel [ch log-model]
  (sc/border-panel :border 5 :vgap 5 :hgap 5
                   :north (create-header ch)
                   :center (create-log log-model)
                   :south (create-convert-button ch)))

(defn initialize [ch log-model]
  (let [frame (sc/frame :title "excel-to-json"
                        :width 800
                        :height 600
                        :on-close :exit)
        panel (create-panel ch log-model)]
    (.add ^javax.swing.JFrame frame panel)
    (sc/invoke-later (sc/show! frame))
    frame))
