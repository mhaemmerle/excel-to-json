(defproject excel-to-json "0.1.0-SNAPSHOT"
  :dependencies [[org.clojure/clojure "1.5.1"]
                 [org.clojure/core.async "0.1.267.0-0d7780-alpha"]
                 [org.clojure/tools.cli "0.3.1"]
                 [cheshire "5.3.1"]
                 [myguidingstar/clansi "1.3.0"]
                 [clojure-watch "0.1.9"]
                 [org.flatland/ordered "1.5.2"]
                 [clj-excel "0.0.1"]
                 [seesaw "1.4.4"]]
  :plugins [[lein-marginalia "0.7.1"]]
  ;; :jvm-opts ["-Djava.awt.headless=true"]
  :main excel-to-json.core)
