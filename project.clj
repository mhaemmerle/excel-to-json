(defproject excel-to-json "0.1.0-SNAPSHOT"
  :dependencies [[org.clojure/clojure "1.5.1"]
                 [org.clojure/tools.cli "0.2.2"]
                 [cheshire "4.0.3"]
                 [myguidingstar/clansi "1.3.0"]
                 [clojure-watch "0.1.9"]
                 [org.flatland/ordered "1.5.1"]
                 [clj-excel "0.0.1"]]
  :plugins [[lein-marginalia "0.7.1"]]
  :jvm-opts ["-Djava.awt.headless=true"]
  :main excel-to-json.core)
