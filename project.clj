(defproject excel-to-json "0.1.0-SNAPSHOT"
  :dependencies [[org.clojure/clojure "1.5.1"]
                 [org.clojure/tools.cli "0.2.2"]
                 [cheshire "4.0.3"]
                 [incanter "1.5.1"]
                 [clojure-watch "0.1.9"]]
  :jvm-opts ["-Djava.awt.headless=true"]
  :main excel-to-json.core)
