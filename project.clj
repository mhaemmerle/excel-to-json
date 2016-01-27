(defproject excel-to-json "0.1.2-SNAPSHOT"
  :dependencies [[org.clojure/clojure "1.5.1"]
                 [org.clojure/core.async "0.1.267.0-0d7780-alpha"]
                 [org.clojure/tools.cli "0.3.3"]
                 [cheshire "5.3.1"]
                 [myguidingstar/clansi "1.3.0"]
                 [fswatch "0.2.0-SNAPSHOT"]
                 [org.flatland/ordered "1.5.2"]
                 [clj-excel "0.0.1"]
                 [seesaw "1.4.4"]
                 [org.clojure/core.match "0.2.1"]]
  :plugins [[lein-marginalia "0.7.1"]]
  :profiles {:uberjar {:aot :all}}
  :main excel-to-json.cli)
