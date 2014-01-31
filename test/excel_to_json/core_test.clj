(ns excel-to-json.core-test
  (:use clojure.test
        excel-to-json.core)
  (:require [cheshire.core :refer [generate-string parse-string]]
            [excel-to-json.converter :as e2j]))

(deftest parse-excel
  (let [file-path "resources/test.xlsx"
        data (e2j/convert file-path)
        expected [["foo" [{:prop_a "baz_1" :value "sun" :name "baz"}
                          {:prop_a "baz_3" :value "moon" :name "qux"}]]
                  ["test" [{:property_2 "abc"
                            :property_1 600
                            :id "foo"
                            :traits
                            {:1
                             {:property_c {:qux "1100%" :baz "70%"}
                              :property_b 9.0
                              :property_a "a"}
                             :2
                             {:property_c {:qux "2200%" :baz "80%"}
                              :property_b 8.0
                              :property_a "b"}
                             :3
                             {:property_c {:qux "3300%" :baz "90%"}
                              :property_b 7.0
                              :property_a "c"}}
                            :properties
                            {:baz_1 {:prop_a 100}
                             :baz_2 {:prop_a 200}
                             :baz_3 {:prop_a 300}}}
                           {:property_2 "def"
                            :property_1 900
                            :id "bar"
                            :traits
                            {:1
                             {:property_c {:qux "4400%" :baz "75%"}
                              :property_b 6.0
                              :property_a "d"}
                             :2
                             {:property_c {:qux "5500%" :baz "85%"}
                              :property_b 5.0
                              :property_a "e"}}
                            :properties
                            {:baz_4 {:prop_a 400}
                             :baz_5 {:prop_a 500}}}]]]]
    (is (= expected (parse-string (generate-string data) true)))))
