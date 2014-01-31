(ns excel-to-json.converter-test
  (:use clojure.test
        excel-to-json.converter))

(deftest convert-all
  (let [configs (convert "resources/test.xlsx")
        expected [["foo"
                   [{:name "baz" :value "sun" :prop_a "baz_1"}
                    {:name "qux" :value "moon" :prop_a "baz_3"}]]
                  ["test"
                   [{:id "foo"
                     :property_1 600
                     :property_2 "abc"
                     :traits
                     {:1
                      {:property_a "a"
                       :property_b 9.0
                       :property_c {:baz "70%" :qux "1100%"}}
                      :2
                      {:property_a "b"
                       :property_b 8.0
                       :property_c {:baz "80%" :qux "2200%"}}
                      :3
                      {:property_a "c"
                       :property_b 7.0
                       :property_c {:baz "90%" :qux "3300%"}}}
                     :properties
                     [{:prop_a "baz_3" :prop_b 300}
                      {:prop_a "baz_2" :prop_b 200}
                      {:prop_a "baz_1" :prop_b 100}]}
                    {:id "bar"
                     :property_1 900
                     :property_2 "def"
                     :traits
                     {:1
                      {:property_a "d"
                       :property_b 6.0
                       :property_c {:baz "75%" :qux "4400%"}}
                      :2
                      {:property_a "e"
                       :property_b 5.0
                       :property_c {:baz "85%" :qux "5500%"}}}
                     :properties
                     [{:prop_a "baz_5" :prop_b 500}
                      {:prop_a "baz_4" :prop_b 400}]}]]]]
    (is (= expected configs))))

(deftest convert-arrays
  (let [configs (convert "resources/test_array.xlsx")
        expected [["test"
                   [{:properties [{:prop_b 100 :prop_a "baz_1"}]
                     :id "foo"}]]]]
    (is (= expected configs))))

(deftest convert-arrays
  (let [configs (convert "resources/test_map.xlsx")
        expected [["test"
                   [{:traits {:first {:a 1 :b 9} :second {:a 2 :b 8}}
                     :id "foo"}]]]]
    (is (= expected configs))))
