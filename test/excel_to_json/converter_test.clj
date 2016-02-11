(ns excel-to-json.converter-test
  (:use clojure.test
        excel-to-json.converter))

(deftest convert-all
  (let [configs (convert "resources/test.xlsx")
        expected [["bar" {:id 1 :foo {:baz "x" :bar "y"} :array [1,2,3]}]
                  ["foo"
                   [{:name "baz" :value "sun" :prop_a "baz_1"}
                    {:name "qux" :value "moon" :prop_a "baz_3"}]]
                  ["test"
                   [{:id "foo"
                     :property_1 600
                     :property_2 true
                     :traits
                     {:1
                      {:property_a "a"
                       :property_b 9.00M
                       :property_c {:baz "70%" :qux "1100%"}}
                      :2
                      {:property_a "b"
                       :property_b 8.00M
                       :property_c {:baz "80%" :qux "2200%"}}
                      :3
                      {:property_a "c"
                       :property_b 7.00M
                       :property_c {:baz "90%" :qux "3300%"}}}
                     :properties
                     [{:prop_a "baz_1" :prop_b 100}
                      {:prop_a "baz_2" :prop_b 200}
                      {:prop_a "baz_3" :prop_b 300}]}
                    {:id "bar"
                     :property_1 900
                     :property_2 false
                     :traits
                     {:1
                      {:property_a "d"
                       :property_b 6.00M
                       :property_c {:baz "75%" :qux "4400%"}}
                      :2
                      {:property_a "e"
                       :property_b 5.00M
                       :property_c {:baz "85%" :qux "5500%"}}}
                     :properties
                     [{:prop_a "baz_4" :prop_b 400}
                      {:prop_a "baz_5" :prop_b 500}]}]]]]
    (is (= expected configs))))

(deftest convert-array
  (let [configs (convert "resources/test_array.xlsx")
        expected [["test_array"
                   [{:properties [{:prop_b 100 :prop_a "baz_1"}]
                     :id "foo"}]]]]
    (is (= expected configs))))

(deftest convert-map
  (let [configs (convert "resources/test_map.xlsx")
        expected [["test_map"
                   [{:traits {:first {:a 1 :b 9} :second {:a 2 :b 8}}
                     :id "foo"}]]]]
    (is (= expected configs))))

(deftest ignore-formatting
  (let [configs (convert "resources/test_formatted_cells.xlsx")
        expected [["test_formatted_cells"
                   [{:id "foo" :map {:a {:value 1}}}
                    {:id "bar" }]]]]
    (is (= expected configs))))

(deftest empty-rows
  (let [configs (convert "resources/test_empty_rows.xlsx")
        expected [["test_empty_rows"
                   [{:value 1 :id "foo"} {:value 2 :id "bar"}]]]]
    (is (= expected configs))))

(deftest empty-columns
  (let [configs (convert "resources/test_empty_columns.xlsx")
        expected [["test_empty_columns"
                   [{:id "foo" :list [{:x 1}]}
                    {:more "b" :id "bar" :list [{:x 2}]}]]]]
    ()
    (is (= expected configs))))

(deftest convert-array-annotation
  (let [configs (convert "resources/test_array_annotation.xlsx")
        expected [["test_array"
                   [{:properties [{:prop ["foo," true 2 "bar," "baz"]}]
                     :value [1 2 3]
                     :id "foo"}]]]]
    (is (= expected configs))))

(deftest sparse-sheet
  (let [configs (convert "resources/test_sparse_sheet.xlsx")
        expected [["test_sparse"
                   {:id 1 :foo {:baz "x" :bar "y"} :array [1,2,3]}]]]
    (is (= expected configs))))
