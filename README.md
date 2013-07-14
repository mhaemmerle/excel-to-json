Usage:

    $ lein deps
    $ lein run path-to-dir-with-excel-files

Worksheet 1

| id            | property_1       | property_2    |
| ------------- | ---------------- | ------------- |
| foo           | 123              | abc           |
| bar           | 345              | def           |

Worksheet 2

| id            | traits        | property_a    | property_b    | property_c.qux  |
| ------------- | ------------- | ------------- | ------------- | --------------- |
| foo           | 1             | a             | 9             | 11              |
| foo           | 2             | b             | 8             | 22              |
| foo           | 3             | c             | 7             | 33              |
| bar           | 1             | d             | 6             | 44              |
| bar           | 2             | e             | 5             | 55              |

Worksheet 3

| id            | properties    | prop_a        | prop_b        |
| ------------- | ------------- | ------------- | ------------- |
| foo           |               | baz_1         | 100           |
| foo           |               | baz_2         | 200           |
| foo           |               | baz_3         | 300           |
| bar           |               | baz_4         | 400           |
| bar           |               | baz_5         | 500           |

Result

```json
[
    {
        "id": "foo",
        "property_1": 123,
        "property_2": "abc",
        "traits": {
            "1": {
                "property_a": "a",
                "property_b": 9,
                "property_c": {
                    "qux": 11
                }
            },
            "2": {
                "property_a": "b",
                "property_b": 8,
                "property_c": {
                    "qux": 22
                }
            },
            "3": {
                "property_a": "c",
                "property_b": 7,
                "property_c": {
                    "qux": 33
                }
            }
        },
        "properties": [
            {
                "prop_a": "baz_1",
                "prop_b": 100
            },
            {
                "prop_a": "baz_2",
                "prop_b": 200
            },
            {
                "prop_a": "baz_3",
                "prop_b": 300
            }
        ]
    },
    {
        "id": "bar",
        "property_1": 345,
        "property_2": "def",
        "traits": {
            "1": {
                "property_a": "d",
                "property_b": 6,
                "property_c": {
                    "qux": 44
                }
            },
            "2": {
                "property_a": "f",
                "property_b": 6,
                "property_c": {
                    "qux": 55
                }
            }
        },
        "properties": [
            {
                "prop_a": "baz_4",
                "prop_b": 400
            },
            {
                "prop_a": "baz_5",
                "prop_b": 500
            }
        ]
    }
]
```
