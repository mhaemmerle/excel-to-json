Opinionated Excel -> JSON converter, where nested structures can be defined on different worksheets.

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

| id            | properties    | property_d    | property_e    |
| ------------- | ------------- | ------------- | ------------- |
| foo           |               | 66            | 100           |
| foo           |               | 77            | 200           |
| foo           |               | 88            | 300           |
| bar           |               | 99            | 400           |
| bar           |               | 111           | 500           |

Will end up as this:

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
                "property_d": 66,
                "property_e": 100
            },
            {
                "property_d": 77,
                "property_e": 200
            },
            {
                "property_d": 88,
                "property_e": 300
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
                "property_d": 99,
                "property_e": 400
            },
            {
                "property_d": 111,
                "property_e": 500
            }
        ]
    }
]
```
