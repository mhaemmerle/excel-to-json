Opinionated Excel -> JSON converter, where nested structures can be defined on different worksheets.

Worksheet 1

| id            | property_1       | property_2    |
| ------------- | ---------------- | ------------- |
| foo           | 123              | abc           |
| bar           | 345              | def           |

Worksheet 2

| id            | traits        | property_a    | property_b    |
| ------------- | ------------- | ------------- | ------------- |
| foo           | 1             | a             | 9             |
| foo           | 2             | b             | 8             |
| foo           | 3             | c             | 7             |
| bar           | 1             | d             | 6             |
| bar           | 2             | e             | 5             |

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
                "property_b": 9
            },
            "2": {
                "property_a": "b",
                "property_b": 8
            },
            "3": {
                "property_a": "c",
                "property_b": 7
            }
        }
    },
    {
        "id": "bar",
        "property_1": 345,
        "property_2": "def",
        "traits": {
            "1": {
                "property_a": "d",
                "property_b": 6
            },
            "2": {
                "property_a": "f",
                "property_b": 6
            }
        }
    }
]
```
