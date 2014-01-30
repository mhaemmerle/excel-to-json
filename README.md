[![Build Status](https://travis-ci.org/mhaemmerle/excel-to-json.png?branch=master)](https://travis-ci.org/mhaemmerle/excel-to-json)

This is a tool that converts Excel files of a certain structure to JSON files. It watches Excel files for any changes and constantly produces JSON files based on the workbooks.

# Usage

    $ lein deps
    $ lein run SOURCEDIR [TARGEDIR]

`SOURCEDIR` is a directory containing a number of Excel workbooks. `TARGETDIR` (optional) is the destination directory of the JSON files (defaults to `SOURCEDIR`). This will start the watcher which will print out progress as files are modified and scanned.

# Rules

## Sheet Rules

The Excel workbook has to follow a few rules to be parsable. Only sheets names ending with ".json" will be converted to JSON files. The name of the JSON file will be the same as the the sheet name. Thus, a workbook containing three sheets, *Example.json*, *Numbers* and *Data.json* will only produce two files, `example.json` and `data.json`.

## Data rules

Data inside a sheet will be parsed as a table and produces a list of row objects. The following workbook:

*Example.json* (Sheet 1)

| id            | property_1       | property_2    |
| ------------- | ---------------- | ------------- |
| foo           | 123              | abc           |
| bar           | 345              | def           |

Will produce `Example.json` containing:

```json
[
  {
    "id": "foo",
    "property_1": 123,
    "property_2": "abc"
  },
  {
    "id": "bar",
    "property_1": 345,
    "property_2": "def"
  }
]
```

### Sub Properties

#### Single Sheet Sub Properties

A column name with dots, will produce a sub tree like structure.

*Example.json* (Sheet 1)

| id            | prop.a           | prop.b        |
| ------------- | ---------------- | ------------- |
| foo           | 19               | 21            |

This sheet will generate `Example.json` containing:

```json
[
  {
    "id": "foo",
    "prop": {
      "a": 19,
      "b": 21
    }
  }
]
```

#### Keyed Sub Properties

Sub properties can also be generated with subsequent sheets. A naming scheme can be employed to split up the generation of one file into several sheets: sheets ending with a hash (`#`) and an identifier will be grouped into the same JSON file. For example, *Example.json*, *Example.json#traits* and *Example.json#properties* will all be merged into a single file, `Example.json`. The value of the identifier is ignored.

When using multiple sheets to generate one file, the subsequent sheets will be assumed to carry sub properties in their second columns.

*Example.json* (Sheet 1)

| id            |
| ------------- |
| foo           |
| bar           |

*Example.json#traits* (Sheet 2)

| id            | traits        | a             | b             |
| ------------- | ------------- | ------------- | ------------- |
| foo           | first         | 1             | 9             |
| foo           | second        | 2             | 8             |
| foo           | third         | 3             | 7             |
| bar           | first         | 4             | 6             |
| bar           | second        | 5             | 5             |

These will together generate `Example.json` containing:

```json
[
  {
    "id": "foo",
    "traits": {
      "first": {
        "a": 1,
        "b": 9
      },
      "second": {
        "a": 2,
        "b": 8
      },
      "third": {
        "a": 3,
        "b": 7
      }
    }
  },
  {
    "id": "bar",
    "traits": {
      "first": {
        "a": 4,
        "b": 6
      },
      "second": {
        "a": 5,
        "b": 5
      }
    }
  }
]
```

The column *traits* will be indexed by the values in that column, and the following colunmns will be group under the resptive key in *traits*.

#### List Sub Properties

*Example.json* (Sheet 1)

| id            |
| ------------- |
| foo           |
| bar           |

*Example.json#properties* (Sheet 2)

| id            | properties    | prop_a        | prop_b        |
| ------------- | ------------- | ------------- | ------------- |
| foo           |               | baz_1         | 100           |
| foo           |               | baz_2         | 200           |
| foo           |               | baz_3         | 300           |
| bar           |               | baz_4         | 400           |
| bar           |               | baz_5         | 500           |

```json
[
  {
    "id": "foo",
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

# Example

*Example.json* (Sheet 1)

| id            | property_1       | property_2    |
| ------------- | ---------------- | ------------- |
| foo           | 123              | abc           |
| bar           | 345              | def           |

*Example.json#traits* (Sheet 2)

| id            | traits        | property_a    | property_b    | property_c.qux  |
| ------------- | ------------- | ------------- | ------------- | --------------- |
| foo           | first         | a             | 9             | 11              |
| foo           | second        | b             | 8             | 22              |
| foo           | third         | c             | 7             | 33              |
| bar           | first         | d             | 6             | 44              |
| bar           | second        | e             | 5             | 55              |

*Example.json#properties* (Sheet 3)

| id            | properties    | prop_a        | prop_b        |
| ------------- | ------------- | ------------- | ------------- |
| foo           |               | baz_1         | 100           |
| foo           |               | baz_2         | 200           |
| foo           |               | baz_3         | 300           |
| bar           |               | baz_4         | 400           |
| bar           |               | baz_5         | 500           |

Results in the following content for `Example.json`:

```json
[
  {
    "id": "foo",
    "property_1": 123,
    "property_2": "abc",
    "traits": {
      "third": {
        "property_c": {
          "qux": 33
        },
        "property_b": 7,
        "property_a": "c"
      },
      "second": {
        "property_c": {
          "qux": 22
        },
        "property_b": 8,
        "property_a": "b"
      },
      "first": {
        "property_c": {
          "qux": 11
        },
        "property_b": 9,
        "property_a": "a"
      }
    },
    "properties": [
      {
        "prop_b": 100,
        "prop_a": "baz_1"
      },
      {
        "prop_b": 200,
        "prop_a": "baz_2"
      },
      {
        "prop_b": 300,
        "prop_a": "baz_3"
      }
    ]
  },
  {
    "id": "bar",
    "property_1": 345,
    "property_2": "def",
    "traits": {
      "second": {
        "property_c": {
          "qux": 55
        },
        "property_b": 6,
        "property_a": "f"
      },
      "first": {
        "property_c": {
          "qux": 44
        },
        "property_b": 6,
        "property_a": "d"
      }
    },
    "properties": [
      {
        "prop_b": 400,
        "prop_a": "baz_4"
      },
      {
        "prop_b": 500,
        "prop_a": "baz_5"
      }
    ]
  }
]
```
