# Excel to JSON Converter

This program can parse Excel sheets to JSON file. 
It is originally for providing convenience of transferring user-friendly Excel data to Unreal Engine's Data Table. 
You can transfer Excel data to JSON file first, and then import the JSON to UE's Data Table.

The key idea of this program is to support singular data in each Excel cell, 
so that we don't have to put nested ((A, B), (D, E, F)) data in one cell. 


## Transfer Excel to JSON: Example

### Source Excel Table:

| ID | +Group | ##ID | ##Guarenteed | ##Sum | ++Item | ###ID | ###Guarenteed | ###Repetitive | ###Contribute |
|----|--------|------|--------------|-------|---------|--------|---------------|---------------|---------------|
| Shanghai.Huangpu.Xujiahui | | Weapons | 0 | 3.5 | | SMG.Uzi.1 | 0 | FALSE | 0.2 |
| | | | | | | SMG.FN_P90.2 | 0 | FALSE | 0.1 |
| | | | | | | Samurai.Nightmare.1 | 0 | FALSE | 0.2 |
| | | Supplies | 2 | 1 | | Ammo.0_75.1 | 1 | TRUE | 0.5 |
| | | | | | | Bandage.Rapid.2 | 0 | TRUE | 0.3 |
| London.Westminster.Soho | | Weapons | 0 | 2 | | Magic.Dragonfire.1 | 0 | FALSE | 0.4 |
| | | | | | | Arcane.Blade.2 | 0 | FALSE | 0.2 |
| | | Supplies | 1 | 1 | | Ammo.Explosive.2 | 0 | TRUE | 0.7 |
| | | | | | | Bandage.Elixir.1 | 0 | TRUE | 0.7 |

### Output JSON:

```json
[
    {
        "ID": "Shanghai.Huangpu.Xujiahui",
        "Group": [
            {
                "ID": "Weapons",
                "Guarenteed": 0,
                "Sum": 3.5,
                "Item": [
                    {
                        "ID": "SMG.Uzi.1",
                        "Guarenteed": 0,
                        "Repetitive": false,
                        "Contribute": 0.2
                    },
                    {
                        "ID": "SMG.FN_P90.2",
                        "Guarenteed": 0,
                        "Repetitive": false,
                        "Contribute": 0.1
                    },
                    {
                        "ID": "Samurai.Nightmare.1",
                        "Guarenteed": 0,
                        "Repetitive": false,
                        "Contribute": 0.2
                    }
                ]
            },
            {
                "ID": "Supplies",
                "Guarenteed": 2,
                "Sum": 1,
                "Item": [
                    {
                        "ID": "Ammo.0_75.1",
                        "Guarenteed": 1,
                        "Repetitive": true,
                        "Contribute": 0.5
                    },
                    {
                        "ID": "Bandage.Rapid.2",
                        "Guarenteed": 0,
                        "Repetitive": true,
                        "Contribute": 0.3
                    }
                ]
            }
        ]
    },
    {
        "ID": "London.Westminster.Soho",
        "Group": [
            {
                "ID": "Weapons",
                "Guarenteed": 0,
                "Sum": 2,
                "Item": [
                    {
                        "ID": "Magic.Dragonfire.1",
                        "Guarenteed": 0,
                        "Repetitive": false,
                        "Contribute": 0.4
                    },
                    {
                        "ID": "Arcane.Blade.2",
                        "Guarenteed": 0,
                        "Repetitive": false,
                        "Contribute": 0.2
                    }
                ]
            },
            {
                "ID": "Supplies",
                "Guarenteed": 1,
                "Sum": 1,
                "Item": [
                    {
                        "ID": "Ammo.Explosive.2",
                        "Guarenteed": 0,
                        "Repetitive": true,
                        "Contribute": 0.7
                    },
                    {
                        "ID": "Bandage.Elixir.1",
                        "Guarenteed": 0,
                        "Repetitive": true,
                        "Contribute": 0.7
                    }
                ]
            }
        ]
    }
]
```

Please see Samples to check out what can be transferred and what cannot:
* Excel files in Excels1, Excels2 show positive examples
* Excel files in Unconvertible show negatives

## Transfer Excel to JSON: How to Use

### To use the program:

#### Transfer multiple files

1. Have an Excel file, let's say, called "ExcelToJson.xlsx"
2. List target Excel paths and desired JSON paths. Let's call it "router"

| Excel | Json | header_row | data_row | start_col |
|-------|------|------------|-----------|------------|
| D:\PythonProjects\TestExcelToJson\Excel\Simple_Dict_List.5.xlsx | D:\PythonProjects\TestExcelToJson\New\1.json | 0 | -1 | 0 |
| D:\PythonProjects\TestExcelToJson\Excel\ChainedList.1.xlsx | D:\PythonProjects\TestExcelToJson\New\2.json | 0 | -1 | 0 |
| D:\PythonProjects\TestExcelToJson\Excel\Simple_Dict_List.2.xlsx | D:\PythonProjects\TestExcelToJson\New\3.json | 0 | -1 | 0 |

Parameters:
- header_row: row of header
- data_row: row of starting data, -1 means just after header_row
- start_col: col of starting data

See TransferList/ExcelToJson.xlsx as an example.

3. In JexcelConfig.ini, specify whether the "router" is in JexcelConfig.ini. For example:

```ini
[Paths]
ExcelManager_ToJson = TransferList/ExcelToJson.xlsx
```

4. Double click ExcelToJson.exe


Actually, you can copy ExcelToJson.exe plus JexcelConig.ini anywhere to run. 
As long as the files are specified correctly in config, the program should run. 


#### Make a single transfer 

To do this, you need to install python, and install pandas. 
Please follow the configuration of "Python: Run jexcel" in .vscode/launch.json to run python command.


### Source Excel File Format

This program can convert Excel data into nested json data: 
one row or multiple rows represent an integrity of a json dict.

To achieve that, a concept of "level" in headers is introduced. 
You can use a prefix of '#' to denote a dict of a header and a prefix of '+' of list.

The level of a column without prefix is 0，and that with prefixes is num of prefixes - 1. 


#### Use '#' to denote a dict

```
#Person  ##FirstName  ##LastName
```

will form a dict like:

```json
{
    "Person": {
        "FirstName": "Ababa",
        "LastName": "JeegooJeegoo"
    }
}
```

#### Use '+' to denote a list

There are two kinds of lists:
1. Lists with object content
2. Lists with dict content

It is free place list content in a new line out of its parent. 

##### Case 1: Object content

| +Pets |
|-------|
| Cat |
| Hamster |
| Dog |

will form:

```json
{
    "Pets": ["Cat", "Hamster", "Dog"]
}
```

##### Case 2: Dict content

| +Date | ##Year | ##Month |
|-------|---------|---------|
| | 2012 | Jan |
| | 2022 | Feb |

will form:

```json
{
    "Date": [
        {"Year": 2012, "Month": "Jan"},
        {"Year": 2022, "Month": "Feb"}
    ]
}
```

For this kind of list, it is HIGHLY suggested to have an "identity" column, either for the list itself or for its content.
A list's content without identity will be all collected as one integrity (See Samples/Unconvertible/ChainedList_WithoutPrimaryKey.1.xlsx)

### Cell Values

Currently, these data types can be parsed:
- Integer
- Float (must be like 1.0)
- Boolean (must be TRUE, True, or true)
- String


# Future Plan

- Apply more approaches of header specification
- Support modifying existing excel from json 