# Excel 转 JSON 转换器

本程序可以将 Excel 表格转换为 JSON 文件。
它最初是为了方便将用户友好的 Excel 数据转换到虚幻引擎的数据表（Data Table）而开发的。
你可以先将 Excel 数据转换为 JSON 文件，然后将 JSON 导入到虚幻引擎的数据表中。

本程序的核心理念是支持在每个 Excel 单元格中放置单一数据，
这样我们就不需要在一个单元格中放置嵌套数据（如 ((A, B), (D, E, F))）。


## Excel 转 JSON：示例

### 源 Excel 表格：

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

### 输出 JSON：

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

请查看示例文件夹了解哪些可以转换，哪些不能转换：
* Excels1、Excels2 文件夹中的 Excel 文件展示了正面示例
* Unconvertible 文件夹中的 Excel 文件展示了负面示例

## Excel 转 JSON：使用方法

### 程序使用方法：

#### 转换多个文件

1. 准备一个 Excel 文件，比如命名为 "ExcelToJson.xlsx"
2. 列出目标 Excel 路径和期望的 JSON 路径。我们称之为"路由表"

| Excel | Json | header_row | data_row | start_col |
|-------|------|------------|-----------|------------|
| D:\PythonProjects\TestExcelToJson\Excel\Simple_Dict_List.5.xlsx | D:\PythonProjects\TestExcelToJson\New\1.json | 0 | -1 | 0 |
| D:\PythonProjects\TestExcelToJson\Excel\ChainedList.1.xlsx | D:\PythonProjects\TestExcelToJson\New\2.json | 0 | -1 | 0 |
| D:\PythonProjects\TestExcelToJson\Excel\Simple_Dict_List.2.xlsx | D:\PythonProjects\TestExcelToJson\New\3.json | 0 | -1 | 0 |

参数说明：
- header_row：表头所在行
- data_row：数据起始行，-1 表示紧接在表头行之后
- start_col：数据起始列

参考 TransferList/ExcelToJson.xlsx 作为示例。

3. 在 JexcelConfig.ini 中指定"路由表"的位置。例如：

```ini
[Paths]
ExcelManager_ToJson = ConvertList_ToJson.xlsx
```

4. 双击 ExcelToJson.exe 运行


实际上，你可以将 ExcelToJson.exe 和 JexcelConig.ini 和包含转换文件的 Excel 复制到任何位置运行。
只要配置文件中正确指定了文件路径，程序就能运行。


#### 进行单次转换

要进行单次转换，你需要安装 Python 和 pandas。
请按照 .vscode/launch.json 中 "Python: Run jexcel" 的配置来运行 Python 命令。


### 源 Excel 文件格式

本程序可以将 Excel 数据转换为嵌套的 JSON 数据：
一行或多行数据表示一个完整的 JSON 字典。

为了实现这一点，引入了表头"级别"的概念。
你可以使用 '#' 前缀来表示表头的字典，使用 '+' 前缀表示列表。

没有前缀的列级别为 0，带前缀的列级别为前缀数量减 1。


#### 使用 '#' 表示字典

```
#Person  ##FirstName  ##LastName
```

将形成如下字典：

```json
{
    "Person": {
        "FirstName": "Ababa",
        "LastName": "JeegooJeegoo"
    }
}
```

#### 使用 '+' 表示列表

有两种类型的列表：
1. 包含对象内容的列表
2. 包含字典内容的列表

可以自由地在父级之外的新行中放置列表内容。

##### 情况 1：对象内容

| +Pets |
|-------|
| Cat |
| Hamster |
| Dog |

将形成：

```json
{
    "Pets": ["Cat", "Hamster", "Dog"]
}
```

##### 情况 2：字典内容

| +Date | ##Year | ##Month |
|-------|---------|---------|
| | 2012 | Jan |
| | 2022 | Feb |

将形成：

```json
{
    "Date": [
        {"Year": 2012, "Month": "Jan"},
        {"Year": 2022, "Month": "Feb"}
    ]
}
```

对于这种类型的列表，强烈建议为列表本身或其内容设置一个"标识"列。
没有标识的列表内容将被全部收集为一个整体（参见 Samples/Unconvertible/ChainedList_WithoutPrimaryKey.1.xlsx）

### 单元格值

目前支持解析以下数据类型：
- 整数
- 浮点数（必须形如 1.0）
- 布尔值（必须是 TRUE、True 或 true）
- 字符串


# 未来计划

- 应用更多的表头规范方法
- 支持从 JSON 修改现有的 Excel
- 支持设置包含或去掉哪些表单