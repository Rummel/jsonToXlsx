My first C# program. It converts a JSON-file to Excel (*.xlsx)

sample JSON
```
{
    "fileNameSource": "HelloWorld.xlsx",
    "fileNameTarget": "HelloWorld1.xlsx",
    "worksheets": [
        {
            "name": "testWB",
            "cells": [
                {
                    "cell": "A10",
                    "value": [
                        [
                            "xa",
                            10,
                            20
                        ],
                        [
                            "xb2",
                            20.0
                        ],
                        [
                            "xc3",
                            30.1
                        ],
                        [
                            "xa4"
                        ]
                    ]
                },
                {
                    "cell": "B2",
                    "dataType": "Number",
                    "style": {
                        "color": 150
                    },
                    "value": "2"
                },
                {
                    "cell": "B1",
                    "value": false
                },
                {
                    "cell": "C2",
                    "style": {
                        "color": 150
                    },
                    "value": 2
                },
                {
                    "cell": "B3",
                    "dataType": "Boolean",
                    "style": {
                        "color": 60
                    },
                    "value": "2"
                },
                {
                    "cell": "B4",
                    "dataType": "Text",
                    "style": {
                        "color": 215907
                    },
                    "value": 77.5
                }
            ]
        }
    ]
}
```



compile
```
dotnet publish -p:PublishSingleFile=true -r win-x64 -c Release --self-contained false .\ExcelTool.sln
```

or
```
dotnet publish -p:PublishSingleFile=true -r win-x64 -c Release --self-contained true .\ExcelTool.sln
```

run
```
jsonToXlsx.exe HelloWorld.json
```

