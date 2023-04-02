# ExcelファイルからODBC接続のクエリー情報をPythonで抽出する

## 動機

- 自分はデータ分析基盤のインフラの方を担当しBIツールとしてExcelは使わない
- データ利用者のExcelファイルをもらって調べるにあたり中に埋め込まれたクエリーを見たいがドライバーの設定がちゃんとしてないとグレーアウトされてSQL文を取り出せない様子
- ExcelのODBC接続の仕組みについてはよくわかってない
- Excel側でドライバーや設定をきちんとするのも良いが、とりあえずクエリー情報が欲しいだけなので、xlsxファイルに情報は入っているはずだから取り出そう

ここら辺の問題に近い

https://answers.microsoft.com/ja-jp/msoffice/forum/all/%E3%83%94%E3%83%9C%E3%83%83%E3%83%88%E3%83%86/cc275ce3-bff0-45a6-aad0-9fad38ee2205

https://learn.microsoft.com/en-US/office/troubleshoot/excel/cannot-modify-odata-connection

https://www.stadlersoftware.com/training/edit-excel-connections-in-xml/

## Excelファイルの構造

openpyxlでなんとかなるのかなぁと思ったけどちょっと見当たらなかったので直接XMLを操作する。

xlsxファイルをzipファイルとして展開して、`xl/connections.xml`を見てみるとここにクエリーの情報が記載されている。

https://learn.microsoft.com/ja-jp/dotnet/api/documentformat.openxml.office2013.excel.dbcommand?view=openxml-2.8.1

抜粋するとこんな感じ

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr16"
    xmlns:xr16="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16">
    <connection id="1" xr16:uid="{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXXX}" name="connection-name" type="100" refreshedVersion="7" savePassword="1" saveData="1" credentials="stored" singleSignOnId="username">
        <extLst>
            <ext uri="{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"
                xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
                <x15:connection id="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx">
                    <x15:oledbPr connection="Provider=MSDASQL.1;Persist Security Info=True;User ID=username;DSN=ODBC-Connection-Name;Password=password">
                        <x15:dbCommand text="SELECT * FROM mytable"/>
                    </x15:oledbPr>
                </x15:connection>
            </ext>
        </extLst>
    </connection>
</connections>
```

## Pythonで取り出す

```python
import sys
import xml.etree.ElementTree as ET
import zipfile


def get_queries(filename):
    """
    Get External queries from Excel file
    """
    # Unzip connection file
    with zipfile.ZipFile(filename) as zf:
        xml = zf.read("xl/connections.xml")

    root = ET.fromstring(xml)

    for connections in root.findall('connections:connection',
                                     {'connections': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}):
        for connection in connections.findall('.//x15:connection/x15:oledbPr/x15:dbCommand',
                                               {'x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'}):
            print(connection.attrib.get('text'))
            print()


if __name__ == '__main__': 
    get_queries(sys.argv[1])
```

めでたし。

```shell
$ python get_queries_from_xlsx.py myfile.xlsx

SELECT * FROM mytable
```
