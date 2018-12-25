# UnityExcelTool
A Unity Excel Tool, which can convert excel files to json, and auto generate table class and ScriptableObjects.

![image1](https://github.com/byfine/UnityExcelTool/raw/master/Readme/p1.jpg)   
![image2](https://github.com/byfine/UnityExcelTool/raw/master/Readme/p2.jpg)   
![image3](https://github.com/byfine/UnityExcelTool/raw/master/Readme/p3.jpg)   

## How to use:
"Window -> Excel Tool Window" to open it.  

1. Choose source file folder.
2. Choose output folder.
3. If you don't see any excel files, click refresh button.
4. Click Json button to Generate Json Files.
5. Click Class button to Generate Table Class Files.
6. Click Asset button to Generate ScriptableObjects Files.
7. Sheet Count means how many sheet you want to conver in every table. If this number is 1, the auto generated class will be named same as table filename, otherwise, it will be named TableName_SheetName.
   
## Notes:
1. In every table, the first 3 rows have a special purpose.
   1. First row : name of data.
   2. Second row : type of data, only support base type and Enum(not recommend).
   3. Third row : comment of data, could be empty.
2. First column data will be the key to find table.
3. All Excel filenames cannot be duplicated (even though they are not in same folders).

## Environment: 
Unity 2018.3.0f2, .Net 4.x, if you are using .Net 2.x, maybe you need to replace dlls in lib to proper version.
[ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader), [Newtonsoft.Json](https://github.com/JamesNK/Newtonsoft.Json)

-------

# UnityExcelTool
一个Unity Excel工具，可以把excel表格转换为Json文件，并且可以自动生成表格类和ScriptableObject文件。

## 使用:
"Window -> Excel Tool Window" 打开窗口.   

1. 选择源文件夹.
2. 选择目标文件夹.
3. 如果没有看到表格列表，点击 Refresh按钮.
4. 点击 Json按钮 生成 Json文件.
5. 点击 Class按钮 生成 类文件.
6. 点击 Asset按钮 生成 ScriptableObject文件.
7. Sheet Count 指每个表格要读几个sheet. 如果这个数是 1, 生成的文件会自动命名为表格的名字, 否则会命名为 表格名_Sheet名.
   
## 注意:
1. 对于每个表，前三行都有特殊意义
   1. 表格第一行：数据名称
   2. 表格第二行：数据类型, 仅支持基本类型和Enum（不推荐）
   3. 表格第三行：注释，可以为空
2. 第一列数据将被用作查找表的Key
3. 表格名不能重复 (即使文件在不同的目录).

## 开发环境: 
Unity 2018.3.0f2, .Net 4.x, 如果你是用的是 .Net 2.x, 可能要更换lib文件夹下dlls到对应的版本.
[ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader), [Newtonsoft.Json](https://github.com/JamesNK/Newtonsoft.Json)
