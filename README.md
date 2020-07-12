# excelSheetsCombination
A simple tool to copy sheets from one excel file to another, and insert the copied sheets to a certain place.

# 功能：
从EXCEL文件A中复制部分表单（目前表单选择依据为表单名）
将从A中复制的表单依次插入EXCEL文件B中的特定位置
同时修改表单名（表单名中包含表单当前所在位置序号）、特定单元格内容（单元格包含表单名）

# 使用方式：
windows下直接打开./dist/excelInsert.exe文件，输入想要插入的位置、选择计算机中文件B与文件A
或
直接运行excelInsert.py文件

# 包含文件说明：
excelInsert.py为主体文件

cvt_icon2py.py文件和icon.py文件以及./icon_src/cat14.ico用于设置界面上的图标。其中，运行cvt_icon2py.py会自动生成icon.py文件（其实是将./icon_src/cat14.ico文件进行base64编码后进行暂时存放）

./icon_src/1.jpg和./icon_src/cuter.ico用于通过pyinstaller生成可执行文件exe的图标。其中，cuter.ico使用1.jpg和格式转化工具png2ico得到，cuter.ico中包含了四种不同尺寸供windows显示。

其余文件为使用pyinstaller生成可执行文件exe时得到的附属产物

