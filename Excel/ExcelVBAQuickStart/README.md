# 10个案例快速上手Excel VBA

- [10个案例快速上手Excel VBA](#10个案例快速上手excel-vba)
  - [例一：第一个VBA程序 - 弹窗](#例一第一个vba程序---弹窗)
  - [例二：变量定义与单元格赋值](#例二变量定义与单元格赋值)
  - [例三：for循环](#例三for循环)

## 例一：第一个VBA程序 - 弹窗

包含VBA的Excel文件要保存为`*.xlsm`扩展名。

在`Excel Option`>`Customize Ribbon`的右侧点选`Developer`以显示开发者菜单组：

![developer menu](img/excel_developer_menu.png)

若希望VBA代码能在Excel文件打开时自动运行，可以将代码写入`ThisWorkbook`对象中的`Open`事件中去，如下：

![sample01-code-run-at-open](img/sample01-code-run-at-open.png)

本例中的源代码：

```VB
<!-- Workbook.Open -->
Private Sub Workbook_Open()
    MsgBox "Hello World!"
    MsgBox "欢迎学习和使用VBA"
End Sub
```

## 例二：变量定义与单元格赋值

VBA中变量声名的语法是：`Dim ... As ...`。

引用单元格的方法：
1. `Range("A1")`
2. `Cells(1,1)`

## 例三：for循环

语法：`For i = 1 to 1000 ... Next i`