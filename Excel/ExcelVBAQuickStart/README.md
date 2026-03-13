# 10个案例快速上手Excel VBA

- [10个案例快速上手Excel VBA](#10个案例快速上手excel-vba)
  - [例一：第一个VBA程序 - 弹窗](#例一第一个vba程序---弹窗)
  - [例二：变量定义与单元格赋值 - 公司销售信息](#例二变量定义与单元格赋值---公司销售信息)
  - [例三：IF判断 - 成绩判断](#例三if判断---成绩判断)
  - [例四：Select Case - 多路径选择改造“成绩判断”](#例四select-case---多路径选择改造成绩判断)
  - [例五：for循环](#例五for循环)

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

## 例二：变量定义与单元格赋值 - 公司销售信息

VBA中变量声名的语法是：`Dim ... As ...`。

引用单元格的方法：
1. `Range("A1")`
2. `Cells(1,1).value`

![02_result](img/02_result.png)

![02_code](img/02_code.png)

源代码：

```VB
<!-- Sheet1: (General).companyrevenue -->
Sub companyrevenue()
    Dim company As String
    Dim renuew As Integer
    company = "ÎÒµÄ¹«Ë¾"
    revenue = 560000
    Range("B1:B2") = company
    Cells(2, 3).Value = revenue
    Cells(2, 4).Value = revenue * 2
End Sub
```

## 例三：IF判断 - 成绩判断

使用`=ROUND(RAND()*100,0)`可以产生0到100之间的随机整数。

![03_1](img/03_1.png)

![03_2](img/03_2.png)

源代码：

```VB
Sub scoreCheck()
    Dim rowNumber As Integer
    For rowNumber = 2 To 21
        If Cells(rowNumber, 2).Value >= 90 Then
            Cells(rowNumber, 3).Value = "优秀"
            Cells(rowNumber, 3).Interior.Color = RGB(0, 255, 0)
        ElseIf Cells(rowNumber, 2).Value >= 60 Then
            Cells(rowNumber, 3).Value = "及格"
            Cells(rowNumber, 3).Interior.Color = RGB(255, 255, 0)
        Else
            Cells(rowNumber, 3).Value = "不及格"
            Cells(rowNumber, 3).Interior.Color = RGB(255, 0, 0)
        End If
    Next rowNumber
End Sub

Sub clearCheck()
    Dim rowNumber As Integer
    For rowNumber = 2 To 21
        Cells(rowNumber, 3).Value = ""
        Cells(rowNumber, 3).Interior.Color = RGB(255, 255, 255)
    Next rowNumber
End Sub
```

## 例四：Select Case - 多路径选择改造“成绩判断”

语法结构为：

```VB
Select Csae variable
  Case Is condition1
    statement1
  Case condition2
    statement2
  ...
  Case Else
    statement-x
End Select
```

![04](img/04.png)

源代码：

```VB
Sub scoreCheck()
    Dim rowNumber As Integer
    Dim score As Integer
    For rowNumber = 2 To 21
        score = Cells(rowNumber, 2).Value
        Select Case score
            Case Is >= 90
                Cells(rowNumber, 3).Value = "优秀"
                Cells(rowNumber, 3).Interior.Color = RGB(0, 255, 0)
            Case 80 To 89
                Cells(rowNumber, 3).Value = "良好"
                Cells(rowNumber, 3).Interior.Color = RGB(255, 255, 0)
            Case 60 To 79
                Cells(rowNumber, 3).Value = "及格"
                Cells(rowNumber, 3).Interior.Color = RGB(255, 128, 128)
            Case Else
                Cells(rowNumber, 3).Value = "不及格"
                Cells(rowNumber, 3).Interior.Color = RGB(255, 0, 0)
        End Select
    Next rowNumber
End Sub

Sub clearCheck()
    Dim rowNumber As Integer
    For rowNumber = 2 To 21
        Cells(rowNumber, 3).Value = ""
        Cells(rowNumber, 3).Interior.Color = RGB(255, 255, 255)
    Next rowNumber
End Sub
```

## 例五：for循环

语法：`For i = 1 to 1000 ... Next i`