# VBA for Excel
 
**In this repo I will dicuss about-**
- Basic of VBA as a programming language
- Recording some basic Macros
- VBA terminology
- Writting some Macros
- Debugging Macros
- Adding the Macros into Ribbon or interface


<br />

### Basic of VBA
VBA stands for Visual Basic Applications.
**VBA is tool for programming, editing and running application code.** Here apllication mean Excel. So thats means actually run code using Excel. VBA is not standalone program it work based on host application like Excel. It is Microsoft Event-driven program.

- custom design of our own functions
- Macros are very hard to undo

### Different way to express the same Range
*indicating C3 cell*
- Range("C3")
- Range("b2").Range("b2")
- Cells(3,3)
- [C3]
- Range("A1").Offset(2,2)

<br />

### selection, Value, Color, Clearing value
```
` Selecting a Cell
Range("C3").Select

` Input a value in a Cell
Selection.Value = "Hello World"

` Change color to yellow to the cell
Selection.Interior.Color = vbYellow

` Clear the value to the cell
Range("C3").Clear
```
<br />

### ActiveSheet, Sheets and Name
```
` Change current sheet name
ActiveSheet.Name = "Porfolio"

` Select a sheet named "sheet1"
sheets("sheet1").Select
```
<br />

### Select Current Region
```
Selection.CurrentRegion.Select
```

<br />

### Basic VBA Scripting
1. **Changing font Styles using VBA (Font Styles.xlsx).**<br />
    1. 
    
	```
    Sub TimeNewRoman()

    ` TimeNewRoman Macro
    ` Keyboard Shortcut: Ctrl+Shift+T

        Cells.Select
        Selection.Font.Name = "Times New Roman"
    End Sub
    ```

    2. 
    
    ```
    Sub TimeNewRoman()
    '
    ' TimeNewRoman Macro
    '
    ' Keyboard Shortcut: Ctrl+Shift+T
    '
        Cells.Select
        With Selection.Font
            .Name = "Times New Roman"
        End With
    End Sub
    ```

    3. From macro recorder

    ```
    Sub verdona()
    '
    ' verdona Macro
    '
    ' Keyboard Shortcut: Ctrl+Shift+S
    '
        With Selection.Font
            .Name = "Verdana"
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
    End Sub
    ```
2. 
