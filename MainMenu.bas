Attribute VB_Name = "MainMenu"
Option Private Module
Sub MostrarHojas()

    Dim Hoja As Worksheet
    
    For Each Hoja In Worksheets
        If Hoja.CodeName <> "Hoja1" Then
            Hoja.Visible = xlSheetVisible
      End If
    Next Hoja
    
End Sub
Sub OcultarHojas()

    Dim Hoja As Worksheet
    
    For Each Hoja In Worksheets
        If Hoja.CodeName <> "Hoja1" Then
            Hoja.Visible = xlSheetVeryHidden
      End If
    Next Hoja
    
End Sub
Sub Menu_Principal()
    
                If Hoja8.Range("H1") = "Administrador" Then
                    frm_Menu.CommandButton2.Enabled = True
                    frm_Menu.CommandButton9.Enabled = True
                    frm_Menu.CommandButton10.Enabled = True
                    frm_Menu.CommandButton11.Enabled = True
                    'frm_Menu.CommandButton12.Enabled = True
                    frm_Menu.CommandButton13.Enabled = True
                End If
    frm_Menu.Show
End Sub
Sub Formatear_Reporte()
'
' Formatear_Reporte Macro
'
    Range("A1:M1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Size = 12
    Selection.Font.Size = 14
    Selection.Font.Size = 16
    Selection.Font.Size = 18
    Selection.Font.Size = 20
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Código"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Nombre"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "Descripción"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "Existencia"
    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A3:A6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B9").Select
    ActiveCell.FormulaR1C1 = "Comprb."
    Range("C9").Select
    ActiveCell.FormulaR1C1 = "Fecha"
    Range("D9").Select
    ActiveCell.FormulaR1C1 = "Factura"
    Range("E9").Select
    ActiveCell.FormulaR1C1 = "Proveedor"
    Range("F9").Select
    ActiveCell.FormulaR1C1 = "Tipo Entrada"
    Range("G9").Select
    ActiveCell.FormulaR1C1 = "Cant. Entrada"
    Range("H9").Select
    ActiveCell.FormulaR1C1 = "Comprb."
    Range("I9").Select
    ActiveCell.FormulaR1C1 = "Fecha"
    Range("J9").Select
    ActiveCell.FormulaR1C1 = "Factura"
    Range("K9").Select
    ActiveCell.FormulaR1C1 = "Destino"
    Range("L9").Select
    ActiveCell.FormulaR1C1 = "Tipo Salida"
    Range("M9").Select
    ActiveCell.FormulaR1C1 = "Cant. Salida"
    Range("M9").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("A3").Select
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").EntireColumn.AutoFit
    Columns("M:M").EntireColumn.AutoFit
    Range("B8:G8").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "ENTRADAS"
    Range("H8:M8").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "SALIDAS"
    Range("B8:G8").Select
    Selection.Font.Size = 12
    Selection.Font.Size = 14
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H8:M8").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Size = 12
    Selection.Font.Size = 14
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B8:G8").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("A3").Select

End Sub

Sub prueba()
Hoja1.Visible = xlSheetVisible
End Sub
