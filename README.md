<div align="center">

## Format and Print a MSFlexGrid


</div>

### Description

Format and Print a MSFlexGrid in a letter sheet, vertical or landscape, adding a Title, a logo, date and time of printing.
 
### More Info
 
MSFlexGrid to be printed

Title

Landscape (true/false)

Works for letter size, suggestions welcome.

Void

Grids with many columns (35 or more) could fullfill printers memory (low memory printers) causing the application to stop.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/format-and-print-a-msflexgrid__1-8186/archive/master.zip)





### Source Code

```
Sub PrintGrid(pGrid As MSFlexGrid, sTitulo As String, pHorizontal As Boolean)
' pGrid = The grid to print
' sTitulo = Page Title
' pHorizontal = True for Landscape
  On Error GoTo ErrorImpresion
  Dim i As Integer
  Dim iMaxRow As Integer
  Dim j As Integer
  Dim msfGrid As MSFlexGrid
  Dim iPaginas As Integer
  Printer.ColorMode = vbPRCMMonochrome
  Printer.PrintQuality = 160
  ' fMainForm.MSFlexGrid1 is an invisible msflexgrid
  ' used only for this routine
  ' put it where your want and reference it apropiately
  Set msfGrid = fMainForm.MSFlexGrid1
  msfGrid.FixedCols = 0
  msfGrid.Clear
  If pHorizontal = True Then
    Printer.Orientation = vbPRORLandscape
    iMaxRow = 44
  Else
    Printer.Orientation = vbPRORPortrait
    iMaxRow = 57
  End If
  ' calcula el número de páginas
  If pGrid.Rows Mod iMaxRow = 0 Then
    iPaginas = pGrid.Rows \ iMaxRow
  Else
    iPaginas = pGrid.Rows \ iMaxRow + 1
  End If
  msfGrid.Rows = iMaxRow
  msfGrid.Cols = pGrid.Cols
  For i = 0 To pGrid.Cols - 1
    msfGrid.ColWidth(i) = pGrid.ColWidth(i)
  Next
  screen.mousepointer = 11 ' hourglass
  ' print some logo -> comment or change as desired
  Printer.PaintPicture fMainForm.ImageList1.ListImages(1).Picture, 0, 0, 4300, 600
  ' imprime título
  Printer.CurrentY = 650
  Printer.FontName = "Courier New"
  Printer.FontBold = True
  Printer.FontSize = 12
  Printer.Print sTitulo
  Printer.Print
  ' justifica a la derecha fecha de impresión
  If pHorizontal = True Then
    Printer.CurrentX = 10000
  Else
    Printer.CurrentX = 7000
  End If
  Printer.CurrentY = 0
  Printer.FontSize = 10
  Printer.Print Now & " - Pág 1 de " & iPaginas
  For i = 0 To pGrid.Rows - 2 + iPaginas
    If i Mod iMaxRow = 0 And i > 0 Then
      With msfGrid
        .Row = 0
        .Col = 0
        .ColSel = 0
        .RowSel = 0
        If pHorizontal Then
          Printer.PaintPicture .Picture, 20, 1250, 15000, 10350
        Else
          Printer.PaintPicture .Picture, 20, 1250, 11400, 13950
        End If
      End With
      Printer.NewPage
      msfGrid.Clear
      For j = 0 To msfGrid.Cols - 1
         ' restablece títulos
        msfGrid.TextMatrix(0, j) = pGrid.TextMatrix(0, j)
      Next
      ' print logo
      Printer.PaintPicture fMainForm.ImageList1.ListImages(23).Picture, 0, 0, 4300, 600
      Printer.CurrentY = 650
      Printer.FontSize = 12
      Printer.Print sTitulo
      Printer.Print
      ' justifica a la derecha fecha de impresión
      If pHorizontal = True Then
        Printer.CurrentX = 10000
      Else
        Printer.CurrentX = 7000
      End If
      Printer.CurrentY = 0
      Printer.FontSize = 10
      Printer.Print Now & " - Pág " & i \ iMaxRow + 1 & " de " & iPaginas
      i = i + 1 ' deja títulos
    End If
    For j = 0 To msfGrid.Cols - 1
      msfGrid.TextMatrix(i Mod iMaxRow, j) = pGrid.TextMatrix(i - i \ iMaxRow, j)
    Next
  Next
  With msfGrid
    .Row = 0
    .Col = 0
    .ColSel = 0
    .RowSel = 0
    If pHorizontal Then
      Printer.PaintPicture .Picture, 20, 1250, 15000, 10350
    Else
      Printer.PaintPicture .Picture, 20, 1250, 11400, 13950
    End If
  End With
  Printer.EndDoc
  MsgBox sTitulo & vbCrLf & "Se ha(n) enviado " & iPaginas & " página(s) a la impresora " & Printer.DeviceName, vbInformation, Printer.Port
salir:
  Set msfGrid = Nothing
  pubCursorDefault
  Exit Sub
ErrorImpresion:
  Printer.KillDoc
  MsgBox "Verify printer", vbCritical, "Printer Error"
  Resume salir
End Sub
```

