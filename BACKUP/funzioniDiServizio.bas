Attribute VB_Name = "funzioniDiServizio"
Sub MacroRaggruppaIntestazioneSettimana()
'
' MacroRaggruppaIntestazioneSettimana Macro
' raggruppa le celle della colonna WEEK
'

'
   Dim cellDelWEEK As Range
  ' Dim Rslt As Range
  ' Dim c1 As Range
  ' Dim c2 As Range
   

   ' raggruppamento turno precedente
  ' Columns("D:P").Columns.Group

   
   UltimaRigaX = Range("A10").End(xlDown).Row
   Set Rslt = Range(FindAll("WEEK", Application.Intersect(ActiveSheet.UsedRange, Range(Range("W9"), Range("W9").End(xlToRight))), xlFormulas, xlPart, SearchFormat:=True).Address)
   For Each cellDelWEEK In Rslt
         c1 = cellDelWEEK.Offset(1, 0).Address
         c2 = cellDelWEEK.Offset(UltimaRigaX - 9, 0).Address
         weekNumber = cellDelWEEK.Offset(-1, 8).Value ' cella nei titoli che contiene il numero di week
         dataLunedi = cellDelWEEK.Offset(-1, 1).Value
         Range(c1, c2).Select
         Selection.Merge
         ActiveCell.FormulaR1C1 = "WEEK " & weekNumber & " da lunedi " & dataLunedi
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter 'xlTop
            .WrapText = False
            .Orientation = 90
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            .Font.Name = "Calibri"
            .Font.Size = 18
        End With
        
        'colore della cella
         With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        

        Col_letter_C1 = Split(cellDelWEEK.Offset(2, 1).Address(True, False), "$")(0)  'ricava la lettera di colonna della cella
        Col_letter_C2 = Split(cellDelWEEK.Offset(2, 10).Address(True, False), "$")(0) 'ricava la lettera di colonna della cella

        colonne = Col_letter_C1 & ":" & Col_letter_C2
        Columns(colonne).Columns.Group
        With ActiveSheet.Outline
           .AutomaticStyles = False
           .SummaryRow = xlBelow
           .SummaryColumn = xlLeft
        End With
      '  Columns("c1:c2").Select
        
   Next
   
End Sub



Sub Colora_NSR_RC()
'
' Macro6 Macro
'

'
    ActiveWindow.SmallScroll Down:=-30
    Cells.Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""N"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""S"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16751204
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""R"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""RC"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16777164
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub




Sub Cancella_Module()
Dim vbMod As Object
    Set vbMod = Application.VBE.ActiveVBProject.VBComponents
    vbMod.Remove VBComponent:=vbMod.Item("Module3")
End Sub
Public Function WeeksInYear(lYear As Long) As Long

    WeeksInYear = DatePart("ww", DateSerial(lYear, 12, 28), vbMonday, vbFirstFourDays)

End Function

Public Function ISOweeknum(ByVal v_Date As Date) As Integer
    ISOweeknum = DatePart("ww", v_Date - Weekday(v_Date, 2) + 4, 2, 2)
End Function

Public Function MondayWeek(intYear As Integer, _
intWeek As Integer) As Date
    Dim dat As Date
    Dim intD As Integer
    
    'calc first day of year of datFld
    dat = DateSerial(intYear, 1, 1)
    intD = Weekday(dat)
    Select Case intD
    Case 1
    dat = dat + 1
    MondayWeek = DateAdd("ww", intWeek - 1, dat)
    Case Else
    'subtract 2 to get first Monday
    dat = dat - (intD - 2)
    MondayWeek = DateAdd("ww", intWeek, dat)
    End Select
    'if datYear is
    '2006 1/2/2006 is first monday
    '2007 1/1/2007 is first monday
    '2008 12/31/2007 is first monday
    
    'add number of weeks minus 1 to first monday
  '  MondayWeek = DateAdd("ww", intWeek - 1, dat)
  ' MondayWeek = DateAdd("ww", intWeek, dat)
End Function


 
Function FindAll(What, Optional SearchWhat As Variant, _
        Optional LookIn, _
        Optional LookAt, _
        Optional SearchOrder, _
        Optional SearchDirection As XlSearchDirection = xlNext, _
        Optional MatchCase As Boolean = False, _
        Optional MatchByte, _
        Optional SearchFormat) As Range
    'LookIn can be xlValues or xlFormulas, _
     LookAt can be xlWhole or xlPart, _
     SearchOrder can be xlByRows or xlByColumns, _
     SearchDirection can be xlNext, xlPrevious, _
     MatchCase, MatchByte, and SearchFormat can be True or False. _
     Before using SearchFormat = True, specify the appropriate settings _
     for the Application.FindFormat object, e.g., _
     Application.FindFormat.NumberFormat = "General;-General;""-"""
    Dim aRng As Range
    If IsMissing(SearchWhat) Then
        On Error Resume Next
        Set aRng = ActiveSheet.UsedRange
        On Error GoTo 0
    ElseIf TypeOf SearchWhat Is Range Then
        If SearchWhat.Cells.Count = 1 Then
            Set aRng = SearchWhat.Parent.UsedRange
        Else
            Set aRng = SearchWhat
            End If
    ElseIf TypeOf SearchWhat Is Worksheet Then
        Set aRng = SearchWhat.UsedRange
    Else
        Exit Function                       '*****
        End If
    If aRng Is Nothing Then Exit Function   '*****
    Dim FirstCell As Range, CurrCell As Range
    With aRng.Areas(aRng.Areas.Count)
    Set FirstCell = .Cells(.Cells.Count)
        'This little 'dance' ensures we get the first matching _
         cell in the range first
        End With
    Set FirstCell = aRng.Find(What:=What, After:=FirstCell, _
        LookIn:=LookIn, LookAt:=LookAt, _
        SearchDirection:=SearchDirection, MatchCase:=MatchCase, _
        MatchByte:=MatchByte, SearchFormat:=SearchFormat)
    If FirstCell Is Nothing Then Exit Function          '*****
    Set CurrCell = FirstCell
    Set FindAll = CurrCell
    Do
        Set FindAll = Application.Union(FindAll, CurrCell)
        'Setting FindAll at the top of the loop ensures _
         the result is arranged in the same sequence as _
         the  matching cells; the duplicate assignment of _
         the first matching cell to FindAll being a small _
         price to pay for the ordered result
        Set CurrCell = aRng.Find(What:=What, After:=CurrCell, _
            LookIn:=LookIn, LookAt:=LookAt, _
            SearchDirection:=SearchDirection, MatchCase:=MatchCase, _
            MatchByte:=MatchByte, SearchFormat:=SearchFormat)
        'FindNext is not reliable because it ignores the FindFormat settings
        Loop Until CurrCell.Address = FirstCell.Address
    End Function


Sub ExamplesOfFindAll()
    'reset any prior find format condition
    Application.FindFormat.Clear
   
    'show the address of the range in the activesheet _
     that contains a value of 1
    MsgBox FindAll(1, , xlValues, xlWhole).Address
   
    'show the address of the range in the activesheet _
     that contains 1 as any part of the value
    MsgBox FindAll(1, , xlValues, xlPart).Address
   
    'show the address of the range in the activesheet _
     where the formula contains a open paren
    MsgBox FindAll("(", , xlFormulas, xlPart).Address
   
    'show the address of the cells in column C of the activesheet _
     that contain a zero
    Application.FindFormat.Clear
    Dim Rslt As Range
    MsgBox FindAll(0, Range("c:c"), xlFormulas, xlWhole).Address
   
    'if a custom number format applies to the entire column C, the below _
     will cause a major performance headache because the find will step _
     through every cell in column C!
    'MsgBox FindAll("", Range("c:c"), _
        xlFormulas, xlPart, SearchFormat:=True).Address
 
    'An alternative to the above is to limit the search to the usedrange.
    Application.FindFormat.Clear
    Application.FindFormat.NumberFormat = "General;-General;""-"""
    MsgBox FindAll("", Application.Intersect( _
            ActiveSheet.UsedRange, Range("c:c")), _
        xlFormulas, xlPart, SearchFormat:=True).Address
 
    'show the address of the range in column C that contains a zero and _
     the specified custom number format
    Application.FindFormat.Clear
    Application.FindFormat.NumberFormat = "General;-General;""-"""
    MsgBox FindAll(0, Range("c:c"), _
        xlFormulas, xlWhole, SearchFormat:=True).Address
   
    'show the address of the range of cells in column C within the _
     activesheet's usedrange that have a fill color of xlThemeColorAccent2
    Application.FindFormat.Clear
    Application.FindFormat.Interior.ThemeColor = xlThemeColorAccent2
    MsgBox FindAll("", Application.Intersect( _
            ActiveSheet.UsedRange, Range("c:c")), _
        xlFormulas, xlPart, SearchFormat:=True).Address
    End Sub


Sub validator_simboliturno(r)
'
' Macro7 Macro
'

'
    Range(r).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:= _
        xlBetween, Formula1:="=simboliturno"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Simbolo non previsto"
        .InputMessage = ""
        .ErrorMessage = "Opportuno usare dei simbili previsti."
        .ShowInput = True
        .ShowError = True
    End With
End Sub

