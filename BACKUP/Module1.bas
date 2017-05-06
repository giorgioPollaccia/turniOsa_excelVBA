Attribute VB_Name = "Module1"
Dim Anno As String
Dim weekIniziale As Long
Dim weekFinale As Long
Dim location As String
Const numColPerWeek = 11
Const rigaIntestazioneData = 8



Private Function OpenConnection() As ADODB.connection
    ' Read type and location of the database, user login and password
    Dim source As String, location As String, user As String, password As String
    source = Range("Source").Value
    location = Range("Location").Value
    'user = TasksSheet.UserInput.Value
    'password = TasksSheet.PasswordInput.Value
 
    ' Handle relative path for the location of Access and SQLite database files
    If (source = "Access" Or source = "SQLite") And Not location Like "?:\*" Then
        location = ActiveWorkbook.Path & "\" & location
    End If
 
    ' Build the connection string depending on the source
    Dim connectionString As String
    Select Case source
        Case "Access"
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & location
        Case "MySQL"
            connectionString = "Driver={MySQL ODBC 5.2a Driver};Server=" & location & ";Database=test;UID=" & user & ";PWD=" & password
        Case "PostgreSQL"
            connectionString = "Driver={PostgreSQL ANSI};Server=" & location & ";Database=test;UID=" & user & ";PWD=" & password
        Case "SQLite"
            connectionString = "Driver={SQLite3 ODBC Driver};Database=" & location
    End Select
 
    ' Create and open a new connection to the selected source
    Set OpenConnection = New ADODB.connection
    Call OpenConnection.Open(connectionString)
End Function

Public Function min_hour(minu)
    Dim flag_segno_meno As Boolean
    
    If minu < 0 Then
       minu = -minu
       flag_segno_meno = True
    Else
       flag_segno_meno = False
    End If
    
   If minu < 10 Then
         'min_hour = CDate("00:0" & minu)
         min_hour = "00:0" & minu
    ElseIf minu >= 10 And minu < 60 Then
        'min_hour = CDate("00:" & minu)
         min_hour = "00:" & minu
    ElseIf minu >= 60 Then
    ore1 = Int(minu / 60)
    minu = minu Mod 60
    If ore1 < 10 Then
    ore1 = "0" & ore1
    End If
    
    If minu < 10 Then
    minu = "0" & minu
    End If
   ' min_hour = CDate(ore1 & ":" & minu)
    min_hour = ore1 & ":" & minu
    If flag_segno_meno = True Then
        min_hour = "-" & min_hour
    End If
    
    End If
    
    
End Function

Public Function hour_min(hour)
       If (hour = "") Then
         hour_min = 0
       Else
        hm = Split(hour, ":")
        h = hm(0)
        M = hm(1)
        If h < 0 Then
            hour_min = h * 60 - M  'le ore:minuti sono negative
        Else
            hour_min = h * 60 + M
        End If
       End If
End Function



Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = v(Y, X)
        Next Y
    Next X
    
    TransposeDim = tempArray


End Function
Public Sub LoadTasksButton_Click()
    Dim output As Range
    Dim recArray As Variant
 
    Dim connection As connection
    Set connection = OpenConnection()
 
    Dim result As ADODB.Recordset

    ' Load all the tasks from the database
    Set result = connection.Execute("SELECT *  FROM reparti")
    
    ' Copy field names to the first row of the worksheet
    For icols = 0 To result.Fields.Count - 1
             Cells(12, icols + 1).Value = result.Fields(icols).Name
        Next
    Range(Cells(12, 1), Cells(12, result.Fields.Count)).Font.Bold = True
   
    fldCount = result.Fields.Count
    recArray = result.GetRows
    ' Determine number of records
    recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array
    
    For iCol = 0 To fldCount - 1
            For iRow = 0 To recCount - 1
                ' Take care of Date fields
                If IsDate(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                ' Take care of OLE object fields or array fields
                ElseIf IsArray(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = "Array Field"
                End If
            Next iRow 'next record
     Next iCol 'next field
            
        ' Transpose and Copy the array to the worksheet,
        ' starting in cell A2
        Cells(13, 1).Resize(recCount, fldCount).Value = TransposeDim(recArray)
    connection.Close
     ' Auto-fit the column widths and row heights
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
End Sub


' ricava il turno precedente dal database filtrando TURNAZIONE_OPERATORE per anno e weekIniziale -1
Sub MacroTurnoPrecedente()
'
'
  ' "SELECT ANAGRAFICA_0.Matricola, ANAGRAFICA_0.Cognome, ANAGRAFICA_0.Nome, TURNAZIONE_OPERATORE_0.IDMATRICE, TURNAZIONE_OPERATORE_0.PROGRESSIVO, TURNAZIONE_OPERATORE_0.d"
    location = Sheets("comandi").Range("Location").Value
    Dim nomeTabellaOuput As String
    nomeTabellaOutput = Anno & "_" & weekIniziale & "_" & weekFinale
   '  With ActiveSheet.ListObjects.Add(SourceType:=0, source:=Array(Array( _
   '     "ODBC;DSN=SQLite3 Datasource;Database=" & location & ";StepAPI=0;SyncPragma=NORMAL;NoTXN=0;Timeout=10000" _
   '     ), Array( _
   '     "0;ShortNames=0;LongNames=0;NoCreat=0;NoWCHAR=0;FKSupport=0;JournalMode=;OEMCP=0;LoadExt=;BigInt=0;JDConv=0;" _
   '     )), Destination:=Range("$A$10")).QueryTable
   '     .CommandText = Array( _
   '     "SELECT ANAGRAFICA_0.Matricola, ANAGRAFICA_0.Cognome, ANAGRAFICA_0.Nome, TURNAZIONE_OPERATORE_0.IDMATRICE, TURNAZIONE_OPERATORE_0.PROGRESSIVO, TURNAZIONE_OPERATORE_0.d" _
   '     , _
   '     "ata_lun, TURNAZIONE_OPERATORE_0.week, TURNAZIONE_OPERATORE_0.anno, TURNAZIONE_OPERATORE_0.lun, TURNAZIONE_OPERATORE_0.mar, TURNAZIONE_OPERATORE_0.mer, TURNAZIONE_OPERATORE_0.gio, TURNAZIONE_OPERATORE_" _
   '     , _
   '     "0.ven, TURNAZIONE_OPERATORE_0.sab, TURNAZIONE_OPERATORE_0.dom" & Chr(13) & "" & Chr(10) & "FROM ANAGRAFICA ANAGRAFICA_0, TURNAZIONE_OPERATORE TURNAZIONE_OPERATORE_0" & Chr(13) & "" & Chr(10) & "WHERE ANAGRAFICA_0.Matricola = TURNAZIONE_OPERATORE_0.MATRICOL" _
   '     , _
   '     "A AND ((TURNAZIONE_OPERATORE_0.anno='" & Anno & "') AND (TURNAZIONE_OPERATORE_0.week='" & weekIniziale - 1 & "'))" _
   '     )
   '     .RowNumbers = False
   '     .FillAdjacentFormulas = False
   '     .PreserveFormatting = True
   '     .RefreshOnFileOpen = False
   '     .BackgroundQuery = True
   '     .RefreshStyle = xlInsertDeleteCells
   '     .SavePassword = False
   '     .SaveData = True
   '     .AdjustColumnWidth = True
   '     .RefreshPeriod = 0
   '     .PreserveColumnInfo = True
   '     .ListObject.DisplayName = "tab_" & nomeTabellaOutput 'Nome della tabella di dati
   '     .Refresh BackgroundQuery:=False
   ' End With
   
  
     
     With ActiveSheet.ListObjects.Add(SourceType:=0, source:=Array(Array( _
        "ODBC;DSN=SQLite3 Datasource;Database=" & location & ";StepAPI=0;SyncPragma=NORMAL;NoTXN=0;Timeout=10000" _
        ), Array( _
        "0;ShortNames=0;LongNames=0;NoCreat=0;NoWCHAR=0;FKSupport=0;JournalMode=;OEMCP=0;LoadExt=;BigInt=0;JDConv=0;" _
        )), Destination:=Range("$A$9")).QueryTable
    
           .CommandText = Array( _
            "SELECT ANAGRAFICA_0.Matricola, ANAGRAFICA_0.Cognome, ANAGRAFICA_0.Nome " _
            & "FROM ANAGRAFICA ANAGRAFICA_0" _
            )
 
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "tab_" & nomeTabellaOutput 'Nome della tabella di dati
        .Refresh BackgroundQuery:=False
   End With

    
    With ActiveSheet.ListObjects.Add(SourceType:=0, source:=Array(Array( _
        "ODBC;DSN=SQLite3 Datasource;Database=" & location & ";StepAPI=0;SyncPragma=NORMAL;NoTXN=0;Timeout=10000" _
        ), Array( _
        "0;ShortNames=0;LongNames=0;NoCreat=0;NoWCHAR=0;FKSupport=0;JournalMode=;OEMCP=0;LoadExt=;BigInt=0;JDConv=0;" _
        )), Destination:=Range("$j$9")).QueryTable
        
       If (weekIniziale > 1) Then
        .CommandText = Array( _
        "SELECT TURNAZIONE_OPERATORE_0.IDMATRICE, TURNAZIONE_OPERATORE_0.PROGRESSIVO, TURNAZIONE_OPERATORE_0.d" _
        , _
        "ata_lun, TURNAZIONE_OPERATORE_0.week, TURNAZIONE_OPERATORE_0.anno, TURNAZIONE_OPERATORE_0.lun, TURNAZIONE_OPERATORE_0.mar, TURNAZIONE_OPERATORE_0.mer, TURNAZIONE_OPERATORE_0.gio, TURNAZIONE_OPERATORE_" _
        , _
        "0.ven, TURNAZIONE_OPERATORE_0.sab, TURNAZIONE_OPERATORE_0.dom" & Chr(13) & "" & Chr(10) & "FROM ANAGRAFICA ANAGRAFICA_0, TURNAZIONE_OPERATORE TURNAZIONE_OPERATORE_0" & Chr(13) & "" & Chr(10) & "WHERE ANAGRAFICA_0.Matricola = TURNAZIONE_OPERATORE_0.MATRICOL" _
        , _
        "A AND ((TURNAZIONE_OPERATORE_0.anno='" & Anno & "') AND (TURNAZIONE_OPERATORE_0.week='" & weekIniziale - 1 & "'))" _
        )
        Else  '''  weekiniziale = 1 PRIMA SETTIMANA DI UN NUOVO ANNO; occorre recuperare il turno precedente rerlativo quindi all'ultima settimana dell'anno precedente
            .CommandText = Array( _
            "SELECT TURNAZIONE_OPERATORE_0.IDMATRICE, TURNAZIONE_OPERATORE_0.PROGRESSIVO, TURNAZIONE_OPERATORE_0.d" _
            , _
            "ata_lun, TURNAZIONE_OPERATORE_0.week, TURNAZIONE_OPERATORE_0.anno, TURNAZIONE_OPERATORE_0.lun, TURNAZIONE_OPERATORE_0.mar, TURNAZIONE_OPERATORE_0.mer, TURNAZIONE_OPERATORE_0.gio, TURNAZIONE_OPERATORE_" _
            , _
            "0.ven, TURNAZIONE_OPERATORE_0.sab, TURNAZIONE_OPERATORE_0.dom" & Chr(13) & "" & Chr(10) & "FROM ANAGRAFICA ANAGRAFICA_0, TURNAZIONE_OPERATORE TURNAZIONE_OPERATORE_0" & Chr(13) & "" & Chr(10) & "WHERE ANAGRAFICA_0.Matricola = TURNAZIONE_OPERATORE_0.MATRICOL" _
            , _
            "A AND ((TURNAZIONE_OPERATORE_0.anno='" & Anno - 1 & "') AND (TURNAZIONE_OPERATORE_0.week='" & 52 & "'))" _
            )
        End If
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "tab_" & nomeTabellaOutput & "turnoBase" 'Nome della tabella di dati
        .Refresh BackgroundQuery:=False
    End With

'''' Raggrauppamente delle colonna relative al turno base
    UltimaRigaX = Range("A10").End(xlDown).Row
    c1 = Cells(10, 9).Address
    c2 = Cells(10, 9).Offset(UltimaRigaX - 10, 0).Address
    Range(c1, c2).Select
    Selection.Merge
      '   ActiveCell.Text = "WEEK "
       ActiveCell.FormulaR1C1 = "TURNO BASE DELLA WEEK " & weekIniziale - 1
       
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
        

       ' Col_letter_C1 = Split(cellDelWEEK.Offset(2, 1).Address(True, False), "$")(0)  'ricava la lettera di colonna della cella
       ' Col_letter_C2 = Split(cellDelWEEK.Offset(2, 10).Address(True, False), "$")(0) 'ricava la lettera di colonna della cella

        'colonne = Col_letter_C1 & ":" & Col_letter_C2
        Columns("J:V").Columns.Group
        With ActiveSheet.Outline
           .AutomaticStyles = False
           .SummaryRow = xlBelow
           .SummaryColumn = xlLeft
        End With
        
        
        
        
 
 

    'Range("tab_2016_45[[#Headers],[Matricola]]").Select
    Range("tab_" & Anno & "_" & weekIniziale & "_" & weekFinale & "[[#Headers],[Matricola]]").Select
    
 '   Selection.AutoFilter

    Columns("L:L").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = weekIniziale '"=R[3]C[6]+1"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = weekFinale
End Sub

'Fa uso del numero di settimana indicata nella cella(1,1)
'Scrive la data e il giorno della settimana in alto nella sheet
Sub IntestazioneDate_Week()

    Dim Months As Double
    Dim SecondDate As Date

    Dim ColonnaBase As Integer

    ColonnaBase = 24
   ' StartDate = InputBox("Enter a date")
   ' SecondDate = CDate(StartDate)
   ' Cells(2, 2) = SecondDate
    
    Dim dd As Date
    Dim datasistema As Date
    Dim costante As Integer
    Dim vettoregiorni(1 To 7) As String
    vettoregiorni(1) = "Lun"
    vettoregiorni(2) = "Mar"
    vettoregiorni(3) = "Mer"
    vettoregiorni(4) = "Gio"
    vettoregiorni(5) = "Ven"
    vettoregiorni(6) = "Sab"
    vettoregiorni(7) = "Dom"

    Sheets(Anno).Name = Anno & "_" & weekIniziale & "_" & weekFinale 'Nome della sheet con la settimana indicata
    For i = 0 To (weekFinale - weekIniziale)                                                              'X settimane
        Cells(rigaIntestazioneData + 1, ColonnaBase - 1 + i * numColPerWeek).Value = "WEEK"
        weekIniziale = Cells(1, 1).Text '+ i
        Cells(rigaIntestazioneData, ColonnaBase + 7 + i * numColPerWeek) = weekIniziale + i                   'data del lunedi della week
        'Cells(1, ColonnaBase + i * 10) = Week2DateNew(weekIniziale)                       'data del lunedi della week
        Cells(rigaIntestazioneData, ColonnaBase + i * numColPerWeek) = MondayWeek(CInt(Anno), weekIniziale + i)
        costante = Weekday(Week2DateNew(weekIniziale), vbUseSystemDayOfWeek)    'Calcola il nome del giorno dalla data
        Cells(rigaIntestazioneData + 1, ColonnaBase + i * numColPerWeek) = vettoregiorni(costante)                        'Scrive il nome del giorno
        'Intestazioni per i giorni della settimana
        For j = 1 To 6
            dd = Cells(rigaIntestazioneData, ColonnaBase + i * numColPerWeek + j - 1)
            'Cells(1, 4 + i * 7 + j) = dd
            Cells(rigaIntestazioneData, ColonnaBase + i * numColPerWeek + j) = DateAdd("D", 1, CDate(Cells(rigaIntestazioneData, ColonnaBase + i * numColPerWeek + j - 1).Value))
            Cells(rigaIntestazioneData + 1, ColonnaBase + i * numColPerWeek + j) = vettoregiorni(costante + j)
        Next
        Cells(rigaIntestazioneData + 1, ColonnaBase + 8 + i * numColPerWeek).Value = "ORE"
        Cells(rigaIntestazioneData + 1, ColonnaBase + 9 + i * numColPerWeek).Value = "MONTE_ORE"
    Next
    
    
    
  
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    ' grassetto e colore delle prime due righe
    Rows("8:9").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.699993896298105
        .PatternTintAndShade = 0
    End With
    
     ' nasconde le colonne al momento non usate
     Columns("E:H").Select
     Selection.EntireColumn.Hidden = True
    Range("A1").Select
End Sub


'Funzione di servizio: restituisce la data del lunedi di una data week
Public Function Week2DateNew(WeekNo As Long, Optional ByVal Yr As Long = 0, _
                        Optional ByVal DOW As VBA.VbDayOfWeek = VBA.VbDayOfWeek.vbUseSystemDayOfWeek, _
                         Optional ByVal FWOY As VBA.VbFirstWeekOfYear = VBA.VbFirstWeekOfYear.vbUseSystem) As Date
    ' Returns First Day of week
    Dim Jan1 As Date
    Dim Sub1 As Boolean
    Dim Ret As Date
    If Yr = 0 Then
      Jan1 = VBA.DateSerial(VBA.Year(VBA.Date()), 1, 1)
    Else
      Jan1 = VBA.DateSerial(Yr, 1, 1)
    End If
    Sub1 = (VBA.Format(Jan1, "ww", DOW, FWOY) = 1)
    Ret = VBA.DateAdd("ww", WeekNo + Sub1, Jan1)
    Ret = Ret - VBA.Weekday(Ret, DOW) + 1
    Week2DateNew = Ret
End Function




Sub CreaSheetTurno()

    Anno = InputBox("Anno")
    weekIniziale = InputBox("Settimana iniziale del nuovo turno")
    weekFinale = InputBox("Settimana finale del nuovo turno")

  ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' VIENE CREATA UNA SHEET "anno_mese" se non esiste
    Dim wsTest As Worksheet
    Dim strSheetName As String
    Set wsTest = Nothing
    On Error Resume Next
    'strSheetName = Year(Now())
    Set wsTest = ActiveWorkbook.Worksheets(Anno)
    On Error GoTo 0
    If wsTest Is Nothing Then
        Worksheets.Add(Sheets(Worksheets.Count), Count:=1, Type:=xlWorksheet).Name = Anno
    End If
    Cells.Select
    Selection.ClearContents

     Call MacroTurnoPrecedente ' ricava il turno precedente dal database filtrando TURNAZIONE_OPERATORE per anno e week
     'Anno = strSheetName
     Call IntestazioneDate_Week
     
     
     '  VIENE CREATO UN BUTTON CON LA MACRO ASSOCIATA
     Dim btn As Button
     Application.ScreenUpdating = False
     Set t = ActiveSheet.Range(Cells(2, 2), Cells(3, 2))
     Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
     With btn
      .OnAction = "btnGeneraTemplate"  ' FUNZIONA ASSOCIATA AL TASTO
      .Caption = "Genera template"
      .Name = "BtnGeneraTemplate"
    End With
    Application.ScreenUpdating = True
    
     '  VIENE CREATO UN BUTTON CON LA MACRO ASSOCIATA
     Dim btnSalvaTurno As Button
     Application.ScreenUpdating = False
     Set t = ActiveSheet.Range(Cells(5, 2), Cells(6, 2))
     Set btnSalvaTurno = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
     With btnSalvaTurno
      .OnAction = "btnSalvaTurnoInDataBase"  ' FUNZIONA ASSOCIATA AL TASTO
      .Caption = "Salva Turno"
      .Name = "BtnGeneraTemplate"
    End With
    
    '  VIENE CREATO UN BUTTON CON LA MACRO ASSOCIATA
     Dim btnequalizzaMattinaPomeriggio As Button
     Application.ScreenUpdating = False
     Set t = ActiveSheet.Range(Cells(2, 4), Cells(3, 4))
     Set btnSalvaTurno = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
     With btnSalvaTurno
      .OnAction = "equalizzaMattinaPomeriggio"  ' FUNZIONA ASSOCIATA AL TASTO
      .Caption = "Forza 9 P"
      .Name = "BtnequalizzaMattinaPomeriggio"
    End With
    Application.ScreenUpdating = True
    

    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    ' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
End Sub

''  Salva il turno visualizzato nella sheet nel data base TURNAZIONE_OPERATORE
''  Effettua un primo inserimento del record nel database usando il turno del lunedi
Sub btnSalvaTurnoInDataBase()
     Dim cellDelLunedi As Range
     Dim Rslt1 As Range
     
     Dim connection As connection
     
     
     rigainiziale = 10 ' prima matricola
     UltimaRigaX = Range("A10").End(xlDown).Row
     
     answer = MsgBox("Vuoi salvare il turno definito nella sheet sul databse ?", vbYesNo + vbQuestion, "Salva turno.")
    
     If answer = vbYes Then
        Set connection = OpenConnection()
     
        Set Rslt = Range(FindAll("Lun", Application.Intersect(ActiveSheet.UsedRange, Range(Range("X9"), Range("X9").End(xlToRight))), xlFormulas, xlPart, SearchFormat:=True).Address)

        ' ciclo for su ogni lunedi indicato nella riga2 della sheet
        ' cellDelLunedi viene usato come riferimento per ricavare i dati dalle celle della sheet da caricare nel database
        ' ricavo data , anno e week per l'iserimento del turno nel database
        For Each cellDelLunedi In Rslt
            dataLunedi = cellDelLunedi.Offset(-1, 0).Value
            week = ISOweeknum(dataLunedi)
            dataLunedi_xDB = Format(dataLunedi, "YYYY-MM-dd")
            anno_xDB = Year(dataLunedi)
            
            For r = rigainiziale To UltimaRigaX              ''ciclo for per riga ; quante sono le matricole sulla colonna A
                matricola = Cells(r, 1).Text
                IdMat_Progressivo = Split(Cells(r, cellDelLunedi.Column + 7).Text, ";")
                idmatrice = IdMat_Progressivo(0)
                progressivo = IdMat_Progressivo(1)
                ID = matricola & "_" & Format(dataLunedi, "YYYY-MM-dd")
                
                lun = Cells(r, cellDelLunedi.Column).Text
                mar = Cells(r, cellDelLunedi.Column + 1).Text
                mer = Cells(r, cellDelLunedi.Column + 2).Text
                gio = Cells(r, cellDelLunedi.Column + 3).Text
                ven = Cells(r, cellDelLunedi.Column + 4).Text
                sab = Cells(r, cellDelLunedi.Column + 5).Text
                dom = Cells(r, cellDelLunedi.Column + 6).Text
                
                
                
                
                ' si inserisce o si modifica il turno;
                QueryinsertReplace = "INSERT or replace INTO turnazione_operatore (id, matricola, idmatrice, progressivo, data_lun, week, anno, lun, mar, mer, gio, ven, sab, dom ) VALUES (" & _
                            "'" & ID & "', " & _
                            "'" & matricola & "', " & _
                            "'" & idmatrice & "', " & _
                            "'" & progressivo & "', " & _
                            "'" & dataLunedi_xDB & "', " & _
                            "'" & week & "', " & _
                            "'" & anno_xDB & "', " & _
                            "'" & lun & "', " & _
                            "'" & mar & "', " & _
                            "'" & mer & "', " & _
                            "'" & gio & "', " & _
                            "'" & ven & "', " & _
                            "'" & sab & "', " & _
                            "'" & dom & "'" & _
                            ")"
                
      
                
                connection.BeginTrans
                connection.Execute QueryinsertReplace
                connection.CommitTrans
                
                ' si completa il turno aggiornandolo con i vaolori veri dei turni mar, mer, gio, ven, sab e dom
                
                
        
            Next
         Next cellDelLunedi
         connection.Close
         Set connection = Nothing
     Else  ' risposta NO nel msgBox
        'do nothing
     End If
     
     rs = MsgBox("Turno descritto nella sheet caricato nel database TURNAZIONE_OPERATORE.", vbInformation, "Upload turno.")

End Sub





'' Crea un turno template nelle settimane scelte.
'' Viene usata come base il turno della settimana N-1 ricavata da database  (OVVIAMENTE deve essere presente un turno coorente)
Sub btnGeneraTemplate()
     rigainiziale = 10 ' prima matricola
     UltimaRigaX = Range("A10").End(xlDown).Row
     colonnaIniziale = 23 'colonna precedente a  dove posizionare il primo girno
     
     
    
     Dim matricola As String
     Dim idmatrice As String
     'Dim progressivo As String
     
     Dim connection As connection
     Set connection = OpenConnection()
     Dim result As ADODB.Recordset
     Dim progressivo As Integer
     
     Application.AutoCorrect.AutoFillFormulasInLists = False  'previene che la formula si propaghi per tutte le righe sulla colonna della tabella in maniera uguale , invece si vuole una formula per ogni righa
     
     Application.ScreenUpdating = False
      
     weekIniziale = Cells(1, 1).Value
     weekFinale = Cells(2, 1).Value
     
      '  calcolo delle risorse per i vari giorni
     Cells(UltimaRigaX + 1, 3).Value = "MATTINA"
     Cells(UltimaRigaX + 2, 3).Value = "POMERIGGIO"
     Cells(UltimaRigaX + 3, 3).Value = "NOTTURNO"
     Cells(UltimaRigaX + 4, 3).Value = "SMONTAGGIO"
     Cells(UltimaRigaX + 5, 3).Value = "FERIE"
     Cells(UltimaRigaX + 6, 3).Value = "RIPOSO"
     Cells(UltimaRigaX + 7, 3).Value = "RIPOSO_RC"
     Cells(UltimaRigaX + 8, 3).Value = "MALATTIA"
     Cells(UltimaRigaX + 9, 3).Value = "L104"
     For i = 0 To (weekFinale - weekIniziale) 'scrive la formula sotto la tabella per il calcolo delle risorse giornaliere; per ogni settimana
      For j = 1 To 7 ' per ogni giorno della settimana, cioe per ogni colonna scrive la formula per calcolare quante risorse per i vari turni M,P ecc
        Cells(UltimaRigaX + 1, colonnaIniziale + j + i * numColPerWeek).Formula = "=countif(" & Cells(rigainiziale, colonnaIniziale + j + i * numColPerWeek).Address & ":" & Cells(UltimaRigaX, colonnaIniziale + j + i * numColPerWeek).Address & "," & Chr(34) & "M" & Chr(34) & ")"
        Range(Cells(UltimaRigaX + 1, colonnaIniziale + j + i * numColPerWeek).Address).Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=12"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        Cells(UltimaRigaX + 2, colonnaIniziale + j + i * numColPerWeek).Formula = "=countif(" & Cells(rigainiziale, colonnaIniziale + j + i * numColPerWeek).Address & ":" & Cells(UltimaRigaX, colonnaIniziale + j + i * numColPerWeek).Address & "," & Chr(34) & "P" & Chr(34) & ")"
        Range(Cells(UltimaRigaX + 2, colonnaIniziale + j + i * numColPerWeek).Address).Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=7"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        Cells(UltimaRigaX + 3, colonnaIniziale + j + i * numColPerWeek).Formula = "=countif(" & Cells(rigainiziale, colonnaIniziale + j + i * numColPerWeek).Address & ":" & Cells(UltimaRigaX, colonnaIniziale + j + i * numColPerWeek).Address & "," & Chr(34) & "N" & Chr(34) & ")"
        Range(Cells(UltimaRigaX + 3, colonnaIniziale + j + i * numColPerWeek).Address).Select
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=2"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        
        Cells(UltimaRigaX + 4, colonnaIniziale + j + i * numColPerWeek).Formula = "=countif(" & Cells(rigainiziale, colonnaIniziale + j + i * numColPerWeek).Address & ":" & Cells(UltimaRigaX, colonnaIniziale + j + i * numColPerWeek).Address & "," & Chr(34) & "S" & Chr(34) & ")"
        Cells(UltimaRigaX + 5, colonnaIniziale + j + i * numColPerWeek).Formula = "=countif(" & Cells(rigainiziale, colonnaIniziale + j + i * numColPerWeek).Address & ":" & Cells(UltimaRigaX, colonnaIniziale + j + i * numColPerWeek).Address & "," & Chr(34) & "F" & Chr(34) & ")"
        Cells(UltimaRigaX + 6, colonnaIniziale + j + i * numColPerWeek).Formula = "=countif(" & Cells(rigainiziale, colonnaIniziale + j + i * numColPerWeek).Address & ":" & Cells(UltimaRigaX, colonnaIniziale + j + i * numColPerWeek).Address & "," & Chr(34) & "R" & Chr(34) & ")"
        Cells(UltimaRigaX + 7, colonnaIniziale + j + i * numColPerWeek).Formula = "=countif(" & Cells(rigainiziale, colonnaIniziale + j + i * numColPerWeek).Address & ":" & Cells(UltimaRigaX, colonnaIniziale + j + i * numColPerWeek).Address & "," & Chr(34) & "RC" & Chr(34) & ")"
        Cells(UltimaRigaX + 8, colonnaIniziale + j + i * numColPerWeek).Formula = "=countif(" & Cells(rigainiziale, colonnaIniziale + j + i * numColPerWeek).Address & ":" & Cells(UltimaRigaX, colonnaIniziale + j + i * numColPerWeek).Address & "," & Chr(34) & "MAL" & Chr(34) & ")"
        Cells(UltimaRigaX + 9, colonnaIniziale + j + i * numColPerWeek).Formula = "=countif(" & Cells(rigainiziale, colonnaIniziale + j + i * numColPerWeek).Address & ":" & Cells(UltimaRigaX, colonnaIniziale + j + i * numColPerWeek).Address & "," & Chr(34) & "L104" & Chr(34) & ")"
      Next
     Next
    
     
     
     
   
     
     For r = rigainiziale To UltimaRigaX              ''ciclo for per riga ; quante sono le matricole sulla colonna A
        matricola = Cells(r, 1).Text
        idmatrice = Cells(r, 10).Text
        progressivo = Cells(r, 11).Text
         
         
        
        For i = 0 To (weekFinale - weekIniziale) 'genera N settimane del turno incrementando il progressivo e tenendo idmatrice
        
                                
                        
           
           ' viene ricavato il progressivo massimo per quella tipologiaDiMatrice (ad esempio OSA_NOTTURNO ha un modulo =11
           Query = "select count(*) from matriceturni where IDMatrice='" & idmatrice & "'"
       
           Set queryMaxProgressivo = connection.Execute(Query)
           MaxProgressivo = queryMaxProgressivo.Fields(0)
           If progressivo = MaxProgressivo Then
                 progressivo = 0
            End If
        
        
            '' SE ESISTE UN TURNO NEL DATABASE RELATIVO A QUELLA SETTIMANA ALLORA VIENE CARICATO QUELLO
            ' verifico se nel database dei turni turnazione_operatore  esiste un turno per quella matricola e per la settimana
            lunediDellasettimana = Cells(rigaIntestazioneData, colonnaIniziale + 1 + i * numColPerWeek).Value
            
            Query = "SELECT lun,mar,mer,gio,ven,sab,dom,idmatrice,progressivo  FROM turnazione_operatore where id=" & " '" & matricola & "_" & Format(Cells(rigaIntestazioneData, colonnaIniziale + 1 + i * numColPerWeek).Value, "YYYY-mm-dd") & "'"
            Set resultQuery1 = New ADODB.Recordset
            resultQuery1.CursorLocation = adUseClient
            resultQuery1.Open Query, connection
           ' resultQuery1 = connection.Execute(Query)
           ' fldCount = resultQuery1.Fields.Count
           ' recArray = resultQuery1.GetRows
           If resultQuery1.RecordCount Then         ''''' se un turno con id matrice_dataLunedi è presente nel database allora viene caricato
                recArray1 = resultQuery1.GetRows
                
                Cells(r, colonnaIniziale + 1 + i * numColPerWeek).Value = recArray1(0, 0)
                Call validator_simboliturno(Cells(r, colonnaIniziale + 1 + i * numColPerWeek).Address)
                Cells(r, colonnaIniziale + 2 + i * numColPerWeek).Value = recArray1(1, 0)
                Call validator_simboliturno(Cells(r, colonnaIniziale + 2 + i * numColPerWeek).Address)
                Cells(r, colonnaIniziale + 3 + i * numColPerWeek).Value = recArray1(2, 0)
                Call validator_simboliturno(Cells(r, colonnaIniziale + 3 + i * numColPerWeek).Address)
                Cells(r, colonnaIniziale + 4 + i * numColPerWeek).Value = recArray1(3, 0)
                Call validator_simboliturno(Cells(r, colonnaIniziale + 4 + i * numColPerWeek).Address)
                Cells(r, colonnaIniziale + 5 + i * numColPerWeek).Value = recArray1(4, 0)
                Call validator_simboliturno(Cells(r, colonnaIniziale + 5 + i * numColPerWeek).Address)
                Cells(r, colonnaIniziale + 6 + i * numColPerWeek).Value = recArray1(5, 0)
                Call validator_simboliturno(Cells(r, colonnaIniziale + 6 + i * numColPerWeek).Address)
                Cells(r, colonnaIniziale + 7 + i * numColPerWeek).Value = recArray1(6, 0)
                Call validator_simboliturno(Cells(r, colonnaIniziale + 7 + i * numColPerWeek).Address)
                resultQuery1.Close
                Cells(rigaIntestazioneData + 1, colonnaIniziale + 8 + i * numColPerWeek).Value = "Tdatab"
                
            Else                                     ''''  ALTRIMENTI VINE GENERATO UN TURNO SEQUNZIALE
            
            

                        
                        ' Viene ricavato il turno dalla matriceturni con lo stesso idmatrice e progressivo incrementato
                        Query = "SELECT lun,mar,mer,gio,ven,sab,dom  FROM matriceturni where idmatrice='" & idmatrice & "' and progressivo= '" & progressivo + 1 & "'"
                        Set result = connection.Execute(Query)
                        
                        fldCount = result.Fields.Count
                        recArray = result.GetRows
                        
                        Cells(r, colonnaIniziale + 1 + i * numColPerWeek).Value = recArray(0, 0)
                        Call validator_simboliturno(Cells(r, colonnaIniziale + 1 + i * numColPerWeek).Address)
                        Cells(r, colonnaIniziale + 2 + i * numColPerWeek).Value = recArray(1, 0)
                        Call validator_simboliturno(Cells(r, colonnaIniziale + 2 + i * numColPerWeek).Address)
                        Cells(r, colonnaIniziale + 3 + i * numColPerWeek).Value = recArray(2, 0)
                        Call validator_simboliturno(Cells(r, colonnaIniziale + 3 + i * numColPerWeek).Address)
                        Cells(r, colonnaIniziale + 4 + i * numColPerWeek).Value = recArray(3, 0)
                        Call validator_simboliturno(Cells(r, colonnaIniziale + 4 + i * numColPerWeek).Address)
                        Cells(r, colonnaIniziale + 5 + i * numColPerWeek).Value = recArray(4, 0)
                        Call validator_simboliturno(Cells(r, colonnaIniziale + 5 + i * numColPerWeek).Address)
                        Cells(r, colonnaIniziale + 6 + i * numColPerWeek).Value = recArray(5, 0)
                        Call validator_simboliturno(Cells(r, colonnaIniziale + 6 + i * numColPerWeek).Address)
                        Cells(r, colonnaIniziale + 7 + i * numColPerWeek).Value = recArray(6, 0)
                        Call validator_simboliturno(Cells(r, colonnaIniziale + 7 + i * numColPerWeek).Address)
                        Cells(rigaIntestazioneData + 1, colonnaIniziale + 8 + i * numColPerWeek).Value = "Tnuovo"
                        
            
             End If
            Cells(r, colonnaIniziale + 8 + i * numColPerWeek).Value = idmatrice & ";" & progressivo + 1
            
            '''' vengono calcolati le ore fatte durante la settimana, ponendo la formula
            FormulaOreFatteSettimana = "min_hour("
            'FormulaOreFatteSettimana = ""
            For giorno = 1 To 7  ''   ciclo for dal lun ... dom
                indirizzoCellaDelGiorno = Cells(r, colonnaIniziale + giorno + i * numColPerWeek).Address 'restituisce le coordinate riga, colonna della cella
                FormulaOreFatteSettimana = FormulaOreFatteSettimana & _
                               "+if(" & indirizzoCellaDelGiorno & "=" & Chr(34) & "M" & Chr(34) & "," & Chr(34) & "420" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & ")" & _
                               "+if(" & indirizzoCellaDelGiorno & "=" & Chr(34) & "P" & Chr(34) & "," & Chr(34) & "420" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & ")" & _
                               "+if(" & indirizzoCellaDelGiorno & "=" & Chr(34) & "F" & Chr(34) & "," & Chr(34) & "420" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & ")" & _
                               "+if(" & indirizzoCellaDelGiorno & "=" & Chr(34) & "N" & Chr(34) & "," & Chr(34) & "600" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & ")" & _
                               "+if(" & indirizzoCellaDelGiorno & "=" & Chr(34) & "M1" & Chr(34) & "," & Chr(34) & "380" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & ")" & _
                               "+if(" & indirizzoCellaDelGiorno & "=" & Chr(34) & "P1" & Chr(34) & "," & Chr(34) & "380" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & ")" & _
                               "+if(" & indirizzoCellaDelGiorno & "=" & Chr(34) & "L104" & Chr(34) & "," & Chr(34) & "420" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & ")" & _
                               "+if(" & indirizzoCellaDelGiorno & "=" & Chr(34) & "MAL" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & ")" & _
                               "+if(" & indirizzoCellaDelGiorno & "=" & Chr(34) & "RC" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & "," & Chr(34) & "0" & Chr(34) & ")"
            Next
            FormulaOreFatteSettimana = FormulaOreFatteSettimana & ")"
            ' mostra le ore fatte nella settimana
            Cells(r, colonnaIniziale + 9 + i * numColPerWeek).Formula = "=" & FormulaOreFatteSettimana
            progressivo = progressivo + 1
          
          
          
          '''''''''''''''''''  calcolo monte ore nelle 12 settimane precedenti
          '''''''''''''''''''
            indirizzoCellaDeiMinutiSettimanali = Cells(r, colonnaIniziale + 9 + i * numColPerWeek).Address
            'Cells(2, colonnaIniziale + 9 + i * numColPerWeek).Value = "ORE"
            indirizzoCellaDeiMinutiSettimanaliPrecedente = Cells(r, colonnaIniziale + i * numColPerWeek - 1).Address
          '  Cells(2, colonnaIniziale + 10 + i * numColPerWeek).Value = "MONTE_ORE"
            If i = 0 Then  '' prima settimana; si stampa nella colonna relativa alla week-1 il monte ore calcolato sulle week precedenti
                 Cells(9, colonnaIniziale - 1).Value = "M_ORE"
                Formula1 = CalcolaResiduoNelle12SettimanePrecedenti_dalDataBase(matricola, i) 'minuti fatti nelle N settimane precedenti
                Cells(r, colonnaIniziale - 1).Formula = "=min_hour(" & Formula1 & ")"
                Cells(r, colonnaIniziale + 10 + i * numColPerWeek).Formula = "=min_hour( hour_min(" & indirizzoCellaDeiMinutiSettimanali & ")+" & Formula1 & "-2280 )"
            Else
                Formula1 = 0
                Cells(r, colonnaIniziale + 10 + i * numColPerWeek).Formula = "=min_hour( hour_min(" & indirizzoCellaDeiMinutiSettimanali & ")+" & "-2280 + hour_min(" & indirizzoCellaDeiMinutiSettimanaliPrecedente & "))"
            End If
            
            
            
            
             
            
        Next  ' fine ciclo for sulle settimane
        
        
     Next     ' fine ciclo for sulle righe delle risorse
      
    

    
     connection.Close
     
    Call Colora_NSR_RC
    
    
      Call MacroRaggruppaIntestazioneSettimana
      
      Application.ScreenUpdating = True
      
     
     Cells.Select
     Cells.EntireColumn.AutoFit
        
     ' splitta il workbook
     'ActiveWindow.SplitColumn = 3

     '' raggruppa le colonne
     'ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
      
      
     
      Range("A1").Select
     resp = MsgBox("Creato un turno template dalla settimana " & weekIniziale & " alla settimana " & weekFinale & " !", vbInformation, "Fine creazione turno template")
End Sub


Sub updateAnagrafica()
    Dim connection As connection
    Set connection = OpenConnection()

    connection.BeginTrans
    connection.Execute "delete from anagrafica" 'ATTENZIONE cancello i dati presenti nel database per aggiornarli con quelli attuali
    connection.CommitTrans
         
    numeroDirisorse = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 5 To numeroDirisorse
        matr = Cells(i, 1).Value
        cog = Cells(i, 2).Value
        N = Cells(i, 3).Value
        S = Cells(i, 4).Value
         ' Compose the INSERT statement.
        
       
         Queryinsert = "insert into anagrafica( " & _
        "matricola ,cognome, nome ,sesso)" & _
        " VALUES (" & _
        "'" & matr & "', " & _
        "'" & cog & "', " & _
        "'" & N & "', " & _
        "'" & S & "'" & _
        ")"
        connection.BeginTrans
        connection.Execute Queryinsert
        connection.CommitTrans
    Next

 
    
    connection.Close
    Set connection = Nothing
    ActiveWorkbook.RefreshAll
End Sub


Sub updateMATRICETURNI()
    Dim connection As connection
    Set connection = OpenConnection()

    connection.BeginTrans
    connection.Execute "delete from matriceturni" 'ATTENZIONE cancello i dati presenti nel database per aggiornarli con quelli attuali
    connection.CommitTrans
         
    numeroDirisorse = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 5 To numeroDirisorse

         ' Compose the INSERT statement.
        
       
         Queryinsert = "insert into matriceturni( " _
         & Cells(4, 1).Value & "," & Cells(4, 2).Value & "," & Cells(4, 3).Value & "," & Cells(4, 4).Value & "," _
         & Cells(4, 5).Value & "," & Cells(4, 6).Value & "," & Cells(4, 7).Value & "," & Cells(4, 8).Value & "," _
         & Cells(4, 9).Value & "," & Cells(4, 10).Value & "," & Cells(4, 11).Value & "," & Cells(4, 12).Value & ")" _
        & " VALUES (" & _
        "'" & Cells(i, 1).Value & "', " & _
        "'" & Cells(i, 2).Value & "', " & _
        "'" & Cells(i, 3).Value & "', " & _
        "'" & Cells(i, 4).Value & "', " & _
        "'" & Cells(i, 5).Value & "', " & _
        "'" & Cells(i, 6).Value & "', " & _
        "'" & Cells(i, 7).Value & "', " & _
        "'" & Cells(i, 8).Value & "', " & _
        "'" & Cells(i, 9).Value & "', " & _
        "'" & Cells(i, 10).Value & "', " & _
        "'" & Cells(i, 11).Value & "', " & _
        "'" & Cells(i, 12).Value & "'" & _
        ")"
        connection.BeginTrans
        connection.Execute Queryinsert
        connection.CommitTrans
    Next

 
    
    connection.Close
    Set connection = Nothing
    ActiveWorkbook.RefreshAll
End Sub
Sub updateASSENZE()
    Dim connection As connection
    Set connection = OpenConnection()

    connection.BeginTrans
    connection.Execute "delete from assenze" 'ATTENZIONE cancello i dati presenti nel database per aggiornarli con quelli attuali
    connection.CommitTrans
         
    numeroDirisorse = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 5 To numeroDirisorse

         ' Compose the INSERT statement.
        
         ID = Cells(i, 1).Value & "_" & Format(Cells(i, 4).Value, "YYYY-MM-dd")
         matricola = Cells(i, 1).Value
         Data = Cells(i, 4).Value
         causa = Cells(i, 5).Value
         
         Queryinsert = "insert into assenze( ID, MATRICOLA, DATA, CAUSA) VALUES (" _
         & "'" & ID & "', " _
         & "'" & matricola & "', " _
         & "'" & Data & "', " _
         & "'" & causa & "'" _
         & ")"
        connection.BeginTrans
        connection.Execute Queryinsert
        connection.CommitTrans
    Next

 
    
    connection.Close
    Set connection = Nothing
    ActiveWorkbook.RefreshAll
End Sub
Sub refreshDalDatase()
        ActiveWorkbook.RefreshAll
End Sub

Function CalcolaResiduoNelle12SettimanePrecedenti_dalDataBase(mat, i)
     Dim week_ii As Integer
     
     rigainiziale = 4 ' prima matricola
     UltimaRigaX = Range("A3").End(xlDown).Row
     colonnaIniziale = 23 'colonna precedente a  dove posizionare il primo girno

     Dim connection As connection
     Set connection = OpenConnection()
    ' Dim result As ADODB.Recordset
    
     weekIniziale_ = Cells(1, 1).Value
     weekFinale_ = Cells(2, 1).Value
   
     settimanePerIlMonteOre = 12
   
     minutiIn12Settimane = 2280 * settimanePerIlMonteOre '2280 minuti in una settimana per il numero di settimane
     

        minutiFatti = 0
        oreFatte = 0
        For week_ii = (i + weekIniziale - 1) To (i + weekIniziale - 1 - settimanePerIlMonteOre + 1) Step -1 ''  i è la settimana corrente, l'indice week_ii serve per leggere dal database il turno delle settimane precedenti
        
       ' For week_ii = (i + weekIniziale - 1) To (i + weekIniziale - 1) Step -1
         '  lunediDellaWeek = DateAdd("d", 1, Week2DateNew(week_ii)) 'data del lunedi della week
           lunediDellaWeek = MondayWeek(CInt(Anno), week_ii)
            ' verifico se nel database dei turni turnazione_operatore  esiste un turno per quella matricola e per la settimana
            Query = "SELECT lun,mar,mer,gio,ven,sab,dom,idmatrice,progressivo  FROM turnazione_operatore where id=" & " '" & mat & "_" & Format(lunediDellaWeek, "YYYY-mm-dd") & "'"
            Set resultQuery1 = New ADODB.Recordset
            resultQuery1.CursorLocation = adUseClient
            resultQuery1.Open Query, connection
            If resultQuery1.RecordCount Then         ''''' se un turno con id matrice_dataLunedi è presente nel database allora viene caricato
                 recArray1 = resultQuery1.GetRows
                 
                 For d = 0 To 6 ' minuti nella settiman
                    If ((recArray1(d, 0) = "MAL") Or (recArray1(d, 0) = "P") Or (recArray1(d, 0) = "M") Or (recArray1(d, 0) = "L104") Or (recArray1(d, 0) = "F")) Then
                        minutiFatti = minutiFatti + 420
                    End If
                    If ((recArray1(d, 0) = "N")) Then
                        minutiFatti = minutiFatti + 600
                    End If
                    If ((recArray1(d, 0) = "M1") Or (recArray1(d, 0) = "P1")) Then
                        minutiFatti = minutiFatti + 380
                    End If
                    
                    'If ((recArray1(d, 0) = "/")) Then  ' giorno NULLO il monte ore da fare si riduce di 7 ore
                    '     minutiIn12Settimane = minutiIn12Settimane + 420
                    'End If
                    
                 Next
            Else
                minutiIn12Settimane = minutiIn12Settimane - 2280 'se il turno non e' nel data base la media mobile verra' calcolata con le settimane con turno nuovo
            End If
        Next

      '  oreinSett = min_hour(minutiIn12Settimane)
      '  oreFatte = min_hour(minutiFatti)
      '  monteOre12Settimane = "text(Max(" & Chr(34) & oreinSett & Chr(34) & "," & Chr(34) & oreFatte & Chr(34) & " ) - Min(" & Chr(34) & oreinSett & Chr(34) & "," & Chr(34) & oreFatte & Chr(34) & ")," & Chr(34) & "-[h]:mm" & Chr(34) & ")"
      '  CalcolaResiduoNelle12SettimanePrecedenti_dalDataBase = monteOre12Settimane
        CalcolaResiduoNelle12SettimanePrecedenti_dalDataBase = minutiFatti - minutiIn12Settimane
        
   '  Next
     
End Function


Function timeDiff(t1 As Date, t0 As Date) As Date 'Return Time1 minus Time0
    Dim units(0 To 2) As String

    units(0) = hour(t1) - hour(t0)
    units(1) = Minute(t1) - Minute(t0)
    units(2) = Second(t1) - Second(t0)

    If units(2) < 0 Then
        units(2) = units(2) + 60
        units(1) = units(1) - 1
    End If

    If units(1) < 0 Then
        units(1) = units(1) + 60
        units(0) = units(0) - 1
    End If

    units(0) = IIf(units(0) < 0, units(0) + 24, units(0))
    timeDiff = Join(units, ":")
End Function


Sub equalizzaMattinaPomeriggio()
   ''  SI FA IN MODO CHE PER OGNI GIORNO CI SIANO NON PIU DI 9 RISORSE PER IL TURNO P
   Set Rslt = Range(FindAll("Lun", Application.Intersect(ActiveSheet.UsedRange, Range(Range("X9"), Range("X9").End(xlToRight))), xlFormulas, xlPart, SearchFormat:=True).Address)
   UltimaRigaX = Range("A10").End(xlDown).Row
   Dim sampleArr() As Variant
   Dim settimanacompleta As Range
   
    Set settimanacompleta = Range("O10:O45")
     For Each cellDelLunedi In Rslt   'per ogni lun trovato, quindi per ogni week
            'c2 = cellDelLunedi.Offset(UltimaRigaX - 9, 6).Address
            For i = 0 To 6           ' scorro per ogni giorno della settimama
                c1 = cellDelLunedi.Offset(1, i).Address
                c2 = cellDelLunedi.Offset(UltimaRigaX - 9, i).Address
                c3 = cellDelLunedi.Offset(15, i).Address '' parte bassa delle righe in maniera da alternare il change con priorita dalla parte bassa alla parte alta in base al giorno
                                                         '' altrimenti i cambi P-->M avverrebbero maggiormente per le stesse persone in alto nelle righe
                Range(c1, c2).Select
                nMattina = Application.CountIf(Range(c1, c2), "M")
               nPomeriggio = Application.CountIf(Range(c1, c2), "P")
                
                
                
                '''  SI CERCA DI ARRIVARE AD UN NUMERO DI P PARI A 9 PER OGNI GIORNO; SCAMBIANDO P CON M  SE SONO MAGGIORI DI 9
                tentativiDiChange = 0 ' un modo per uscire dal while se non c'è modi di portare il numero di P a 9  (per ferie permessi ecc)
                While ((nPomeriggio > 9) And (tentativiDiChange < 30))
                        tentativiDiChange = tentativiDiChange + 1
                        ' per alternare il cambiamento P in M sulle matricole in base al giorno ; il lunedi mercoledi e venerdi viene cercata la P dalla prima matricola a scendere
                        If ((i = 0) Or (i = 2) Or (i = 4) Or (i = 6)) Then
                             Set Rslt1 = Range(FindAll("P", Application.Intersect(ActiveSheet.UsedRange, Range(c1, c2)), xlFormulas, xlPart, SearchFormat:=True).Address)
                             For Each p In Rslt1
                               If ((p.Offset(0, -2).Value <> "S") And (i > 1)) Then
                                     '  p.Select
                                        p.Value = "M"
                                        Exit For  ' sostituiamo un solo P alla volta
                               End If
                               If ((p.Offset(0, -6).Value <> "S") And (i = 1)) Then 'per il martedi si controlla che 2 turni prima della settimana precedente non sia S , per evitare di cambiare SRP
                                      '  p.Select
                                        p.Value = "M"
                                        Exit For  ' sostituiamo un solo P alla volta
                               End If
                               If ((p.Offset(0, -7).Value <> "S") And (i = 0)) Then 'per il lundedi si controlla che 2 turni prima della settimana precedente non sia S , per evitare di cambiare SRP
                                     '   p.Select
                                        p.Value = "M"
                                        Exit For  ' sostituiamo un solo P alla volta
                               End If
                            Next
                        ElseIf ((i = 1) Or (i = 3) Or (i = 5)) Then                 ' per i girni di martedi , giovedi e sabato le P vengono cercate dall'ultima matricola a salire
                            'Set findP = Range(c1, c2).Find("P", , , , , xlPrevious)
                            'Set p = Range(findP.Address)
                            'p.Select
                             trovatoPNellaParteBassa = 0  'flag usato nel caso che non si trovino P da sostituire nella parte bassa allora si cercheranno nella parte alta
                             Set Rslt1 = Range(FindAll("P", Application.Intersect(ActiveSheet.UsedRange, Range(c3, c2)), xlFormulas, xlPart, SearchFormat:=True).Address)
                             For Each p In Rslt1
                                ' For lngCounter = Rslt1.Cells.Count To 1 Step -1
                                ' Set p = cells(lngCounter,).Address
                                'For Each p In Rslt1
                               If ((p.Offset(0, -2).Value <> "S") And (i > 1)) Then
                                     '  p.Select
                                       trovatoPNellaParteBassa = 1
                                        p.Value = "M"
                                        Exit For  ' sostituiamo un solo P alla volta
                               End If
                               If ((p.Offset(0, -6).Value <> "S") And (i = 1)) Then 'per il martedi si controlla che 2 turni prima della settimana precedente non sia S , per evitare di cambiare SRP
                                      '  p.Select
                                        trovatoPNellaParteBassa = 1
                                        p.Value = "M"
                                        Exit For  ' sostituiamo un solo P alla volta
                               End If
                               If ((p.Offset(0, -7).Value <> "S") And (i = 0)) Then 'per il lundedi si controlla che 2 turni prima della settimana precedente non sia S , per evitare di cambiare SRP
                                       ' p.Select
                                        trovatoPNellaParteBassa = 1
                                        p.Value = "M"
                                        Exit For  ' sostituiamo un solo P alla volta
                               End If
                            Next
                            If (trovatoPNellaParteBassa = 0) Then  '' non si sono trovati P da sostituire nella parte bassa allora si cercano nella parte alta
                             Set Rslt1 = Range(FindAll("P", Application.Intersect(ActiveSheet.UsedRange, Range(c1, c3)), xlFormulas, xlPart, SearchFormat:=True).Address)
                             For Each p In Rslt1
                                ' For lngCounter = Rslt1.Cells.Count To 1 Step -1
                                ' Set p = cells(lngCounter,).Address
                                'For Each p In Rslt1
                               If ((p.Offset(0, -2).Value <> "S") And (i > 1)) Then
                                      ' p.Select
                                       
                                        p.Value = "M"
                                        Exit For  ' sostituiamo un solo P alla volta
                               End If
                               If ((p.Offset(0, -6).Value <> "S") And (i = 1)) Then 'per il martedi si controlla che 2 turni prima della settimana precedente non sia S , per evitare di cambiare SRP
                                       ' p.Select
                                       
                                        p.Value = "M"
                                        Exit For  ' sostituiamo un solo P alla volta
                               End If
                               If ((p.Offset(0, -7).Value <> "S") And (i = 0)) Then 'per il lundedi si controlla che 2 turni prima della settimana precedente non sia S , per evitare di cambiare SRP
                                       ' p.Select
                                      
                                        p.Value = "M"
                                        Exit For  ' sostituiamo un solo P alla volta
                               End If
                            Next
                            End If
                            
                        End If
                        
                       
                      nPomeriggio = Application.CountIf(Range(c1, c2), "P")
                Wend  '''  mentre le P sono > di 9
                
                
                
                
                    '''  SI CERCA DI ARRIVARE AD UN NUMERO DI P PARI A 9 PER OGNI GIORNO; SCAMBIANDO M CON P  SE SONO MINORI DI 9
                tentativiDiChange = 0 ' un modo per uscire dal while se non c'è modi di portare il numero di P a 9  (per ferie permessi ecc)
                While ((nPomeriggio < 9) And (tentativiDiChange < 30))
                        tentativiDiChange = tentativiDiChange + 1
                        ' per alternare il cambiamento M in P sulle matricole in base al giorno ; il lunedi mercoledi e venerdi viene cercata la M dalla prima matricola a scendere
                        If ((i = 0) Or (i = 2) Or (i = 4) Or (i = 6)) Then
                             Set Rslt1 = Range(FindAll("M", Application.Intersect(ActiveSheet.UsedRange, Range(c1, c2)), xlFormulas, xlPart, SearchFormat:=True).Address)
                             For Each p In Rslt1
                               If ((p.Offset(0, 1).Value <> "N") And (i > 1)) Then
                                      'p.Select
                                        p.Value = "P"
                                        Exit For  ' sostituiamo un solo M alla volta
                               End If
                               If ((p.Offset(0, 5).Value <> "N") And (i = 6)) Then  'per la domenica si controlla che 1 turni prima della settimana successiva non sia N , per evitare di cambiare MNS
                                        'p.Select
                                        p.Value = "P"
                                        Exit For
                               End If
                            Next
                        ElseIf ((i = 1) Or (i = 3) Or (i = 5)) Then                 ' per i girni di martedi , giovedi e sabato le P vengono cercate dall'ultima matricola a salire
                             trovatoMNellaParteBassa = 0  'flag usato nel caso che non si trovino P da sostituire nella parte bassa allora si cercheranno nella parte alta
                             Set Rslt1 = Range(FindAll("M", Application.Intersect(ActiveSheet.UsedRange, Range(c3, c2)), xlFormulas, xlPart, SearchFormat:=True).Address)
                             For Each p In Rslt1
                              p.Select
                                If ((p.Offset(0, 1).Value <> "N")) Then
                                      ' p.Select
                                       trovatoMNellaParteBassa = 1
                                        p.Value = "P"
                                        Exit For
                               End If
                               If ((p.Offset(0, 5).Value <> "N") And (i = 6)) Then  'per la domenica si controlla che 1 turni prima della settimana successiva non sia N , per evitare di cambiare MNS
                                       ' p.Select
                                        trovatoMNellaParteBassa = 1
                                        p.Value = "P"
                                        Exit For  ' sostituiamo un solo M alla volta
                               End If
                            Next
                            If (trovatoMNellaParteBassa = 0) Then  '' non si sono trovati P da sostituire nella parte bassa allora si cercano nella parte alta
                             Set Rslt1 = Range(FindAll("M", Application.Intersect(ActiveSheet.UsedRange, Range(c1, c3)), xlFormulas, xlPart, SearchFormat:=True).Address)
                             For Each p In Rslt1
                                 If ((p.Offset(0, 1).Value <> "N")) Then
                                      ' p.Select
                                       
                                        p.Value = "P"
                                        Exit For
                               End If
                               If ((p.Offset(0, 5).Value <> "N") And (i = 6)) Then  'per la domenica si controlla che 1 turni prima della settimana successiva non sia N , per evitare di cambiare MNS
                                        'p.Select
                                       
                                        p.Value = "P"
                                        Exit For  ' sostituiamo un solo M alla volta
                               End If
                            Next
                            End If
                            
                        End If
                        
                       
                      nPomeriggio = Application.CountIf(Range(c1, c2), "P")
                Wend  '''  mentre le P sono > di 9
                
                
                
                
                
                
                
    
            Next ''  per ogni giorno della settimana
     Next '' per ogni lenedi trovato , quindi per ogni settimana
      Range("A1").Select
End Sub
