VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf2 
   Caption         =   "Australia QGC Automation Tool Run 2: Clearing"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16560
   OleObjectBlob   =   "QGC_UF2_Clearing & Trend.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'------------------------------------
'Developed By
'Nicholas Tay (MYNTAQ)
'Nicholas.N.Tay@shell.com
'https://www.linkedin.com/in/nicholastayyauyung/
'------------------------------------



'**************************************************************************************************************
'Declaration of module variables
'**************************************************************************************************************

Private StartTime   As Double, SecondsElapsed As Double, MinutesElapsed As String
Private inputError  As String

Private wsSettings As Worksheet
Private wsRef As Worksheet, wsProf As Worksheet

Private myFile_Clearing As String, myFile_GST As String

Private myFolderNames(0 To 10) As String

Private myPath As String, myPath_Prev As String


Private wbClearing As Workbook
Private wsC_Data As Worksheet, wsC_Control As Worksheet, wsC_Check As Worksheet, wsC_Check2 As Worksheet, wsC_Input As Worksheet, wsC_Output As Worksheet





'**************************************************************************************************************
'Main Routine
'**************************************************************************************************************


Private Sub BtnRun_Click()
    StartTime = Timer
    inputError = ""
    
    checkFields
    
    If inputError <> "" Then
        MsgBox ("Please ensure the following fields are filled before proceeding:" & inputError)
        Exit Sub
    End If
    
    
    RunPauseAll
    
    openFiles
    
    'If Me.cbTrend.Value = True Then
    If Me.cbGST_BGIA.Value = True Or Me.cbGST_QCLNG.Value = True Or Me.cbGST_QGC.Value = True Or Me.cbGST_Single.Value = True Then
        doTrend
    End If
    
    'If Me.cbClearing.Value = True Then
    If Me.cbClearing_BGIA.Value = True Or Me.cbClearing_QCLNG.Value = True Or Me.cbClearing_QGC.Value = True Or Me.cbClearing_Single.Value = True Then
        doClearing
    End If
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    
    RunActivateAll
    
    MsgBox ("Completed. Time taken: " & MinutesElapsed)
    Unload Me
    
End Sub



'**************************************************************************************************************
'Requirement Runs
'**************************************************************************************************************


Private Sub UserForm_Initialize()
    'ShowTitleBar Me
    Set wsSettings = ThisWorkbook.Worksheets("Settings")
    Set wsRef = ThisWorkbook.Worksheets("Ref")
    Set wsProf = ThisWorkbook.Worksheets("ProfitMatch")
    
    If Not Environ("username") = wsSettings.Range("B1").Value Then
        wsSettings.Range("B:B").EntireColumn.Clear
        wsSettings.Range("B1").Value = Environ("username")
    Else
        Me.tbSaveLocation.Value = wsSettings.Range("B6").Value
        Me.tbClearing.Value = wsSettings.Range("B7").Value
    End If
    
    
End Sub


Private Sub BtnCancel_Click()
    Unload Me
End Sub



Private Sub btnClearing_Click()
    Me.tbClearing.Value = SearchFileLocation
    wsSettings.Range("B7").Value = Me.tbClearing.Value
End Sub


Private Sub btnSaveLocation_Click()
    Me.tbSaveLocation.Value = SearchFolderLocation
    wsSettings.Range("B6").Value = Me.tbSaveLocation.Value
End Sub


'**************************************************************************************************************
'Processing Runs
'**************************************************************************************************************



Sub checkFields()
    
    If Me.tbSaveLocation.Value = "" Then
        inputError = inputError & vbLf & " - Save Location (1)"
    ElseIf Dir(Me.tbSaveLocation.Value, vbDirectory) = "" Then
        inputError = inputError & vbLf & " - Save Location (2)"
    End If
    
    If Me.cbGST_BGIA.Value = False And Me.cbGST_QCLNG.Value = False And Me.cbGST_QGC.Value = False And Me.cbGST_Single.Value = False And Me.cbClearing_BGIA.Value = False And Me.cbClearing_QCLNG.Value = False And Me.cbClearing_QGC.Value = False And Me.cbClearing_Single.Value = False Then
        inputError = inputError & vbLf & " - No action selected"
        
    ElseIf Me.cbClearing_BGIA.Value = True Or Me.cbClearing_QCLNG.Value = True Or Me.cbClearing_QGC.Value = True Or Me.cbClearing_Single.Value = True Then
        If Me.tbClearing.Value = "" Then
            inputError = inputError & vbLf & " - Clearing Form Template (1)"
        ElseIf Dir(Me.tbClearing.Value, vbNormal) = "" Then
            inputError = inputError & vbLf & " - Clearing Form Template (2)"
        End If
    End If
    
    
End Sub


Sub openFiles()
    Dim xCount As Long
    
    If Dir(Me.tbSaveLocation.Value, vbDirectory) = "" Then
        MsgBox ("Source folder not found. Please ensure the correct file path is selected.")
        End
    End If
    
    myFile_Clearing = ""
    myFile_GST = ""
    
    
    If Right(Me.tbSaveLocation.Value, 12) = Format(DateAdd("m", -1, Date), "yyyy") & "\" & Format(DateAdd("m", -1, Date), "mm mmm") & "\" Then
        myPath = Me.tbSaveLocation.Value
        
        If Dir(Left(Me.tbSaveLocation.Value, Len(Me.tbSaveLocation.Value) - 12) & Format(DateAdd("m", -2, Date), "yyyy") & "\" & Format(DateAdd("m", -2, Date), "mm mmm") & "\", vbDirectory) = "" Then
            MsgBox ("Previous month source file not found. Please ensure the correct file path is selected.")
            End
        Else
            myPath_Prev = Left(Me.tbSaveLocation.Value, Len(Me.tbSaveLocation.Value) - 12) & Format(DateAdd("m", -2, Date), "yyyy") & "\" & Format(DateAdd("m", -2, Date), "mm mmm") & "\"
        End If
        
    ElseIf Right(Me.tbSaveLocation.Value, 5) = Format(DateAdd("m", -1, Date), "yyyy") & "\" Then
        If Dir(Me.tbSaveLocation.Value & Format(DateAdd("m", -1, Date), "mm mmm") & "\", vbDirectory) = "" Then
            MsgBox ("Source folder for current month not found. Please ensure the correct file path is selected.")
            End
        Else
            myPath = Me.tbSaveLocation.Value & Format(DateAdd("m", -1, Date), "mm mmm") & "\"
            If Dir(Left(Me.tbSaveLocation.Value, Len(Me.tbSaveLocation.Value) - 5) & Format(DateAdd("m", -2, Date), "yyyy") & "\" & Format(DateAdd("m", -2, Date), "mm mmm") & "\", vbDirectory) = "" Then
                MsgBox ("Previous month source file not found. Please ensure the correct file path is selected.")
                End
            Else
                myPath_Prev = Left(Me.tbSaveLocation.Value, Len(Me.tbSaveLocation.Value) - 5) & Format(DateAdd("m", -2, Date), "yyyy") & "\" & Format(DateAdd("m", -2, Date), "mm mmm") & "\"
            End If
        End If
        
    ElseIf Not Dir(Me.tbSaveLocation.Value & Format(DateAdd("m", -1, Date), "yyyy") & "\" & Format(DateAdd("m", -1, Date), "mm mmm") & "\", vbDirectory) = "" Then
        myPath = Me.tbSaveLocation.Value & Format(DateAdd("m", -1, Date), "yyyy") & "\" & Format(DateAdd("m", -1, Date), "mm mmm") & "\"
        
        If Dir(Me.tbSaveLocation.Value & Format(DateAdd("m", -2, Date), "yyyy") & "\" & Format(DateAdd("m", -2, Date), "mm mmm") & "\", vbDirectory) = "" Then
            MsgBox ("Previous month source file not found. Please ensure the correct file path is selected.")
            End
        Else
            myPath_Prev = Me.tbSaveLocation.Value & Format(DateAdd("m", -2, Date), "yyyy") & "\" & Format(DateAdd("m", -2, Date), "mm mmm") & "\"
        End If
        
    Else
        MsgBox ("Source folder not found. Please ensure the correct file path is selected.")
        End
    End If
    
    
    myFolderNames(0) = "1100 - BGIA (QGC Upstream)\"
    myFolderNames(1) = "5000 - QGC Group\"
    myFolderNames(2) = "5000 - QGC JV\"
    myFolderNames(3) = "5030 - Toll Co 2\"
    myFolderNames(4) = "5031 - Toll Co 2 (2)\"
    myFolderNames(5) = "5033 - Toll Co 1\"
    myFolderNames(6) = "5036 - QCLNG (OpCo)\"
    myFolderNames(7) = "5037 - Train 1\"
    myFolderNames(8) = "5038 - Train 2\"
    myFolderNames(9) = "5045 - T1 UJV\"
    myFolderNames(10) = "5046 - T2 UJV\"
    
    
    For xCount = 0 To 10
        If Dir(myPath & myFolderNames(xCount), vbDirectory) = "" Then
            MsgBox ("Folder " & myFolderNames(xCount) & " is missing for the current month. Please ensure the correct file path is selected.")
            End
        ElseIf Dir(myPath_Prev & myFolderNames(xCount), vbDirectory) = "" Then
            MsgBox ("Folder " & myFolderNames(xCount) & " is missing for the previous month. Please ensure the correct file path is selected.")
            End
        End If
    Next
    
    If Me.cbClearing_BGIA.Value = True Or Me.cbClearing_QCLNG.Value = True Or Me.cbClearing_QGC.Value = True Or Me.cbClearing_Single.Value = True Then
    'If Me.cbClearing.Value = True Then
        Set wbClearing = Workbooks.Open(Me.tbClearing.Value, ReadOnly:=True)
        
        On Error GoTo errorClearing
        Set wsC_Control = wbClearing.Worksheets("Control")
        Set wsC_Data = wbClearing.Worksheets("Data")
        Set wsC_Check = wbClearing.Worksheets("Check")
        Set wsC_Check2 = wbClearing.Worksheets("Check2")
        Set wsC_Input = wbClearing.Worksheets("Input")
        Set wsC_Output = wbClearing.Worksheets("Output")
        On Error GoTo 0
    End If
    
    Exit Sub
    
    
errorClearing:
    wbClearing.Close False
    MsgBox ("Error found in clearing file. Please ensure the correct file is selected.")
    End
    
End Sub






Sub doTrend()
    Dim xCount As Long
    Dim curFileTrend As String, curFileInput As String, tempFile As String, wsName As String
    Dim wbTrend As Workbook, wbInput As Workbook, wsTrend As Worksheet, wsInput As Worksheet, ws As Worksheet
    Dim myRow As Long
    
    
    myFolderNames(0) = "1100 - BGIA (QGC Upstream)\"
    myFolderNames(1) = "5000 - QGC Group\"
    myFolderNames(2) = "5000 - QGC JV\"
    myFolderNames(3) = "5030 - Toll Co 2\"
    myFolderNames(4) = "5031 - Toll Co 2 (2)\"
    myFolderNames(5) = "5033 - Toll Co 1\"
    myFolderNames(6) = "5036 - QCLNG (OpCo)\"
    myFolderNames(7) = "5037 - Train 1\"
    myFolderNames(8) = "5038 - Train 2\"
    myFolderNames(9) = "5045 - T1 UJV\"
    myFolderNames(10) = "5046 - T2 UJV\"
    
    
    For xCount = 0 To 10
        If xCount = 2 Then
            GoTo nextItem
        ElseIf xCount = 0 Then
            If Me.cbGST_BGIA.Value = False Then GoTo nextItem
        ElseIf xCount = 1 Then
            If Me.cbGST_QGC.Value = False Then GoTo nextItem
        ElseIf xCount = 6 Then
            If Me.cbGST_QCLNG.Value = False Then GoTo nextItem
        Else
            If Me.cbGST_Single.Value = False Then GoTo nextItem
        End If
    
        Select Case xCount
        Case 0
            curFileInput = "BGIA Group Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            wsName = "Data Entry - BGIA Group"
        Case 1
            curFileInput = "QGC Group Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            wsName = "Data Entry - QGC Group"
        Case 3
            curFileInput = "5030 GST Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            wsName = "Data Entry - 5030"
        Case 4
            curFileInput = "5031 GST Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            wsName = "Data Entry - 5031"
        Case 5
            curFileInput = "5033 GST Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            wsName = "Data Entry - 5033"
        Case 6
            curFileInput = "QCLNG Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            wsName = "Data Entry - QCLNG Group"
        Case 7
            curFileInput = "5037 GST Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            wsName = "Data Entry - 5037"
        Case 8
            curFileInput = "5038 GST Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            wsName = "Data Entry - 5038"
        Case 9
            curFileInput = "5045 GST Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            wsName = "Data Entry - 5045"
        Case 10
            curFileInput = "5046 GST Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            wsName = "Data Entry - 5046"
        End Select
        
        tempFile = Dir(myPath_Prev & myFolderNames(xCount), vbDirectory)
        Do While Not tempFile = ""
            If InStr(1, tempFile, "GST Trend Analysis", vbTextCompare) <> 0 Then
                curFileTrend = tempFile
                Exit Do
            End If
            tempFile = Dir
        Loop
        
        If Not (curFileTrend = "" Or Dir(myPath & myFolderNames(xCount) & curFileInput) = "") Then
            Set wbTrend = Workbooks.Open(myPath_Prev & myFolderNames(xCount) & curFileTrend, ReadOnly:=True)
            Set wsTrend = wbTrend.Worksheets("Summary")
            Set wbInput = Workbooks.Open(myPath & myFolderNames(xCount) & curFileInput, ReadOnly:=True)
            For Each ws In wbInput.Worksheets
                If ws.Name = wsName Then
                    Set wsInput = ws
                    Exit For
                End If
            Next
            If wsInput Is Nothing Then
                wbTrend.Close False
                wbInput.Close False
                inputError = inputError & vbLf & " - Input Form issue on " & Left(myFolderNames(xCount), Len(myFolderNames(xCount)) - 1)
                GoTo nothingHere
            End If
            
            wsTrend.Range("A1:L25").Interior.ColorIndex = 0
            
            If Month(DateAdd("m", -1, Date)) = 1 Then
                wsTrend.Range("A14:K25").Copy
                wsTrend.Range("A2").PasteSpecial (xlPasteValues)
                wsTrend.Range("A14:K25").Value = 0
                
                wsTrend.Range("A14").Value = DateSerial(Year(DateAdd("m", -1, Date)), Month(DateAdd("m", -1, Date)), 1)
                wsTrend.Range("A15").Value = DateSerial(Year(Date), Month(Date), 1)
                wsTrend.Range("A16").Value = DateSerial(Year(DateAdd("m", 1, Date)), Month(DateAdd("m", 1, Date)), 1)
                wsTrend.Range("A17").Value = DateSerial(Year(DateAdd("m", 2, Date)), Month(DateAdd("m", 2, Date)), 1)
                wsTrend.Range("A18").Value = DateSerial(Year(DateAdd("m", 3, Date)), Month(DateAdd("m", 3, Date)), 1)
                wsTrend.Range("A19").Value = DateSerial(Year(DateAdd("m", 4, Date)), Month(DateAdd("m", 4, Date)), 1)
                wsTrend.Range("A20").Value = DateSerial(Year(DateAdd("m", 5, Date)), Month(DateAdd("m", 5, Date)), 1)
                wsTrend.Range("A21").Value = DateSerial(Year(DateAdd("m", 6, Date)), Month(DateAdd("m", 6, Date)), 1)
                wsTrend.Range("A22").Value = DateSerial(Year(DateAdd("m", 7, Date)), Month(DateAdd("m", 7, Date)), 1)
                wsTrend.Range("A23").Value = DateSerial(Year(DateAdd("m", 8, Date)), Month(DateAdd("m", 8, Date)), 1)
                wsTrend.Range("A24").Value = DateSerial(Year(DateAdd("m", 9, Date)), Month(DateAdd("m", 9, Date)), 1)
                wsTrend.Range("A25").Value = DateSerial(Year(DateAdd("m", 10, Date)), Month(DateAdd("m", 10, Date)), 1)
                
                wsTrend.Range("A2:A25").NumberFormat = "mmm-yy"
                
            End If
            
            wsTrend.Range("B" & Month(DateAdd("m", -1, Date)) + 13 & ":L" & Month(DateAdd("m", -1, Date)) + 13).Interior.Color = RGB(221, 235, 247)
            wsTrend.Range("B" & Month(DateAdd("m", -1, Date)) + 13).Value = -Round(wsInput.Range("J9").Value / 1000000, 2)
            wsTrend.Range("C" & Month(DateAdd("m", -1, Date)) + 13).Value = -Round(wsInput.Range("J10").Value / 1000000, 2)
            wsTrend.Range("D" & Month(DateAdd("m", -1, Date)) + 13).Value = -Round(wsInput.Range("J11").Value / 1000000, 2)
            wsTrend.Range("E" & Month(DateAdd("m", -1, Date)) + 13).Value = -Round(wsInput.Range("J12").Value / 1000000, 2)
            wsTrend.Range("F" & Month(DateAdd("m", -1, Date)) + 13).Value = -Round(wsInput.Range("J14").Value / 1000000, 2)
            wsTrend.Range("G" & Month(DateAdd("m", -1, Date)) + 13).Value = -Round(wsInput.Range("J38").Value / 1000000, 2)
            wsTrend.Range("H" & Month(DateAdd("m", -1, Date)) + 13).Value = Round(wsInput.Range("J19").Value / 1000000, 2)
            wsTrend.Range("I" & Month(DateAdd("m", -1, Date)) + 13).Value = Round(wsInput.Range("J20").Value / 1000000, 2)
            wsTrend.Range("J" & Month(DateAdd("m", -1, Date)) + 13).Value = Round(wsInput.Range("J23").Value / 1000000, 2)
            wsTrend.Range("K" & Month(DateAdd("m", -1, Date)) + 13).Value = Round(wsInput.Range("J26").Value / 1000000, 2)
            
            wbTrend.SaveAs myPath & myFolderNames(xCount) & curFileTrend
            wbTrend.Close True
            wbInput.Close False
            
        End If
        
nothingHere:
nextItem:
    
        Set wsInput = Nothing
        curFileInput = ""
        curFileTrend = ""
        
    Next
    
End Sub


Sub doClearing()
    Dim xRow As Long, xLastRow As Long, rowIn As Long, rowOut As Long, curRow As Long, tempRow As Long, xCount As Long
    Dim curPath As String, myRef As String, tempName As String, myName As String
    
    
    Dim wbInput As Workbook, wsInput As Worksheet
    Dim wbOutput As Workbook, wsOutput As Worksheet
    
    Dim wbTemp As Workbook, wsTemp As Worksheet
    
    
    wsC_Control.Range("C2").Value = "'" & Format(DateAdd("m", -1, Date), "mmm yy")
    
    For xCount = 0 To 10
        If xCount = 2 Then
            GoTo skipClearingThis
        ElseIf xCount = 0 Then
            If Me.cbClearing_BGIA.Value = False Then GoTo skipClearingThis
            
        ElseIf xCount = 1 Then
            If Me.cbClearing_QGC.Value = False Then GoTo skipClearingThis
            
        ElseIf xCount = 6 Then
            If Me.cbClearing_QCLNG.Value = False Then GoTo skipClearingThis
            
        Else
            If Me.cbClearing_Single.Value = False Then GoTo skipClearingThis
            
        End If
        
        xLastRow = Application.WorksheetFunction.Max(wsC_Data.Cells(wsC_Data.Rows.Count, "B").End(xlUp).Row, wsC_Data.Cells(wsC_Data.Rows.Count, "J").End(xlUp).Row, wsC_Data.Cells(wsC_Data.Rows.Count, "L").End(xlUp).Row, wsC_Data.Cells(wsC_Data.Rows.Count, "T").End(xlUp).Row)
        If xLastRow >= 4 Then
            wsC_Data.Range("A4:A" & xLastRow + 10).EntireRow.Delete
        End If
        
        wsC_Check2.Cells.Delete
        
        
        wsC_Data.Range("M3").Value = "Amount in LC (USD)"
        
        Select Case xCount
        Case 0
            tempName = "BGIA"
            wsC_Data.Range("M3").Value = "Amount in LC (AUD)"
        Case 1
            tempName = "QGC"
            wsC_Data.Range("M3").Value = "Amount in LC (AUD)"
        Case 2
            tempName = "QGC JV"
        Case 3
            tempName = "5030"
        Case 4
            tempName = "5031"
        Case 5
            tempName = "5033"
        Case 6
            tempName = "QCLNG"
        Case 7
            tempName = "5037"
        Case 8
            tempName = "5038"
        Case 9
            tempName = "5045"
        Case 10
            tempName = "5046"
        End Select
        
        
        If xCount = 0 Then
            Set wbTemp = Workbooks.Add
            Set wsTemp = wbTemp.Worksheets(1)
            wsC_Data.Copy before:=wsTemp
            wsTemp.Delete
            Set wsTemp = wbTemp.Worksheets(1)
            
            
            
            xRow = 1
            myName = "BGIA Group Cross Co Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & ".xlsm"
            While Dir(myPath & myFolderNames(xCount) & myName) <> ""
                myName = "BGIA Group Cross Co Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & " (" & xRow & ").xlsm"
                xRow = xRow + 1
            Wend
            
            wbTemp.SaveAs myPath & myFolderNames(xCount) & myName, FileFormat:=52
            
            xLastRow = Application.WorksheetFunction.Max(wsTemp.Cells(wsTemp.Rows.Count, "B").End(xlUp).Row, wsTemp.Cells(wsTemp.Rows.Count, "J").End(xlUp).Row, wsTemp.Cells(wsTemp.Rows.Count, "L").End(xlUp).Row, wsTemp.Cells(wsTemp.Rows.Count, "T").End(xlUp).Row)
            If xLastRow < 4 Then
                xLastRow = 4
            End If
            wsTemp.Range("A4:A" & xLastRow + 10).EntireRow.Delete
            
            
            xRow = 1
            myName = "BGIA 1100 Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & ".xlsm"
            While Dir(myPath & myFolderNames(xCount) & myName) <> ""
                myName = "BGIA 1100 Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & " (" & xRow & ").xlsm"
                xRow = xRow + 1
            Wend
            
            wbClearing.SaveAs myPath & myFolderNames(xCount) & myName
            
            
            wsC_Data.Range("O3").Value = "Amount in LC (GBP)"
            
            Call doInputOutput(1, myPath & myFolderNames(xCount))
            Call doDataCheck
            
            Call doCrossCo(wsTemp, 1)
            
            
            xRow = 1
            myName = "BGIA 1106 1122 Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & ".xlsm"
            While Dir(myPath & myFolderNames(xCount) & myName) <> ""
                myName = "BGIA 1106 1122 Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & " (" & xRow & ").xlsm"
                xRow = xRow + 1
            Wend
            
            wbClearing.SaveAs myPath & myFolderNames(xCount) & myName
            
            
            wsC_Data.Range("O3").Value = "Amount in LC (AUD)"
            
            Call doInputOutput(2, myPath & myFolderNames(xCount))
            Call doDataCheck
            
            Call doCrossCo(wsTemp, 2)
            
            wbTemp.Save
            
            
        ElseIf xCount = 6 Then
            xRow = 1
            myName = "QCLNG Group Cross Co Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & ".xlsm"
            While Dir(myPath & myFolderNames(xCount) & myName) <> ""
                myName = "QCLNG Group Cross Co Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & " (" & xRow & ").xlsm"
                xRow = xRow + 1
            Wend
            
            wbTemp.SaveAs myPath & myFolderNames(xCount) & myName
            
            xLastRow = Application.WorksheetFunction.Max(wsTemp.Cells(wsTemp.Rows.Count, "B").End(xlUp).Row, wsTemp.Cells(wsTemp.Rows.Count, "J").End(xlUp).Row, wsTemp.Cells(wsTemp.Rows.Count, "L").End(xlUp).Row, wsTemp.Cells(wsTemp.Rows.Count, "T").End(xlUp).Row)
            If xLastRow < 4 Then
                xLastRow = 4
            End If
            wsTemp.Range("A4:A" & xLastRow + 10).EntireRow.Delete
            
            xRow = 1
            myName = tempName & " Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & ".xlsm"
            While Dir(myPath & myFolderNames(xCount) & myName) <> ""
                myName = tempName & " Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & " (" & xRow & ").xlsm"
                xRow = xRow + 1
            Wend
            
            wbClearing.SaveAs myPath & myFolderNames(xCount) & myName
            
            Call doInputOutput(3, myPath & myFolderNames(xCount))
            Call doDataCheck
            
            
            'do the cross co thingy
            
            Call doCrossCo(wsTemp, 3)
            
            wbTemp.Save
            wbTemp.Close True
            Set wbTemp = Nothing
            Set wsTemp = Nothing
            
        Else
            xRow = 1
            myName = tempName & " Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & ".xlsm"
            While Dir(myPath & myFolderNames(xCount) & myName) <> ""
                myName = tempName & " Clearing Journal " & Format(DateAdd("m", -1, Date), "MMM YYYY") & " (" & xRow & ").xlsm"
                xRow = xRow + 1
            Wend
            
            wbClearing.SaveAs myPath & myFolderNames(xCount) & myName
            
            Call doInputOutput(4, myPath & myFolderNames(xCount))
            Call doDataCheck
            
        End If
        
skipClearingThis:
        
    Next
    
    wbClearing.Close True
    
End Sub



Sub doCrossCo(ws As Worksheet, myType As Long)
    Dim xRow As Long, xLastRow As Long, xLastRowA As Long
    Dim myRow As Long
    
    
            
            
    
    With ws
        If myType = 1 Then
            'BGIA
            
            .Range("B4").Value = Format(Date, "dd.mm.yyyy")
            .Range("B6").Value = Format(Date, "dd.mm.yyyy")
            .Range("B8").Value = Format(Date, "dd.mm.yyyy")
            
            .Range("C4").Value = Format(Date, "dd.mm.yyyy")
            .Range("C6").Value = Format(Date, "dd.mm.yyyy")
            .Range("C8").Value = Format(Date, "dd.mm.yyyy")
            
            .Range("D4").Value = "ST"
            .Range("D6").Value = "ST"
            .Range("D8").Value = "ST"
            
            .Range("E4").Value = 1100
            .Range("E6").Value = 1106
            .Range("E8").Value = 1122
            
            .Range("F4").Value = "AUD"
            .Range("F6").Value = "AUD"
            .Range("F8").Value = "AUD"
            
            .Range("H4").Value = "1100 Cross Co " & Format(DateAdd("m", -1, Date), "MMM YY") & " AUD Clearing"
            .Range("H6").Value = "1106 Cross Co " & Format(DateAdd("m", -1, Date), "MMM YY") & " AUD Clearing"
            .Range("H8").Value = "1122 Cross Co " & Format(DateAdd("m", -1, Date), "MMM YY") & " AUD Clearing"
            
            .Range("I4").Formula = "=H4"
            .Range("I6").Formula = "=H6"
            .Range("I8").Formula = "=H8"
            
            .Range("K4:K9").Value = "20610000"
            
            .Range("P4").Value = 1100
            .Range("P5").Value = 5000
            .Range("P6").Value = 1106
            .Range("P7").Value = 5000
            .Range("P8").Value = 1122
            .Range("P9").Value = 5000
            
            .Range("T4").Value = 30991
            .Range("T5").Value = 31512
            .Range("T6").Value = 31186
            .Range("T7").Value = 31512
            .Range("T8").Value = 31540
            .Range("T9").Value = 31512
            
            .Range("W4").Value = 115000
            .Range("W5").Value = 111100
            .Range("W6").Value = 115000
            .Range("W7").Value = 111106
            .Range("W8").Value = 115000
            .Range("W9").Value = 111122
            
            .Range("X4").Value = 31512
            .Range("X5").Value = 30991
            .Range("X6").Value = 31512
            .Range("X7").Value = 31886
            .Range("X8").Value = 31512
            .Range("X9").Value = 31540
            
            .Range("AB4:AB9").Value = "915"
            
            .Range("AD4").Formula = "=H4"
            .Range("AD5").Formula = "=H4"
            .Range("AD6").Formula = "=H6"
            .Range("AD7").Formula = "=H6"
            .Range("AD8").Formula = "=H8"
            .Range("AD9").Formula = "=H8"
            
            .Range("L4:O4").Interior.Color = RGB(255, 255, 0)
            .Range("L6:O6").Interior.Color = RGB(255, 255, 0)
            .Range("L8:O8").Interior.Color = RGB(255, 255, 0)
            
            .Range("B5:AI5").Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range("B5:AI5").Borders(xlEdgeBottom).Weight = xlThin
            .Range("B7:AI7").Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range("B7:AI7").Borders(xlEdgeBottom).Weight = xlThin
            .Range("B9:AI9").Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range("B9:AI9").Borders(xlEdgeBottom).Weight = xlThin
            
            .Range("M4").Value = "= 0"
            .Range("N4").Value = "= 0"
            .Range("O4").Value = "= 0"
                
            .Range("M6").Value = "= 0"
            .Range("N6").Value = "= 0"
            .Range("O6").Value = "= 0"
                
            .Range("M8").Value = "= 0"
            .Range("N8").Value = "= 0"
            .Range("O8").Value = "= 0"
            
            
            '1100 only
            
            xLastRow = wsC_Data.Cells(wsC_Data.Rows.Count, "J").End(xlUp).Row
            'wsC_Data.Range("B3:O" & xLastRow).CopyPicture
            '.Range("B12").PasteSpecial
            
            xLastRowA = wsC_Data.Cells(wsC_Data.Rows.Count, "H").End(xlUp).Row
            xRow = 4
            
            While Not wsC_Data.Range("E" & xRow).Value = ""
                myRow = xRow
                While Not wsC_Data.Range("K" & myRow).Value = 20610000
                    myRow = myRow + 1
                Wend
                
                If wsC_Data.Range("E" & xRow).Value = 1100 Then
                    If wsC_Data.Range("J" & myRow).Value = 40 Then
                        .Range("M4").Formula = .Range("M4").Formula & " - " & wsC_Data.Range("M" & myRow).Value
                        .Range("N4").Formula = .Range("N4").Formula & " - " & wsC_Data.Range("N" & myRow).Value
                        .Range("O4").Formula = .Range("O4").Formula & " - " & wsC_Data.Range("O" & myRow).Value
                    Else
                        .Range("M4").Formula = .Range("M4").Formula & " + " & wsC_Data.Range("M" & myRow).Value
                        .Range("N4").Formula = .Range("N4").Formula & " + " & wsC_Data.Range("N" & myRow).Value
                        .Range("O4").Formula = .Range("O4").Formula & " + " & wsC_Data.Range("O" & myRow).Value
                    End If
                End If
                
                xRow = myRow + 1
            Wend
            
        ElseIf myType = 2 Then
            '1106 and 1122
            
            
            xLastRow = wsC_Data.Cells(wsC_Data.Rows.Count, "J").End(xlUp).Row
            'wsC_Data.Range("B3:O" & xLastRow).CopyPicture
            '.Range("T12").PasteSpecial (xlPasteAll)
            
            xLastRowA = wsC_Data.Cells(wsC_Data.Rows.Count, "H").End(xlUp).Row
            xRow = 4
            
            While Not wsC_Data.Range("E" & xRow).Value = ""
                myRow = xRow
                While Not wsC_Data.Range("K" & myRow).Value = 20610000
                    myRow = myRow + 1
                Wend
                
                If wsC_Data.Range("E" & xRow).Value = 1106 Then
                    If wsC_Data.Range("J" & myRow).Value = 40 Then
                        .Range("M6").Formula = .Range("M6").Formula & " - " & wsC_Data.Range("M" & myRow).Value
                        .Range("N6").Formula = .Range("N6").Formula & " - " & wsC_Data.Range("N" & myRow).Value
                        .Range("O6").Formula = .Range("O6").Formula & " - " & wsC_Data.Range("O" & myRow).Value
                    Else
                        .Range("M6").Formula = .Range("M6").Formula & " + " & wsC_Data.Range("M" & myRow).Value
                        .Range("N6").Formula = .Range("N6").Formula & " + " & wsC_Data.Range("N" & myRow).Value
                        .Range("O6").Formula = .Range("O6").Formula & " + " & wsC_Data.Range("O" & myRow).Value
                    End If
                    
                ElseIf wsC_Data.Range("E" & xRow).Value = 1122 Then
                    If wsC_Data.Range("J" & myRow).Value = 40 Then
                        .Range("M8").Formula = .Range("M8").Formula & " - " & wsC_Data.Range("M" & myRow).Value
                        .Range("N8").Formula = .Range("N8").Formula & " - " & wsC_Data.Range("N" & myRow).Value
                        .Range("O8").Formula = .Range("O8").Formula & " - " & wsC_Data.Range("O" & myRow).Value
                    Else
                        .Range("M8").Formula = .Range("M8").Formula & " + " & wsC_Data.Range("M" & myRow).Value
                        .Range("N8").Formula = .Range("N8").Formula & " + " & wsC_Data.Range("N" & myRow).Value
                        .Range("O8").Formula = .Range("O8").Formula & " + " & wsC_Data.Range("O" & myRow).Value
                    End If
                End If
                
                xRow = myRow + 1
            Wend
            
            If .Range("M4").Value >= 0 Then
                .Range("J4").Value = 40
                .Range("J5").Value = 50
            Else
                .Range("M4").Formula = "=ABS(" & Right(.Range("M4").Formula, Len(.Range("M4").Formula) - 2) & ")"
                .Range("N4").Formula = "=ABS(" & Right(.Range("N4").Formula, Len(.Range("N4").Formula) - 2) & ")"
                .Range("O4").Formula = "=ABS(" & Right(.Range("O4").Formula, Len(.Range("O4").Formula) - 2) & ")"
                .Range("J4").Value = 50
                .Range("J5").Value = 40
            End If
            
            If .Range("M6").Value >= 0 Then
                .Range("J6").Value = 40
                .Range("J7").Value = 50
            Else
                .Range("M6").Formula = "=ABS(" & Right(.Range("M6").Formula, Len(.Range("M6").Formula) - 2) & ")"
                .Range("N6").Formula = "=ABS(" & Right(.Range("N6").Formula, Len(.Range("N6").Formula) - 2) & ")"
                .Range("O6").Formula = "=ABS(" & Right(.Range("O6").Formula, Len(.Range("O6").Formula) - 2) & ")"
                .Range("J6").Value = 50
                .Range("J7").Value = 40
            End If
            
            If .Range("M8").Value >= 0 Then
                .Range("J8").Value = 40
                .Range("J9").Value = 50
            Else
                .Range("M8").Formula = "=ABS(" & Right(.Range("M8").Formula, Len(.Range("M8").Formula) - 2) & ")"
                .Range("N8").Formula = "=ABS(" & Right(.Range("N8").Formula, Len(.Range("N8").Formula) - 2) & ")"
                .Range("O8").Formula = "=ABS(" & Right(.Range("O8").Formula, Len(.Range("O8").Formula) - 2) & ")"
                .Range("J8").Value = 50
                .Range("J9").Value = 40
            End If
            
            .Range("L4").Formula = "=M4"
            .Range("L5").Formula = "=L4"
            .Range("M5").Formula = "=M4"
            .Range("N5").Formula = "=N4"
            .Range("O5").Formula = "=O4"
            
            .Range("L6").Formula = "=M6"
            .Range("L7").Formula = "=L6"
            .Range("M7").Formula = "=M6"
            .Range("N7").Formula = "=N6"
            .Range("O7").Formula = "=O6"
            
            .Range("L8").Formula = "=M8"
            .Range("L9").Formula = "=L8"
            .Range("M9").Formula = "=M8"
            .Range("N9").Formula = "=N8"
            .Range("O9").Formula = "=O8"
            
            
        Else
            'QCLNG
            
            .Range("B4").Value = Format(Date, "dd.mm.yyyy")
            .Range("C4").Value = Format(Date, "dd.mm.yyyy")
            .Range("D4").Value = "ST"
            .Range("E4").Value = "5039"
            .Range("F4").Value = "AUD"
            
            .Range("H4").Value = "OPCO CrossCo " & Format(DateAdd("m", -1, Date), "MMM YY") & " AUD Clearing"
            .Range("I4").Formula = "=H4"
            
            .Range("J4:P4").Interior.Color = RGB(255, 255, 0)
            
            .Range("K4:K5").Value = "20610000"
            
            .Range("P4").Value = "5039"
            .Range("P5").Value = "5036"
            
            .Range("T4").Value = "31683"
            .Range("T5").Value = "31680"
            
            .Range("W4").Value = "115036"
            .Range("W5").Value = "115039"
            
            .Range("X4").Value = "31680"
            .Range("X5").Value = "31683"
            
            .Range("AB4").Value = "915"
            .Range("AB5").Value = "915"
            
            .Range("AD4").Formula = "=I4"
            .Range("AD5").Formula = "=I4"
            
            .Range("M4").Value = "= 0"
            .Range("N4").Value = "= 0"
            .Range("O4").Value = "= 0"
            
            
            xLastRow = wsC_Data.Cells(wsC_Data.Rows.Count, "J").End(xlUp).Row
            'wsC_Data.Range("B3:O" & xLastRow).CopyPicture
            '.Range("B8").PasteSpecial
            
            
            xLastRowA = wsC_Data.Cells(wsC_Data.Rows.Count, "H").End(xlUp).Row
            xRow = 4
            
            While Not wsC_Data.Range("E" & xRow).Value = ""
                myRow = xRow
                While Not wsC_Data.Range("K" & myRow).Value = 20610000
                    myRow = myRow + 1
                Wend
                
                If wsC_Data.Range("E" & xRow).Value = 5039 Then
                    If wsC_Data.Range("J" & myRow).Value = 40 Then
                        .Range("M4").Formula = .Range("M4").Formula & " - " & wsC_Data.Range("M" & myRow).Value
                        .Range("N4").Formula = .Range("N4").Formula & " - " & wsC_Data.Range("N" & myRow).Value
                        .Range("O4").Formula = .Range("O4").Formula & " - " & wsC_Data.Range("O" & myRow).Value
                    Else
                        .Range("M4").Formula = .Range("M4").Formula & " + " & wsC_Data.Range("M" & myRow).Value
                        .Range("N4").Formula = .Range("N4").Formula & " + " & wsC_Data.Range("N" & myRow).Value
                        .Range("O4").Formula = .Range("O4").Formula & " + " & wsC_Data.Range("O" & myRow).Value
                    End If
                End If
                
                xRow = myRow + 1
            Wend
            
            If .Range("M4").Value >= 0 Then
                .Range("J4").Value = 40
                .Range("J5").Value = 50
            Else
                .Range("M4").Formula = "=ABS(" & Right(.Range("M4").Formula, Len(.Range("M4").Formula) - 2) & ")"
                .Range("N4").Formula = "=ABS(" & Right(.Range("N4").Formula, Len(.Range("N4").Formula) - 2) & ")"
                .Range("O4").Formula = "=ABS(" & Right(.Range("O4").Formula, Len(.Range("O4").Formula) - 2) & ")"
                .Range("J4").Value = 50
                .Range("J5").Value = 40
            End If
            
            .Range("L4").Formula = "=M4"
            .Range("L5").Formula = "=L4"
            .Range("M5").Formula = "=M4"
            .Range("N5").Formula = "=N4"
            .Range("O5").Formula = "=O4"
            
            
        End If
    End With
    
End Sub






Sub doInputOutput(myType As Long, selPath As String)
    Dim curPath As String
    Dim xLastRow As Long, xRow As Long
    
    Dim myPvtCache As PivotCache
    Dim myRng As Range
    Dim rngStr As String
    
    wsC_Input.Cells.Delete
    wsC_Output.Cells.Delete
    
    Dim wbInput As Workbook, wsInput As Worksheet
    Dim wbOutput As Workbook, wsOutput As Worksheet
    
    
    curPath = Dir(selPath)
    Do While curPath <> ""
        If Right(curPath, 23) = "Input " & Format(DateAdd("m", -1, Date), "MMM YYYY") & " (1).xlsx" Then
            wsC_Input.Cells.Delete
            
            Set wbInput = Workbooks.Open(selPath & curPath, ReadOnly:=True)
            Set wsInput = wbInput.Worksheets(1)
            
            xLastRow = Application.WorksheetFunction.Max(wsInput.Cells(wsInput.Rows.Count, "H").End(xlUp).Row, wsInput.Cells(wsInput.Rows.Count, "I").End(xlUp).Row)
        
            While wsInput.Range("I" & xLastRow).Interior.Color = RGB(255, 255, 0)
                wsInput.Range("I" & xLastRow).EntireRow.Delete
                xLastRow = xLastRow - 1
            Wend
            
            xRow = 2
            While Not wsInput.Range("A" & xRow).Value = ""
                If InStr(1, wsInput.Range("Q" & xRow).Value, "Clearing", vbTextCompare) <> 0 Or InStr(1, wsInput.Range("Q" & xRow).Value, "Valuation", vbTextCompare) <> 0 Then
                    wsInput.Range("Q" & xRow).EntireRow.Delete
                    xRow = xRow - 1
                End If
                xRow = xRow + 1
            Wend
            
            
            xLastRow = Application.WorksheetFunction.Max(wsInput.Cells(wsInput.Rows.Count, "H").End(xlUp).Row, wsInput.Cells(wsInput.Rows.Count, "I").End(xlUp).Row)
        
            wsInput.Range("A1:Q" & xLastRow).Copy
            wsC_Input.Range("A1").PasteSpecial (xlPasteValues)
            
            wbInput.Close False
            Set wbInput = Nothing
            Set wsInput = Nothing
            
            
            xLastRow = wsC_Input.Cells(wsC_Input.Rows.Count, "A").End(xlUp).Row
            xRow = 2
            
            Select Case myType
            Case 1 'BGIA 1100
                While Not xRow > xLastRow
                    If Not wsC_Input.Range("A" & xRow).Value = "1100" Then
                        wsC_Input.Range("A" & xRow).EntireRow.Delete
                        xRow = xRow - 1
                    End If
                    
                    xRow = xRow + 1
                    xLastRow = wsC_Input.Cells(wsC_Input.Rows.Count, "A").End(xlUp).Row
                Wend
                
            Case 2 'BGIA 1106 1122
                While Not xRow > xLastRow
                    If Not (wsC_Input.Range("A" & xRow).Value = "1106" Or wsC_Input.Range("A" & xRow).Value = "1122") Then
                        wsC_Input.Range("A" & xRow).EntireRow.Delete
                        xRow = xRow - 1
                    End If
                    
                    xRow = xRow + 1
                    xLastRow = wsC_Input.Cells(wsC_Input.Rows.Count, "A").End(xlUp).Row
                Wend
                
            Case 3 'QCLNG
                
                'While Not xRow > xLastRow
                '    If Not wsC_Input.Range("A" & xRow).Value = "5039" Then
                '        wsC_Input.Range("A" & xRow).EntireRow.Delete
                '        xRow = xRow - 1
                '    End If
                    
                '    xRow = xRow + 1
                '    xLastRow = wsC_Input.Cells(wsC_Input.Rows.Count, "A").End(xlUp).Row
                'Wend
                
            End Select
            
            
            xLastRow = wsC_Input.Cells(wsC_Input.Rows.Count, "A").End(xlUp).Row
            
            If xLastRow >= 2 Then
                
                If wsC_Input.Range("R1").Value = "" Then
                    rngStr = wsC_Input.Name & "!" & wsC_Input.Range("A1:Q" & xLastRow).Address(ReferenceStyle:=xlR1C1)
                Else
                    rngStr = wsC_Input.Name & "!" & wsC_Input.Range("A1:R" & xLastRow).Address(ReferenceStyle:=xlR1C1)
                End If
                
                
            
                wbClearing.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                rngStr, Version:=6).CreatePivotTable TableDestination:= _
                "Input!R1C20", TableName:="PivotTable1", DefaultVersion:=6
                
                With wsC_Input.PivotTables("PivotTable1")
                    .ColumnGrand = True
                    .HasAutoFormat = True
                    .DisplayErrorString = False
                    .DisplayNullString = True
                    .EnableDrilldown = True
                    .ErrorString = ""
                    .MergeLabels = False
                    .NullString = ""
                    .PageFieldOrder = 2
                    .PageFieldWrapCount = 0
                    .PreserveFormatting = True
                    .RowGrand = True
                    .SaveData = True
                    .PrintTitles = False
                    .RepeatItemsOnEachPrintedPage = True
                    .TotalsAnnotation = False
                    .CompactRowIndent = 1
                    .InGridDropZones = False
                    .DisplayFieldCaptions = True
                    .DisplayMemberPropertyTooltips = False
                    .DisplayContextTooltips = True
                    .ShowDrillIndicators = True
                    .PrintDrillIndicators = False
                    .AllowMultipleFilters = False
                    .SortUsingCustomLists = True
                    .FieldListSortAscending = False
                    .ShowValuesRow = False
                    .CalculatedMembersInFilters = False
                    .RowAxisLayout xlCompactRow
                End With
                
                With wsC_Input.PivotTables("PivotTable1").PivotCache
                    .RefreshOnFileOpen = False
                    .MissingItemsLimit = xlMissingItemsDefault
                End With
                
                wsC_Input.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
                
                With wsC_Input.PivotTables("PivotTable1").PivotFields("Company Code")
                    .Orientation = xlRowField
                    .Position = 1
                End With
                
                With wsC_Input.PivotTables("PivotTable1").PivotFields("Document currency")
                    .Orientation = xlRowField
                    .Position = 2
                End With
                
                With wsC_Input.PivotTables("PivotTable1").PivotFields("Profit Center")
                    .Orientation = xlRowField
                    .Position = 3
                End With
                
                With wsC_Input.PivotTables("PivotTable1")
                    .AddDataField wsC_Input.PivotTables("PivotTable1").PivotFields("Amount in doc. curr."), "Sum of Amount in doc. curr.", xlSum
                    .AddDataField wsC_Input.PivotTables("PivotTable1").PivotFields("Amount in local currency"), "Sum of Amount in local currency", xlSum
                    .AddDataField wsC_Input.PivotTables("PivotTable1").PivotFields("Amount in loc.curr.2"), "Sum of Amount in loc.curr.2", xlSum
                    .AddDataField wsC_Input.PivotTables("PivotTable1").PivotFields("Amt in loc.curr. 3"), "Sum of Amt in loc.curr. 3", xlSum
                    
                    .PivotFields("Company Code").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Document currency").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Amount in doc. curr.").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Local Currency").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Amount in local currency").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Local currency 2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Amount in loc.curr.2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Local currency 3").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Amt in loc.curr. 3").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Document Number").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                        
                    .PivotFields("Document Type").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Posting Date").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Document Date").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Posting Key").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Tax code").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Profit Center").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Text").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .RowAxisLayout xlTabularRow
                    .RepeatAllLabels xlRepeatLabels
                    
                End With
                
                wbClearing.Save
            End If
            
        ElseIf Right(curPath, 24) = "Output " & Format(DateAdd("m", -1, Date), "MMM YYYY") & " (1).xlsx" Then
            wsC_Output.Cells.Delete
            
            Set wbOutput = Workbooks.Open(selPath & curPath, ReadOnly:=True)
            Set wsOutput = wbOutput.Worksheets(1)
            
            xLastRow = Application.WorksheetFunction.Max(wsOutput.Cells(wsOutput.Rows.Count, "H").End(xlUp).Row, wsOutput.Cells(wsOutput.Rows.Count, "I").End(xlUp).Row)
        
            While wsOutput.Range("I" & xLastRow).Interior.Color = RGB(255, 255, 0)
                wsOutput.Range("I" & xLastRow).EntireRow.Delete
                xLastRow = xLastRow - 1
            Wend
            
            
            xRow = 2
            While Not wsOutput.Range("A" & xRow).Value = ""
                If InStr(1, wsOutput.Range("Q" & xRow).Value, "Clearing", vbTextCompare) <> 0 Or InStr(1, wsOutput.Range("Q" & xRow).Value, "Valuation", vbTextCompare) <> 0 Then
                    wsOutput.Range("Q" & xRow).EntireRow.Delete
                    xRow = xRow - 1
                End If
                xRow = xRow + 1
            Wend
            
            
            xLastRow = Application.WorksheetFunction.Max(wsOutput.Cells(wsOutput.Rows.Count, "H").End(xlUp).Row, wsOutput.Cells(wsOutput.Rows.Count, "I").End(xlUp).Row)
        
            wsOutput.Range("A1:Q" & xLastRow).Copy
            wsC_Output.Range("A1").PasteSpecial (xlPasteValues)
            
            wbOutput.Close False
            Set wbOutput = Nothing
            Set wsOutput = Nothing
            
            xLastRow = wsC_Output.Cells(wsC_Output.Rows.Count, "A").End(xlUp).Row
            xRow = 2
            
            Select Case myType
            Case 1 'BGIA 1100
                While Not xRow > xLastRow
                    If Not wsC_Output.Range("A" & xRow).Value = "1100" Then
                        wsC_Output.Range("A" & xRow).EntireRow.Delete
                        xRow = xRow - 1
                    End If
                    
                    xRow = xRow + 1
                    xLastRow = wsC_Output.Cells(wsC_Output.Rows.Count, "A").End(xlUp).Row
                Wend
                
            Case 2 'BGIA 1106 1122
                While Not xRow > xLastRow
                    If Not (wsC_Output.Range("A" & xRow).Value = "1106" Or wsC_Output.Range("A" & xRow).Value = "1122") Then
                        wsC_Output.Range("A" & xRow).EntireRow.Delete
                        xRow = xRow - 1
                    End If
                    
                    xRow = xRow + 1
                    xLastRow = wsC_Output.Cells(wsC_Output.Rows.Count, "A").End(xlUp).Row
                Wend
                
            Case 3 'QCLNG
                'While Not xRow > xLastRow
                '    If Not wsC_Output.Range("A" & xRow).Value = "5039" Then
                '        wsC_Output.Range("A" & xRow).EntireRow.Delete
                '        xRow = xRow - 1
                '    End If
                '
                '    xRow = xRow + 1
                '    xLastRow = wsC_Output.Cells(wsC_Output.Rows.Count, "A").End(xlUp).Row
                'Wend
                '
            End Select
            
            xLastRow = wsC_Output.Cells(wsC_Output.Rows.Count, "A").End(xlUp).Row
            
            If xLastRow >= 2 Then
                
                If wsC_Output.Range("R1").Value = "" Then
                    rngStr = wsC_Output.Name & "!" & wsC_Output.Range("A1:Q" & xLastRow).Address(ReferenceStyle:=xlR1C1)
                Else
                    rngStr = wsC_Output.Name & "!" & wsC_Output.Range("A1:R" & xLastRow).Address(ReferenceStyle:=xlR1C1)
                End If
                
                
                wbClearing.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                rngStr, Version:=6).CreatePivotTable TableDestination:= _
                "Output!R1C20", TableName:="PivotTable2", DefaultVersion:=6
                
                With wsC_Output.PivotTables("PivotTable2")
                    .ColumnGrand = True
                    .HasAutoFormat = True
                    .DisplayErrorString = False
                    .DisplayNullString = True
                    .EnableDrilldown = True
                    .ErrorString = ""
                    .MergeLabels = False
                    .NullString = ""
                    .PageFieldOrder = 2
                    .PageFieldWrapCount = 0
                    .PreserveFormatting = True
                    .RowGrand = True
                    .SaveData = True
                    .PrintTitles = False
                    .RepeatItemsOnEachPrintedPage = True
                    .TotalsAnnotation = False
                    .CompactRowIndent = 1
                    .InGridDropZones = False
                    .DisplayFieldCaptions = True
                    .DisplayMemberPropertyTooltips = False
                    .DisplayContextTooltips = True
                    .ShowDrillIndicators = True
                    .PrintDrillIndicators = False
                    .AllowMultipleFilters = False
                    .SortUsingCustomLists = True
                    .FieldListSortAscending = False
                    .ShowValuesRow = False
                    .CalculatedMembersInFilters = False
                    .RowAxisLayout xlCompactRow
                End With
                
                With wsC_Output.PivotTables("PivotTable2").PivotCache
                    .RefreshOnFileOpen = False
                    .MissingItemsLimit = xlMissingItemsDefault
                End With
                
                wsC_Output.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
                
                With wsC_Output.PivotTables("PivotTable2").PivotFields("Company Code")
                    .Orientation = xlRowField
                    .Position = 1
                End With
                
                With wsC_Output.PivotTables("PivotTable2").PivotFields("Document currency")
                    .Orientation = xlRowField
                    .Position = 2
                End With
                
                With wsC_Output.PivotTables("PivotTable2").PivotFields("Profit Center")
                    .Orientation = xlRowField
                    .Position = 3
                End With
                
                With wsC_Output.PivotTables("PivotTable2")
                    .AddDataField wsC_Output.PivotTables("PivotTable2").PivotFields("Amount in doc. curr."), "Sum of Amount in doc. curr.", xlSum
                    .AddDataField wsC_Output.PivotTables("PivotTable2").PivotFields("Amount in local currency"), "Sum of Amount in local currency", xlSum
                    .AddDataField wsC_Output.PivotTables("PivotTable2").PivotFields("Amount in loc.curr.2"), "Sum of Amount in loc.curr.2", xlSum
                    .AddDataField wsC_Output.PivotTables("PivotTable2").PivotFields("Amt in loc.curr. 3"), "Sum of Amt in loc.curr. 3", xlSum
                    
                    .PivotFields("Company Code").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Document currency").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Amount in doc. curr.").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Local Currency").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Amount in local currency").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Local currency 2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Amount in loc.curr.2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Local currency 3").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Amt in loc.curr. 3").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Document Number").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                        
                    .PivotFields("Document Type").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Posting Date").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Document Date").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Posting Key").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Tax code").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Profit Center").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .PivotFields("Text").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                    .RowAxisLayout xlTabularRow
                    .RepeatAllLabels xlRepeatLabels
                    
                End With
                
                wbClearing.Save
            End If
        End If
        curPath = Dir
    Loop
    
    wbClearing.Save
    
End Sub


Sub doDataCheck()
    Dim xRow As Long, xLastRow As Long, rowIn As Long, rowOut As Long, tempRow As Long
    Dim myRef As String, tempStr As String
    
    Dim tempVal As Double
    
    Dim gotIn As Boolean, gotOut As Boolean
    
    Dim rngDocType As Range
    
    
    
    xRow = 4
    
    If wsC_Input.Range("T1").Value = "Company Code" Then
        rowIn = 2
    Else
        rowIn = 3
    End If
    
    If wsC_Output.Range("T1").Value = "Company Code" Then
        rowOut = 2
    Else
        rowOut = 3
    End If
    
    myRef = "" 'company code & document currency
    While wsC_Input.Range("T" & rowIn).Value <> "" Or wsC_Output.Range("T" & rowOut).Value <> ""
        gotIn = False
        gotOut = False
        
        If myRef = "" Then
            wsC_Data.Range("B" & xRow).Value = Format(Date, "dd.mm.yyyy")
            wsC_Data.Range("C" & xRow).Value = Format(Date, "dd.mm.yyyy")
            Set rngDocType = wsC_Data.Range("D" & xRow)
            
            If wsC_Input.Range("T" & rowIn).Value <> "" And wsC_Output.Range("T" & rowOut).Value <> "" Then
                If wsC_Input.Range("T" & rowIn).Value = wsC_Output.Range("T" & rowOut).Value Then
                    If wsC_Input.Range("U" & rowIn).Value = wsC_Output.Range("U" & rowOut).Value Or wsC_Input.Range("U" & rowIn).Value = "AUD" Then
                        myRef = wsC_Input.Range("T" & rowIn).Value & wsC_Input.Range("U" & rowIn).Value
                        wsC_Data.Range("E" & xRow).Value = wsC_Input.Range("T" & rowIn).Value
                        wsC_Data.Range("F" & xRow).Value = wsC_Input.Range("U" & rowIn).Value
                        
                        tempStr = wsC_Input.Range("V" & rowIn).Value 'profit center
                        
                    ElseIf wsC_Output.Range("U" & rowOut).Value = "AUD" Then
                        myRef = wsC_Output.Range("T" & rowOut).Value & wsC_Output.Range("U" & rowOut).Value
                        wsC_Data.Range("E" & xRow).Value = wsC_Output.Range("T" & rowOut).Value
                        wsC_Data.Range("F" & xRow).Value = wsC_Output.Range("U" & rowOut).Value
                        
                        tempStr = wsC_Output.Range("V" & rowOut).Value 'profit center
                        
                    ElseIf wsC_Input.Range("U" & rowIn).Value = "EUR" Then
                        myRef = wsC_Input.Range("T" & rowIn).Value & wsC_Input.Range("U" & rowIn).Value
                        wsC_Data.Range("E" & xRow).Value = wsC_Input.Range("T" & rowIn).Value
                        wsC_Data.Range("F" & xRow).Value = wsC_Input.Range("U" & rowIn).Value
                        
                        tempStr = wsC_Input.Range("V" & rowIn).Value 'profit center
                        
                    ElseIf wsC_Output.Range("U" & rowOut).Value = "EUR" Then
                        myRef = wsC_Output.Range("T" & rowOut).Value & wsC_Output.Range("U" & rowOut).Value
                        wsC_Data.Range("E" & xRow).Value = wsC_Output.Range("T" & rowOut).Value
                        wsC_Data.Range("F" & xRow).Value = wsC_Output.Range("U" & rowOut).Value
                        
                        tempStr = wsC_Output.Range("V" & rowOut).Value 'profit center
                        
                    ElseIf wsC_Input.Range("U" & rowIn).Value = "GBP" Then
                        myRef = wsC_Input.Range("T" & rowIn).Value & wsC_Input.Range("U" & rowIn).Value
                        wsC_Data.Range("E" & xRow).Value = wsC_Input.Range("T" & rowIn).Value
                        wsC_Data.Range("F" & xRow).Value = wsC_Input.Range("U" & rowIn).Value
                        
                        tempStr = wsC_Input.Range("V" & rowIn).Value 'profit center
                        
                    Else
                        myRef = wsC_Output.Range("T" & rowOut).Value & wsC_Output.Range("U" & rowOut).Value
                        wsC_Data.Range("E" & xRow).Value = wsC_Output.Range("T" & rowOut).Value
                        wsC_Data.Range("F" & xRow).Value = wsC_Output.Range("U" & rowOut).Value
                        
                        tempStr = wsC_Output.Range("V" & rowOut).Value 'profit center
                        
                    End If
                    
                ElseIf wsC_Input.Range("T" & rowIn).Value < wsC_Output.Range("T" & rowOut).Value Then
                    myRef = wsC_Input.Range("T" & rowIn).Value & wsC_Input.Range("U" & rowIn).Value
                    wsC_Data.Range("E" & xRow).Value = wsC_Input.Range("T" & rowIn).Value
                    wsC_Data.Range("F" & xRow).Value = wsC_Input.Range("U" & rowIn).Value
                    
                    tempStr = wsC_Input.Range("V" & rowIn).Value 'profit center
                ElseIf wsC_Input.Range("T" & rowIn).Value > wsC_Output.Range("T" & rowOut).Value Then
                    myRef = wsC_Output.Range("T" & rowOut).Value & wsC_Output.Range("U" & rowOut).Value
                    wsC_Data.Range("E" & xRow).Value = wsC_Output.Range("T" & rowOut).Value
                    wsC_Data.Range("F" & xRow).Value = wsC_Output.Range("U" & rowOut).Value
                    
                    tempStr = wsC_Output.Range("V" & rowOut).Value 'profit center
                End If
                
            ElseIf wsC_Input.Range("T" & rowIn).Value <> "" Then
                myRef = wsC_Input.Range("T" & rowIn).Value & wsC_Input.Range("U" & rowIn).Value
                wsC_Data.Range("E" & xRow).Value = wsC_Input.Range("T" & rowIn).Value
                wsC_Data.Range("F" & xRow).Value = wsC_Input.Range("U" & rowIn).Value
                
                tempStr = wsC_Input.Range("V" & rowIn).Value 'profit center
            ElseIf wsC_Output.Range("T" & rowOut).Value <> "" Then
                myRef = wsC_Output.Range("T" & rowOut).Value & wsC_Output.Range("U" & rowOut).Value
                wsC_Data.Range("E" & xRow).Value = wsC_Output.Range("T" & rowOut).Value
                wsC_Data.Range("F" & xRow).Value = wsC_Output.Range("U" & rowOut).Value
                
                tempStr = wsC_Output.Range("V" & rowOut).Value 'profit center
            End If
            
        
            wsC_Data.Range("H" & xRow).Formula = "=E" & xRow & " & "" "" & Control!$C$2 & "" "" & F" & xRow & " & Control!$C$3"
            'wsC_Data.Range("H4").Formula = "=E4 & "" "" & Control!$C$2 & "" "" & F4 & "" "" & Control!$C$3"
            wsC_Data.Range("I" & xRow).Formula = "=H" & xRow
            
        End If
    
        While myRef = wsC_Input.Range("T" & rowIn).Value & wsC_Input.Range("U" & rowIn).Value
            gotIn = True
            wsC_Data.Range("K" & xRow).Value = 20610101
            
            
            'If xRow >= 96 Then Stop
            
            tempVal = tempVal + wsC_Input.Range("W" & rowIn).Value
            
            If wsC_Input.Range("W" & rowIn).Value >= 0 Then
                wsC_Data.Range("J" & xRow).Value = 50
                wsC_Data.Range("L" & xRow).Value = wsC_Input.Range("W" & rowIn).Value
                wsC_Data.Range("L" & xRow).Font.Color = RGB(0, 0, 0)
            Else
                wsC_Data.Range("J" & xRow).Value = 40
                wsC_Data.Range("L" & xRow).Value = wsC_Input.Range("W" & rowIn).Value * -1
                wsC_Data.Range("L" & xRow).Font.Color = RGB(255, 0, 0)
                'tempVal = tempVal - wsC_Input.Range("W" & rowIn).Value
            End If
            
            If wsC_Input.Range("X" & rowIn).Value >= 0 Then
                wsC_Data.Range("M" & xRow).Value = wsC_Input.Range("X" & rowIn).Value
                wsC_Data.Range("M" & xRow).Font.Color = RGB(0, 0, 0)
            Else
                wsC_Data.Range("M" & xRow).Value = wsC_Input.Range("X" & rowIn).Value * -1
                wsC_Data.Range("M" & xRow).Font.Color = RGB(255, 0, 0)
            End If
            
            If wsC_Input.Range("Y" & rowIn).Value >= 0 Then
                wsC_Data.Range("N" & xRow).Value = wsC_Input.Range("Y" & rowIn).Value
                wsC_Data.Range("N" & xRow).Font.Color = RGB(0, 0, 0)
            Else
                wsC_Data.Range("N" & xRow).Value = wsC_Input.Range("Y" & rowIn).Value * -1
                wsC_Data.Range("N" & xRow).Font.Color = RGB(255, 0, 0)
            End If
            
            If wsC_Input.Range("Z" & rowIn).Value >= 0 Then
                wsC_Data.Range("O" & xRow).Value = wsC_Input.Range("Z" & rowIn).Value
                wsC_Data.Range("O" & xRow).Font.Color = RGB(0, 0, 0)
            Else
                wsC_Data.Range("O" & xRow).Value = wsC_Input.Range("Z" & rowIn).Value * -1
                wsC_Data.Range("O" & xRow).Font.Color = RGB(255, 0, 0)
            End If
            
            wsC_Data.Range("T" & xRow).Value = wsC_Input.Range("V" & rowIn).Value
            
            wsC_Data.Range("AB" & xRow).Value = 915
            wsC_Data.Range("AD" & xRow).Formula = "=$H$" & rngDocType.Row
            
            If tempStr <> wsC_Input.Range("V" & rowIn).Value Then
                If rngDocType.Value = "" Then
                    tempStr = ""
                    rngDocType.Value = "SF"
                End If
            End If
            
            
            rowIn = rowIn + 1
            xRow = xRow + 1
            
            
            While wsC_Input.Range("T" & rowIn).Value = "(blank)"
                rowIn = rowIn + 1
            Wend
            
            If wsC_Input.Range("T" & rowIn).Value = "Grand Total" Then
                rowIn = rowIn + 1
            End If
            
        Wend
        
        tempRow = xRow - 1
        
        While myRef = wsC_Output.Range("T" & rowOut).Value & wsC_Output.Range("U" & rowOut).Value
            gotOut = True
            wsC_Data.Range("K" & xRow).Value = 20610102
            
            
            tempVal = tempVal + wsC_Output.Range("W" & rowOut).Value
            
            If wsC_Output.Range("W" & rowOut).Value >= 0 Then
                wsC_Data.Range("J" & xRow).Value = 50
                wsC_Data.Range("L" & xRow).Value = wsC_Output.Range("W" & rowOut).Value
                wsC_Data.Range("L" & xRow).Font.Color = RGB(0, 0, 0)
            Else
                wsC_Data.Range("J" & xRow).Value = 40
                wsC_Data.Range("L" & xRow).Value = wsC_Output.Range("W" & rowOut).Value * -1
                wsC_Data.Range("L" & xRow).Font.Color = RGB(255, 0, 0)
            End If
            
            If wsC_Output.Range("X" & rowOut).Value >= 0 Then
                wsC_Data.Range("M" & xRow).Value = wsC_Output.Range("X" & rowOut).Value
                wsC_Data.Range("M" & xRow).Font.Color = RGB(0, 0, 0)
            Else
                wsC_Data.Range("M" & xRow).Value = wsC_Output.Range("X" & rowOut).Value * -1
                wsC_Data.Range("M" & xRow).Font.Color = RGB(255, 0, 0)
            End If
            
            If wsC_Output.Range("Y" & rowOut).Value >= 0 Then
                wsC_Data.Range("N" & xRow).Value = wsC_Output.Range("Y" & rowOut).Value
                wsC_Data.Range("N" & xRow).Font.Color = RGB(0, 0, 0)
            Else
                wsC_Data.Range("N" & xRow).Value = wsC_Output.Range("Y" & rowOut).Value * -1
                wsC_Data.Range("N" & xRow).Font.Color = RGB(255, 0, 0)
            End If
            
            If wsC_Output.Range("Z" & rowOut).Value >= 0 Then
                wsC_Data.Range("O" & xRow).Value = wsC_Output.Range("Z" & rowOut).Value
                wsC_Data.Range("O" & xRow).Font.Color = RGB(0, 0, 0)
            Else
                wsC_Data.Range("O" & xRow).Value = wsC_Output.Range("Z" & rowOut).Value * -1
                wsC_Data.Range("O" & xRow).Font.Color = RGB(255, 0, 0)
            End If
            
            wsC_Data.Range("T" & xRow).Value = wsC_Output.Range("V" & rowOut).Value
            
            wsC_Data.Range("AB" & xRow).Value = 915
            wsC_Data.Range("AD" & xRow).Formula = "=$H$" & rngDocType.Row
            
            If rngDocType.Value = "" Then
                If tempStr <> wsC_Output.Range("V" & rowOut).Value Then
                    tempStr = ""
                    rngDocType.Value = "SF"
                End If
            End If
            
            rowOut = rowOut + 1
            xRow = xRow + 1
            
            
            While wsC_Output.Range("T" & rowOut).Value = "(blank)"
                rowOut = rowOut + 1
            Wend
            
            If wsC_Output.Range("T" & rowOut).Value = "Grand Total" Then
                rowOut = rowOut + 1
            End If
            
        Wend
        
        'If rngDocType.Value = "" Then
        '    rngDocType.Value = "SB"
        '    tempStr = ""
        'End If
        
        If wsC_Data.Range("K" & xRow - 1).Value <> 20610000 Then
            
            With wsC_Data
                .Range("K" & xRow).Value = 20610000
                
                
                If gotIn = True And gotOut = True Then
                    .Range("L" & xRow).Formula = "=ABS(SUM(L" & rngDocType.Row & ":L" & tempRow & ") - SUM(L" & tempRow + 1 & ":L" & xRow - 1 & "))"
                    .Range("M" & xRow).Formula = "=ABS(SUM(M" & rngDocType.Row & ":M" & tempRow & ") - SUM(M" & tempRow + 1 & ":M" & xRow - 1 & "))"
                    .Range("N" & xRow).Formula = "=ABS(SUM(N" & rngDocType.Row & ":N" & tempRow & ") - SUM(N" & tempRow + 1 & ":N" & xRow - 1 & "))"
                    .Range("O" & xRow).Formula = "=ABS(SUM(O" & rngDocType.Row & ":O" & tempRow & ") - SUM(O" & tempRow + 1 & ":O" & xRow - 1 & "))"
                    
                    .Range("J" & xRow).Formula = "="
                    
                ElseIf gotIn = True Then
                    .Range("L" & xRow).Formula = "=ABS(SUM(L" & rngDocType.Row & ":L" & xRow - 1 & "))"
                    .Range("M" & xRow).Formula = "=ABS(SUM(M" & rngDocType.Row & ":M" & xRow - 1 & "))"
                    .Range("N" & xRow).Formula = "=ABS(SUM(N" & rngDocType.Row & ":N" & xRow - 1 & "))"
                    .Range("O" & xRow).Formula = "=ABS(SUM(O" & rngDocType.Row & ":O" & xRow - 1 & "))"
                    
                Else
                    .Range("L" & xRow).Formula = "=ABS(-SUM(L" & rngDocType.Row & ":L" & xRow - 1 & "))"
                    .Range("M" & xRow).Formula = "=ABS(-SUM(M" & rngDocType.Row & ":M" & xRow - 1 & "))"
                    .Range("N" & xRow).Formula = "=ABS(-SUM(N" & rngDocType.Row & ":N" & xRow - 1 & "))"
                    .Range("O" & xRow).Formula = "=ABS(-SUM(O" & rngDocType.Row & ":O" & xRow - 1 & "))"
                    
                End If
                
                
                If tempVal >= 0 Then
                    .Range("J" & xRow).Value = 40
                Else
                    .Range("J" & xRow).Value = 50
                End If
                
                tempVal = 0
                
                tempRow = 2
                Do While wsProf.Range("A" & tempRow).Value <> ""
                    If wsProf.Range("A" & tempRow).Value = .Range("E" & rngDocType.Row).Value Then
                        .Range("T" & xRow).Value = wsProf.Range("B" & tempRow).Value
                        Exit Do
                    End If
                    tempRow = tempRow + 1
                Loop
                
                
                If rngDocType.Value = "" Then
                    If tempStr <> .Range("T" & xRow).Value Then
                        rngDocType.Value = "SF"
                    Else
                        rngDocType.Value = "SB"
                    End If
                    
                End If
                tempStr = ""
                
                
                .Range("AB" & xRow).Value = 915
                .Range("AD" & xRow).Formula = "=$H$" & rngDocType.Row
                
                .Range("B" & xRow & ":AI" & xRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range("B" & xRow & ":AI" & xRow).Borders(xlEdgeBottom).Weight = xlThin
                
            End With
        End If
        
        xRow = xRow + 1
        myRef = ""
    Wend
    
    
    wsC_Data.Range("L:O").NumberFormat = "#,##0.00"
    
    wsC_Data.Range("N" & xRow + 3).Value = "Input Tax"
    wsC_Data.Range("N" & xRow + 4).Value = "Output Tax"
    
    If Left(wsC_Data.Parent.Name, 3) = "QGC" Or Left(wsC_Data.Parent.Name, 4) = "BGIA" Then
        wsC_Data.Range("O" & xRow + 3).Formula = "=SUMIFS(M4:M" & xRow - 1 & ",J4:J" & xRow - 1 & ",50,K4:K" & xRow - 1 & ",20610101) - SUMIFS(M4:M" & xRow - 1 & ",J4:J" & xRow - 1 & ",40,K4:K" & xRow - 1 & ",20610101)"
        wsC_Data.Range("O" & xRow + 4).Formula = "=SUMIFS(M4:M" & xRow - 1 & ",J4:J" & xRow - 1 & ",40,K4:K" & xRow - 1 & ",20610102) - SUMIFS(M4:M" & xRow - 1 & ",J4:J" & xRow - 1 & ",50,K4:K" & xRow - 1 & ",20610102)"
    
    Else
        wsC_Data.Range("O" & xRow + 3).Formula = "=SUMIFS(O4:O" & xRow - 1 & ",J4:J" & xRow - 1 & ",50,K4:K" & xRow - 1 & ",20610101) - SUMIFS(O4:O" & xRow - 1 & ",J4:J" & xRow - 1 & ",40,K4:K" & xRow - 1 & ",20610101)"
        wsC_Data.Range("O" & xRow + 4).Formula = "=SUMIFS(O4:O" & xRow - 1 & ",J4:J" & xRow - 1 & ",40,K4:K" & xRow - 1 & ",20610102) - SUMIFS(O4:O" & xRow - 1 & ",J4:J" & xRow - 1 & ",50,K4:K" & xRow - 1 & ",20610102)"
    
    End If
    
    
    wbClearing.Save
    
    'do check
    
    wsC_Check.Range("B3").Formula = "=SUMIFS(Data!L:L,Data!J:J,Check!A3)"
    wsC_Check.Range("B4").Formula = "=SUMIFS(Data!L:L,Data!J:J,Check!A4)"
    
    wsC_Check.Range("C3").Formula = "=SUMIFS(Data!M:M,Data!J:J,Check!A3)"
    wsC_Check.Range("C4").Formula = "=SUMIFS(Data!M:M,Data!J:J,Check!A4)"
    
    wsC_Check.Range("D3").Formula = "=SUMIFS(Data!N:N,Data!J:J,Check!A3)"
    wsC_Check.Range("D4").Formula = "=SUMIFS(Data!N:N,Data!J:J,Check!A4)"
    
    wsC_Check.Range("E3").Formula = "=SUMIFS(Data!O:O,Data!J:J,Check!A3)"
    wsC_Check.Range("E4").Formula = "=SUMIFS(Data!O:O,Data!J:J,Check!A4)"
    
    wbClearing.Save
    
    
    'do check2
    
    With wsC_Check2
        .Cells.Delete
        
        .Range("E1").Value = "Doc"
        .Range("F1").Value = "AUD"
        .Range("G1").Value = "GBP"
        .Range("H1").Value = "AUD"
        
        
        If xRow >= 4 Then
            .Range("C2:H" & xRow - 3).Formula = "=Data!J4"
            .Range("E:H").NumberFormat = "#,##0.00"
        End If
        
        
        .Range("K1").Value = 40
        .Range("K2").Value = 50
        .Range("K3").Value = 40
        .Range("K4").Value = 50
        .Range("K5").Value = 40
        .Range("K6").Value = 50
        
        .Range("L1").Value = 20610101
        .Range("L2").Value = 20610101
        .Range("L3").Value = 20610102
        .Range("L4").Value = 20610102
        .Range("L5").Value = 20610000
        .Range("L6").Value = 20610000
        .Range("L7").Value = 20610101
        .Range("L8").Value = 20610102
        .Range("L9").Value = 20610000
        
        .Range("M1:P6").Formula = "=SUMIFS(E:E,$C:$C,$K1,$D:$D,$L1)*IF($K1 = 50,-1,1)"
        .Range("M7:P7").Formula = "=M1 + M2"
        .Range("M8:P8").Formula = "=M3 + M4"
        .Range("M9:P9").Formula = "=M5 + M6"
        .Range("M10:P10").Formula = "=SUM(M7:M9)"
        
        .Range("M10:P10").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("M10:P10").Borders(xlEdgeTop).Weight = xlThin
        .Range("M10:P10").Borders(xlEdgeTop).LineStyle = xlDouble
        .Range("M10:P10").Borders(xlEdgeTop).Weight = xlThick
        
        .Range("M1:P10").NumberFormat = "#,##0.00_);(#,##0.00)"
        
        .Range("K1:P2").Interior.Color = RGB(296, 215, 155)
        .Range("K5:P6").Interior.Color = RGB(252, 213, 180)
        
        '.Range("K5:P5").Interior.Color = RGB(252, 213, 180)
        '.Range("K6:P6").Interior.Color = RGB(296, 215, 155)
        
    End With
    
    wbClearing.Save
    
End Sub























