VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf1 
   Caption         =   "Australia SAPL & SEHAL Automation"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14730
   OleObjectBlob   =   "NAS_UF1_ETL.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "uf1"
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
'instructions to run
'**************************************************************************************************************
'download 11 Spool reports (6 from SAPL, 5 from SEHAL) from SAP

Const runTest As Boolean = False
Const testFilePath As String = "C:\Users\Nicholas.N.Tay\Desktop\Project\Aus SS\Testing\To test\"



'**************************************************************************************************************
'Declaration of module variables
'**************************************************************************************************************

Private StartTime       As Double, SecondsElapsed As Double, MinutesElapsed As String
Private wsSettings      As Worksheet, wsRef As Worksheet, wsSample As Worksheet, wsExclude As Worksheet
Private inputError      As String
Private Path_Main       As String, Path_SAPL_M As String, Path_SEHAL_M As String
Private xCount          As Long, mySeq As Long

Private finalMessage    As String
Private myPath(11)      As String, myFileName(10) As String, myType(10) As String
Private spoolNo(10)     As String, BacNo(10) As String

Private wbMain          As Workbook
Private wsData          As Worksheet, wsSupporting As Worksheet, wsException As Worksheet, wsExchangeRate As Worksheet, wsGSTExclude As Worksheet
Private myTable

Private wbSA1 As Workbook, wbSA2 As Workbook, wbSA3 As Workbook, wbSA4 As Workbook, wbSA5 As Workbook, wbSA6 As Workbook
Private wbSE1 As Workbook, wbSE2 As Workbook, wbSE3 As Workbook, wbSE4 As Workbook, wbSE5 As Workbook

Private wsSA1a As Worksheet, wsSA2a As Worksheet, wsSA3a As Worksheet, wsSA4a As Worksheet, wsSA5a As Worksheet, wsSA6a As Worksheet
Private wsSE1a As Worksheet, wsSE2a As Worksheet, wsSE3a As Worksheet, wsSE4a As Worksheet, wsSE5a As Worksheet

Private wsSA1b As Worksheet, wsSA2b As Worksheet, wsSA3b As Worksheet, wsSA4b As Worksheet, wsSA5b As Worksheet, wsSA6b As Worksheet
Private wsSE1b As Worksheet, wsSE2b As Worksheet, wsSE3b As Worksheet, wsSE4b As Worksheet, wsSE5b As Worksheet

Private wsSA1c As Worksheet, wsSA2c As Worksheet, wsSA3c As Worksheet, wsSA4c As Worksheet, wsSA5c As Worksheet, wsSA6c As Worksheet
Private wsSE1c As Worksheet, wsSE2c As Worksheet, wsSE3c As Worksheet, wsSE4c As Worksheet, wsSE5c As Worksheet

Private wsSA1d As Worksheet, wsSA2d As Worksheet, wsSA3d As Worksheet, wsSA4d As Worksheet, wsSA5d As Worksheet, wsSA6d As Worksheet
Private wsSE1d As Worksheet, wsSE2d As Worksheet, wsSE3d As Worksheet, wsSE4d As Worksheet, wsSE5d As Worksheet

Private wsSA1e As Worksheet, wsSA2e As Worksheet, wsSA3e As Worksheet, wsSA4e As Worksheet, wsSA5e As Worksheet, wsSA6e As Worksheet
Private wsSE1e As Worksheet, wsSE2e As Worksheet, wsSE3e As Worksheet, wsSE4e As Worksheet, wsSE5e As Worksheet


Private wbOth1 As Workbook, wbOth2 As Workbook

Private wsOth1a As Worksheet, wsOth2a As Worksheet, wsOth1b As Worksheet, wsOth2b As Worksheet, wsOth1c As Worksheet, wsOth2c As Worksheet, wsOth1d As Worksheet, wsOth2d As Worksheet, wsOth1e As Worksheet, wsOth2e As Worksheet

Private wbFBL3N As Workbook, wsFBL3N As Worksheet
Private wsExch As Worksheet

Private myDocNo As String

Private wsXX As Worksheet

Private wbCrosscheck As Workbook, wsCrosscheck As Worksheet

Private mySPSite As String



Private doPDF As Boolean


'**************************************************************************************************************
'Main Routine
'**************************************************************************************************************

Private Sub BtnRun_Click()
    
    
    Me.CheckBox1.Value = doPDF
    
    StartTime = Timer
    inputError = ""
    
    'check if required fields are filled
    checkFields
    
    If inputError <> "" Then
        MsgBox ("Please ensure the following issues are corrected before proceeding:" & inputError)
        Exit Sub
    End If
    
    If runTest = False Then
        'check if SAP BP is open
        checkOpenSAP
        If openP16 = False Then
            MsgBox ("Please ensure you have SAP Blueprint open.")
            Exit Sub
        End If
        
    End If
    RunPauseAll
    
    
    
    Me.Frame1.Visible = True
    Me.btnClose.Visible = False
    Me.Height = 196
    Me.Width = 239.5
    
    
    
    
    '*************not done
    'open file and check tabs
    openFile
    
    If runTest = False Then
        'create save path folder
        createAllFolders
        
        'initiate downloading in queue
        StartDownloadQueue
    Else
        readFolders
    End If
    
    'initiate population of input form
    setAllFiles
    
    'initiate population of input form
    setInputForm
    
    
    'doSP
    doCrossCheck
    
    
    RunActivateAll
    Call closeFile
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MinutesElapsed = VBA.Format((Timer - StartTime) / 86400, "hh:mm:ss")
    
    
    If finalMessage = "" Then
        Call ChangeProgress("Automation process completed for Australia SAPL & SEHAL Preparation" & vbLf & "Time taken: " & MinutesElapsed, 1)
        
        'MsgBox ("Automation process completed for Australia SAPL & SEHAL Preparation" & vbLf & "Time taken: " & MinutesElapsed)
    Else
        Call ChangeProgress("Automation process completed for Australia SAPL & SEHAL Preparation" & vbLf & "Time taken: " & MinutesElapsed & vbLf & vbLf & "Errors: " & finalMessage, 1)
        'MsgBox ("Automation process completed for Australia SAPL & SEHAL Preparation" & vbLf & "Time taken: " & MinutesElapsed & vbLf & vbLf & "Errors: " & finalMessage)
    End If
    
    Me.btnClose.Visible = True
    
    'Unload Me
    
    Exit Sub
    
    
noERP:
    MsgBox ("Please ensure you have SAP Blueprint open.")
    Exit Sub
    
End Sub




'**************************************************************************************************************
'Requirement Runs
'**************************************************************************************************************


Private Sub UserForm_Initialize()
    'ShowTitleBar Me
    Me.Frame1.Visible = False
    
    
    Set wsSettings = ThisWorkbook.Worksheets("Settings")
    Set wsRef = ThisWorkbook.Worksheets("Data Entry")
    Set wsSample = ThisWorkbook.Worksheets("Sample")
    Set wsExclude = ThisWorkbook.Worksheets("To Exclude")
    Set wsExch = ThisWorkbook.Worksheets("Exchange Rate")
    Set wsXX = ThisWorkbook.Worksheets("Main")
    
    If Not VBA.Environ("username") = wsSettings.Range("B1").Value Then
        wsSettings.Range("B:B").EntireColumn.Clear
        wsSettings.Range("B1").Value = VBA.Environ("username")
    Else
        Me.tbSaveLocation.Value = wsSettings.Range("B2").Value
        Me.tbCrosscheck.Value = wsSettings.Range("B3").Value
    End If
    
    If IsNumeric(wsXX.Range("P15").Value) Then
        Me.tb_TotalGST.Value = wsXX.Range("P15").Value
    End If
    
    If IsNumeric(wsXX.Range("P17").Value) Then
        Me.tb_EmpContribution.Value = wsXX.Range("P17").Value
    End If
    
    If IsNumeric(wsXX.Range("P19").Value) Then
        Me.tb_GSTPayable.Value = wsXX.Range("P19").Value
    End If
    
    
End Sub

Private Sub btnSaveLocation_Click()
    Me.tbSaveLocation.Value = SearchFolderLocation
    wsSettings.Range("B2").Value = Me.tbSaveLocation.Value
End Sub

Private Sub btnCrossCheck_Click()
    Me.tbCrosscheck.Value = SearchFileLocation
    wsSettings.Range("B3").Value = Me.tbCrosscheck.Value
End Sub



'Private Sub btnReport_Click()
'    Me.tbReport.Value = SearchFileLocation
'    wsSettings.Range("B3").Value = Me.tbReport.Value
'End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub


'**************************************************************************************************************
'Processing Runs
'**************************************************************************************************************

Sub ChangeProgress(VarTitle As String, PercValue As Double)
    
    Me.LabelCaption.Caption = VarTitle
    
    If PercValue > 1 Then
        LabelProgress.Width = PercValue / 100 * FrameProgress.Width
    Else
        LabelProgress.Width = PercValue * FrameProgress.Width
    End If
    
    Me.Repaint
End Sub







Sub checkFields()
    'If Me.tbReport.Value = "" Then
    '    inputError = inputError & vbLf & " - SAPL SEHAL Form (1)"
    'ElseIf Dir(Me.tbReport.Value, vbNormal) = "" Then
    '    inputError = inputError & vbLf & " - SAPL SEHAL Form (2)"
    'End If
    
    If Me.tbSaveLocation.Value = "" Then
        inputError = inputError & vbLf & " - Save Location (1)"
    ElseIf Dir(Me.tbSaveLocation.Value, vbDirectory) = "" Then
        inputError = inputError & vbLf & " - Save Location (2)"
    End If
    
    
    If Me.tbCrosscheck.Value = "" Then
        inputError = inputError & vbLf & " - CrossCheck File (1)"
    ElseIf Dir(Me.tbCrosscheck.Value) = "" Then
        inputError = inputError & vbLf & " - CrossCheck File (2)"
    End If
    
    
    If Me.tb_TotalGST.Value = "" Then
    ElseIf Not (IsNumeric(Me.tb_TotalGST.Value)) Then
        inputError = inputError & vbLf & " - Total GST not numeric"
    End If
    
    If Me.tb_EmpContribution.Value = "" Then
    ElseIf Not (IsNumeric(Me.tb_EmpContribution.Value)) Then
        inputError = inputError & vbLf & " - Employee Contribution not numeric"
    End If
    
    If Me.tb_GSTPayable.Value = "" Then
    ElseIf Not (IsNumeric(Me.tb_GSTPayable.Value)) Then
        inputError = inputError & vbLf & " - Total GST Payable not numeric"
    End If
    
    
End Sub


Sub openFile()
    
    Call ChangeProgress("Setting Crosscheck File", 0.01)
    Dim ws As Worksheet
    
    Set wbCrosscheck = Workbooks.Open(Me.tbCrosscheck.Value, ReadOnly:=True)
    
    For Each ws In wbCrosscheck.Worksheets
        If ws.Name = "GST Cross Check" Then
            Set wsCrosscheck = ws
            Exit For
        End If
    Next
    
    If wsCrosscheck Is Nothing Then
        wbCrosscheck.Close False
        Call RunActivateAll
        MsgBox ("Tab missing in CrossCheck file. Please ensure the right file is selected.")
        End
    End If
    
    
    
    Call ChangeProgress("Creating Input Forms", 0.012)
    
    Set wbMain = Workbooks.Add
    
    Set wsSupporting = wbMain.Worksheets(1)
    wsSupporting.Name = "Supporting Data"
    
    'wsRef.Visible = xlSheetVisible
    wsRef.Copy before:=wbMain.Worksheets(1)
    'wsRef.Visible = xlSheetVeryHidden
    Set wsData = wbMain.Worksheets(1)
    
    Set wsException = wbMain.Worksheets.Add(after:=wbMain.Sheets(wbMain.Sheets.Count))
    wsException.Name = "Exception Rule"
    
    Set wsExchangeRate = wbMain.Worksheets.Add(after:=wbMain.Sheets(wbMain.Sheets.Count))
    wsExchangeRate.Name = "Exchange Rate"
    
    Set wsGSTExclude = wbMain.Worksheets.Add(after:=wbMain.Sheets(wbMain.Sheets.Count))
    wsGSTExclude.Name = "SEHAL GST GROUP"
    
    
    Call ChangeProgress("Clearing and setting input form fields ", 0.02)
    
    With wsData
        .Range("G1").Value = Month(DateAdd("m", -1, VBA.Date))
        .Range("H1").Value = Year(DateAdd("m", -1, VBA.Date))
        .Range("H3").Value = VBA.Date
    
        .Range("I94").Value = "Fleetplus " & VBA.Format(DateAdd("m", -1, VBA.Date), "MMMM YYYY")
    
        .Range("G56").Value = 0
        .Range("G57").Value = 0
        .Range("G58").Value = 0
        .Range("G59").Value = 0
        .Range("G60").Value = 0
        .Range("G61").Value = 0

        .Range("G86").Value = 0

        .Range("G90").Value = 0
        .Range("G91").Value = 0
        
    End With
    
    wsSupporting.Activate
    wsSupporting.Select
End Sub


Sub readFolders()
    Dim x As Long
    
    Path_Main = testFilePath & "Part 1\"
    
    Path_SAPL_M = Path_Main & "SAPL\"
    Path_SEHAL_M = Path_Main & "SEHAL\"
        
    myFileName(0) = "ACP41"
    myFileName(1) = "ACP52"
    myFileName(2) = "NTP48"
    myFileName(3) = "PALTA"
    myFileName(4) = "SEDNA (CRUX)"
    myFileName(5) = "PRELUDE"
    
    myFileName(6) = "AU01"
    myFileName(7) = "AU02"
    myFileName(8) = "AU10"
    myFileName(9) = "AU11"
    myFileName(10) = "AU03-AU09"
    
    For x = 0 To 10
        If x < 6 Then
            
            myPath(x) = Path_SAPL_M & myFileName(x) & "\"
        Else
            myPath(x) = Path_SEHAL_M & myFileName(x) & "\"
        End If
    Next
        
    myPath(11) = Path_SEHAL_M
        
        
        
End Sub


Sub createAllFolders()
    Dim x As Long, xx As Long
    Dim pathTemp As String
    
    
    pathTemp = createFolder(Me.tbSaveLocation.Value, "AUS SAPL SEHAL")
    MkDir pathTemp & "Part 1\"
    MkDir pathTemp & "Part 2\"
    
    wsSettings.Range("B5").Value = pathTemp
    
    Me.LabelCaption.Caption = "Creating input files"
    
    
    For xx = 1 To 2
        If xx = 1 Then
            Call ChangeProgress("Creating folders Part 2", 0.03)
            
            Path_Main = pathTemp & "Part 2\"
        Else
            Call ChangeProgress("Creating folders Part 1", 0.05)
            Path_Main = pathTemp & "Part 1\"
        End If
    
        Path_SAPL_M = Path_Main & "SAPL\"
        Path_SEHAL_M = Path_Main & "SEHAL\"
        
        MkDir Path_SAPL_M
        MkDir Path_SEHAL_M
        
        myFileName(0) = "ACP41"
        myFileName(1) = "ACP52"
        myFileName(2) = "NTP48"
        myFileName(3) = "PALTA"
        myFileName(4) = "SEDNA (CRUX)"
        myFileName(5) = "PRELUDE"
        
        myFileName(6) = "AU01"
        myFileName(7) = "AU02"
        myFileName(8) = "AU10"
        myFileName(9) = "AU11"
        myFileName(10) = "AU03-AU09"
        
        For x = 0 To 10
            If x < 6 Then
                myPath(x) = Path_SAPL_M & myFileName(x) & "\"
            Else
                myPath(x) = Path_SEHAL_M & myFileName(x) & "\"
            End If
            
            MkDir myPath(x)
        Next
        
        myPath(11) = Path_SEHAL_M
        'MkDir myPath(11)
    Next
    
End Sub


Sub StartDownloadQueue()
    SAP_BP_Clear
    
    SAP_BP_1
    SAP_BP_2
    SAP_BP_3
    SAP_BP_4 'download pdf
    
    SAP_BP_7
    
    SAP_BP_6
    
    
    Call ChangeProgress("Opening FBL3N", 0.68)
    
    Set wbFBL3N = Workbooks.Open(Path_Main & "FBL3N.xlsx", ReadOnly:=True)
    Set wsFBL3N = wbFBL3N.Worksheets(1)
    
    
    
End Sub


Sub setAllFiles()
    Dim wb As Workbook
    
    Me.LabelCaption.Caption = "Creating input files"
    
    
    Call ChangeProgress("Preparing Input Form Files", 0.7)
    
    wbMain.SaveAs myPath(0) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(0) & " Input Form.xlsx"
    wbMain.SaveAs myPath(1) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(1) & " Input Form.xlsx"
    wbMain.SaveAs myPath(2) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(2) & " Input Form.xlsx"
    wbMain.SaveAs myPath(3) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(3) & " Input Form.xlsx"
    wbMain.SaveAs myPath(4) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(4) & " Input Form.xlsx"
    wbMain.SaveAs myPath(5) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(5) & " Input Form.xlsx"
    wbMain.SaveAs myPath(6) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(6) & " Input Form.xlsx"
    wbMain.SaveAs myPath(7) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(7) & " Input Form.xlsx"
    wbMain.SaveAs myPath(8) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(8) & " Input Form.xlsx"
    wbMain.SaveAs myPath(9) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(9) & " Input Form.xlsx"
    
    wbMain.SaveAs myPath(11) & Format(DateAdd("m", -1, Date), "YYYYMM") & " Shell Australia Services Company Pty Ltd Input Form.xlsx"
    wbMain.SaveAs myPath(11) & Format(DateAdd("m", -1, Date), "YYYYMM") & " Shell Australia Lubricants Production Pty Ltd Input Form.xlsx"
    
    
    Set wbSA1 = Workbooks.Open(myPath(0) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(0) & " Input Form.xlsx")
    Set wsSA1a = wbSA1.Worksheets("Data Entry")
    Set wsSA1b = wbSA1.Worksheets("Supporting Data")
    Set wsSA1c = wbSA1.Worksheets("Exception Rule")
    Set wsSA1d = wbSA1.Worksheets("Exchange Rate")
    Set wsSA1e = wbSA1.Worksheets("SEHAL GST GROUP")
    
    Set wbSA2 = Workbooks.Open(myPath(1) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(1) & " Input Form.xlsx")
    Set wsSA2a = wbSA2.Worksheets("Data Entry")
    Set wsSA2b = wbSA2.Worksheets("Supporting Data")
    Set wsSA2c = wbSA2.Worksheets("Exception Rule")
    Set wsSA2d = wbSA2.Worksheets("Exchange Rate")
    Set wsSA2e = wbSA2.Worksheets("SEHAL GST GROUP")
    
    Set wbSA3 = Workbooks.Open(myPath(2) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(2) & " Input Form.xlsx")
    Set wsSA3a = wbSA3.Worksheets("Data Entry")
    Set wsSA3b = wbSA3.Worksheets("Supporting Data")
    Set wsSA3c = wbSA3.Worksheets("Exception Rule")
    Set wsSA3d = wbSA3.Worksheets("Exchange Rate")
    Set wsSA3e = wbSA3.Worksheets("SEHAL GST GROUP")
    
    Set wbSA4 = Workbooks.Open(myPath(3) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(3) & " Input Form.xlsx")
    Set wsSA4a = wbSA4.Worksheets("Data Entry")
    Set wsSA4b = wbSA4.Worksheets("Supporting Data")
    Set wsSA4c = wbSA4.Worksheets("Exception Rule")
    Set wsSA4d = wbSA4.Worksheets("Exchange Rate")
    Set wsSA4e = wbSA4.Worksheets("SEHAL GST GROUP")
    
    Set wbSA5 = Workbooks.Open(myPath(4) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(4) & " Input Form.xlsx")
    Set wsSA5a = wbSA5.Worksheets("Data Entry")
    Set wsSA5b = wbSA5.Worksheets("Supporting Data")
    Set wsSA5c = wbSA5.Worksheets("Exception Rule")
    Set wsSA5d = wbSA5.Worksheets("Exchange Rate")
    Set wsSA5e = wbSA5.Worksheets("SEHAL GST GROUP")
    
    Set wbSA6 = Workbooks.Open(myPath(5) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(5) & " Input Form.xlsx")
    Set wsSA6a = wbSA6.Worksheets("Data Entry")
    Set wsSA6b = wbSA6.Worksheets("Supporting Data")
    Set wsSA6c = wbSA6.Worksheets("Exception Rule")
    Set wsSA6d = wbSA6.Worksheets("Exchange Rate")
    Set wsSA6e = wbSA6.Worksheets("SEHAL GST GROUP")
    
    Set wbSE1 = Workbooks.Open(myPath(6) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(6) & " Input Form.xlsx")
    Set wsSE1a = wbSE1.Worksheets("Data Entry")
    Set wsSE1b = wbSE1.Worksheets("Supporting Data")
    Set wsSE1c = wbSE1.Worksheets("Exception Rule")
    Set wsSE1d = wbSE1.Worksheets("Exchange Rate")
    Set wsSE1e = wbSE1.Worksheets("SEHAL GST GROUP")
    
    Set wbSE2 = Workbooks.Open(myPath(7) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(7) & " Input Form.xlsx")
    Set wsSE2a = wbSE2.Worksheets("Data Entry")
    Set wsSE2b = wbSE2.Worksheets("Supporting Data")
    Set wsSE2c = wbSE2.Worksheets("Exception Rule")
    Set wsSE2d = wbSE2.Worksheets("Exchange Rate")
    Set wsSE2e = wbSE2.Worksheets("SEHAL GST GROUP")
    
    Set wbSE3 = Workbooks.Open(myPath(8) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(8) & " Input Form.xlsx")
    Set wsSE3a = wbSE3.Worksheets("Data Entry")
    Set wsSE3b = wbSE3.Worksheets("Supporting Data")
    Set wsSE3c = wbSE3.Worksheets("Exception Rule")
    Set wsSE3d = wbSE3.Worksheets("Exchange Rate")
    Set wsSE3e = wbSE3.Worksheets("SEHAL GST GROUP")
    
    Set wbSE4 = Workbooks.Open(myPath(9) & Format(DateAdd("m", -1, Date), "YYYYMM") & " " & myFileName(9) & " Input Form.xlsx")
    Set wsSE4a = wbSE4.Worksheets("Data Entry")
    Set wsSE4b = wbSE4.Worksheets("Supporting Data")
    Set wsSE4c = wbSE4.Worksheets("Exception Rule")
    Set wsSE4d = wbSE4.Worksheets("Exchange Rate")
    Set wsSE4e = wbSE4.Worksheets("SEHAL GST GROUP")
    
    Set wbOth1 = Workbooks.Open(myPath(11) & Format(DateAdd("m", -1, Date), "YYYYMM") & " Shell Australia Services Company Pty Ltd Input Form.xlsx")
    Set wsOth1a = wbOth1.Worksheets("Data Entry")
    Set wsOth1b = wbOth1.Worksheets("Supporting Data")
    Set wsOth1c = wbOth1.Worksheets("Exception Rule")
    Set wsOth1d = wbOth1.Worksheets("Exchange Rate")
    Set wsOth1e = wbOth1.Worksheets("SEHAL GST GROUP")
    
    Set wbOth2 = Workbooks.Open(myPath(11) & Format(DateAdd("m", -1, Date), "YYYYMM") & " Shell Australia Lubricants Production Pty Ltd Input Form.xlsx")
    Set wsOth2a = wbOth2.Worksheets("Data Entry")
    Set wsOth2b = wbOth2.Worksheets("Supporting Data")
    Set wsOth2c = wbOth2.Worksheets("Exception Rule")
    Set wsOth2d = wbOth2.Worksheets("Exchange Rate")
    Set wsOth2e = wbOth2.Worksheets("SEHAL GST GROUP")
    
    wsSA1a.Range("B3").Value = "ACP/41 JOINT VENTURE"
    wsSA1a.Range("B4").Value = "14 009 663 576"
    wsSA1a.Range("B5").Value = "ACP41"
    
    wsSA2a.Range("B3").Value = "ACP52 JOINT VENTURE"
    wsSA2a.Range("B4").Value = "14 009 663 576"
    wsSA2a.Range("B5").Value = "ACP52"

    wsSA3a.Range("B3").Value = "NTP48 JOINT VENTURE"
    wsSA3a.Range("B4").Value = "14 009 663 576"
    wsSA3a.Range("B5").Value = "NTP48"

    wsSA4a.Range("B3").Value = "PALTA JOINT VENTURE"
    wsSA4a.Range("B4").Value = "14 009 663 576"
    wsSA4a.Range("B5").Value = "PALTA JOINT VENTURE"

    wsSA5a.Range("B3").Value = "AC/L9 CRUX JV"
    wsSA5a.Range("B4").Value = "14 009 663 576"
    wsSA5a.Range("B5").Value = "SEDNA (CRUX)"

    wsSA6a.Range("B3").Value = "PRELUDE JOINT VENTURE"
    wsSA6a.Range("B4").Value = "14 009 663 576"
    wsSA6a.Range("B5").Value = "PRELUDE"

    wsSE1a.Range("B3").Value = "Shell Australia Pty Ltd"
    wsSE1a.Range("B4").Value = "14 009 663 576"
    wsSE1a.Range("B5").Value = "AU01"

    wsSE2a.Range("B3").Value = "Shell Energy Holdings Australia Ltd, ABN"
    wsSE2a.Range("B4").Value = "14 009 663 576"
    wsSE2a.Range("B5").Value = "AU02"

    wsSE3a.Range("B3").Value = "Shell FLNG"
    wsSE3a.Range("B4").Value = "32 008 551 068"
    wsSE3a.Range("B5").Value = "AU10"

    wsSE4a.Range("B3").Value = "Shell Global Solution Australia Pty Ltd"
    wsSE4a.Range("B4").Value = "34 091 702 448"
    wsSE4a.Range("B5").Value = "AU11"
    
    wsOth1a.Range("B3").Value = "Shell Australia Services Company Pty Ltd"
    wsOth1a.Range("B4").Value = ""
    wsOth1a.Range("B5").Value = "2002"

    wsOth2a.Range("B3").Value = "Shell Australia Lubricants Production Pty Ltd"
    wsOth2a.Range("B4").Value = ""
    wsOth2a.Range("B5").Value = "1050"
    
    
    wsData.Range("B4").Value = "14 009 663 576"
    
    wsData.Range("B3").Value = "North West Shelf LNG Pty Ltd"
    wsData.Range("B5").Value = "AU03"
    wbMain.SaveAs myPath(10) & Format(DateAdd("m", -1, Date), "YYYYMM") & " AU03 Input Form.xlsx"
    
    wsData.Range("B3").Value = "SAPL (PSC19)"
    wsData.Range("B5").Value = "AU06"
    wbMain.SaveAs myPath(10) & Format(DateAdd("m", -1, Date), "YYYYMM") & " AU06 Input Form.xlsx"
    
    wsData.Range("B3").Value = "SAPL (PSC20)"
    wsData.Range("B5").Value = "AU07"
    wbMain.SaveAs myPath(10) & Format(DateAdd("m", -1, Date), "YYYYMM") & " AU07 Input Form.xlsx"
    
    wsData.Range("B3").Value = "Shell Aus Nat Gas Shipping"
    wsData.Range("B5").Value = "AU08"
    wbMain.SaveAs myPath(10) & Format(DateAdd("m", -1, Date), "YYYYMM") & " AU08 Input Form.xlsx"
    
    wsData.Range("B3").Value = "Shell Energy Investments"
    wsData.Range("B5").Value = "AU09"
    wbMain.SaveAs myPath(10) & Format(DateAdd("m", -1, Date), "YYYYMM") & " AU09 Input Form.xlsx"
    
    
    wbMain.Close True
    wbOth1.Close True
    'wbOth2.Close True
    
    Set wbMain = Nothing
    Set wsData = Nothing
    Set wsSupporting = Nothing
    Set wsException = Nothing
    
    Set wbOth1 = Nothing
    Set wsOth1a = Nothing
    Set wsOth1b = Nothing
    Set wsOth1c = Nothing
    
    Set wbOth2 = Nothing
    Set wsOth2a = Nothing
    Set wsOth2b = Nothing
    Set wsOth2c = Nothing
    
End Sub



Sub setInputForm()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim myCount As Long
    
    Call ChangeProgress("Preparing Input Form Formatting", 0.72)
    
    For myCount = 0 To 9
        Select Case myCount
        Case 0
            Set ws1 = wsSA1a
            Set ws2 = wsSA1b
            Set ws3 = wsSA1c
        Case 1
            Set ws1 = wsSA2a
            Set ws2 = wsSA2b
            Set ws3 = wsSA2c
        Case 2
            Set ws1 = wsSA3a
            Set ws2 = wsSA3b
            Set ws3 = wsSA3c
        Case 3
            Set ws1 = wsSA4a
            Set ws2 = wsSA4b
            Set ws3 = wsSA4c
        Case 4
            Set ws1 = wsSA5a
            Set ws2 = wsSA5b
            Set ws3 = wsSA5c
        Case 5
            Set ws1 = wsSA6a
            Set ws2 = wsSA6b
            Set ws3 = wsSA6c
        Case 6
            Set ws1 = wsSE1a
            Set ws2 = wsSE1b
            Set ws3 = wsSE1c
        Case 7
            Set ws1 = wsSE2a
            Set ws2 = wsSE2b
            Set ws3 = wsSE2c
        Case 8
            Set ws1 = wsSE3a
            Set ws2 = wsSE3b
            Set ws3 = wsSE3c
        Case 9
            Set ws1 = wsSE4a
            Set ws2 = wsSE4b
            Set ws3 = wsSE4c
        End Select
        
        
        Call ChangeProgress("Getting Supporting Values - " & myFileName(myCount), 0.72 + myCount * 0.02)
        
        Call moveSupporting(ws2, myCount)
        
        If myCount = 6 Then
            Call specialAU01
        End If
        
        ws2.Cells.WrapText = False
        ws2.Cells.EntireColumn.AutoFit
        
        
        Call ChangeProgress("Update Data Entry values - " & myFileName(myCount), 0.72 + myCount * 0.029)
        
        Call doUpdateData(ws1, ws2, ws3, myCount)
        
    Next
    
    
End Sub


Sub moveSupporting(myWS As Worksheet, myCount) 'myID As Long)
    myWS.Cells.Delete
    'Call setBAS_Alt(myWs, myPath(myCount) & "BAS " & myFileName(myCount) & ".pdf")
    
    doPDF = True
    If doPDF = True Then
    'If Me.CheckBox1.Value = True Then
        If Not Len(Dir(myPath(myCount) & "BAS " & myFileName(myCount) & ".pdf")) = 0 Then
            
            Me.LabelCaption.Caption = "Reading from PDF - " & myFileName(myCount)
            
            Call collectBAS(myWS, myPath(myCount) & "BAS " & myFileName(myCount) & ".pdf")
        End If
    End If
    
    Me.LabelCaption.Caption = "Filtering Exception Rules - " & myFileName(myCount)
    
    copySpool (myCount)
    
    Me.LabelCaption.Caption = "Copy Exchange Rates - " & myFileName(myCount)
    
    copyFBL3N (myCount)
    
    If myCount >= 6 And myCount <= 9 Then
        If runTest = False Then
            Me.LabelCaption.Caption = "Calling FS10N value - " & myFileName(myCount)
            Call SAP_BP_5(myFileName(myCount), myWS)
        End If
        
        With myWS.Range("G3:K21")
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
        
        With myWS.Range("M3:Q21")
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
        
    End If
    
    
End Sub


Public Function ReadAcrobatDocument(strFileName As String) As String
    Dim AcroApp As CAcroApp, AcroAVDoc As CAcroAVDoc, AcroPDDoc As CAcroPDDoc
    Dim AcroHiliteList As CAcroHiliteList, AcroTextSelect As CAcroPDTextSelect
    Dim PageNumber, PageContent, Content, i, j
    
    Dim repCount As Long
    repCount = 0
    
    
startAgain:
    Err.Clear
    On Error GoTo startAgain2
    
    
    repCount = repCount + 1
    
    If repCount >= 30 Then
        If MsgBox("Multiple attempts has been done to extract PDF. Do you want to continue attempt to read PDF?", vbYesNo) = vbYes Then
            repCount = 1
        Else
            doPDF = False
            ReadAcrobatDocument = ""
            Exit Function
        End If
    End If
    
    Set AcroApp = CreateObject("AcroExch.App")
    Set AcroAVDoc = CreateObject("AcroExch.AVDoc")
    If AcroAVDoc.Open(strFileName, vbNull) <> True Then Exit Function
    ' The following While-Wend loop shouldn't be necessary but timing issues may occur.
    While AcroAVDoc Is Nothing
      Set AcroAVDoc = AcroApp.GetActiveDoc
    Wend
    Set AcroPDDoc = AcroAVDoc.GetPDDoc
    For i = 0 To AcroPDDoc.GetNumPages - 1
      Set PageNumber = AcroPDDoc.AcquirePage(i)
      Set PageContent = CreateObject("AcroExch.HiliteList")
      If PageContent.Add(0, 9000) <> True Then Exit Function
      Set AcroTextSelect = PageNumber.CreatePageHilite(PageContent)
      ' The next line is needed to avoid errors with protected PDFs that can't be read
      On Error Resume Next
      For j = 0 To AcroTextSelect.GetNumText - 1
        Content = Content & AcroTextSelect.GetText(j)
      Next j
    Next i
    
    
    AcroAVDoc.Close True
    AcroApp.Exit
    Set AcroAVDoc = Nothing: Set AcroApp = Nothing
    
startAgain2:
    
    ReadAcrobatDocument = Content
    
End Function


Sub collectBAS(myWS As Worksheet, myName As String)
    Dim myStrSplit() As String
    Dim nextSplit() As String
    Dim xRow As Long
    Dim strCount As Long
    Dim tempStr As String
    Dim repLoop As Long
    
    repLoop = 0
    
    myWS.Range("A1").Value = "BAS"
    
    With myWS.Range("A9:C28")
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    
    
    
    
restartAgain:
    Err.Clear
    repLoop = repLoop + 1
    
    If repLoop >= 20 Then
        If MsgBox("Multiple attempts has been done to extract PDF. You can try ending task of AcroTray from Task Manager or to skip reading from PDF. Do you want to continue attempt to read PDF?", vbYesNo) = vbYes Then
            repLoop = 1
        Else
            doPDF = False
        End If
    End If
    
    
    If doPDF = False Then Exit Sub
    
    tempStr = ReadAcrobatDocument(myName)
    
    If tempStr = "" Then GoTo restartAgain
    
    myStrSplit = Split(tempStr, Chr(10))
    On Error GoTo restartAgain
    'Debug.Print UBound(myStrSplit)
    If UBound(myStrSplit) < 23 Then GoTo restartAgain
    'If IsError(myStrSplit(23)) Then GoTo restartAgain
    
    myWS.Range("A3").Value = myStrSplit(21)
    myWS.Range("A4").Value = myStrSplit(22)
    myWS.Range("A5").Value = myStrSplit(23)
    myWS.Range("A6").Value = myStrSplit(20)
    
    On Error GoTo 0
    
    xRow = 9
    For strCount = 0 To 19
        nextSplit = Split(Left(myStrSplit(strCount), Len(myStrSplit(strCount)) - 4), ":")
        myWS.Range("A" & xRow).Value = Trim(nextSplit(0))
        myWS.Range("B" & xRow).Value = Trim(nextSplit(1))
        myWS.Range("C" & xRow).Value = "AUD"
        
        xRow = xRow + 1
    Next
    
    
    
    
End Sub


Sub specialAU01()
    
    Me.LabelCaption.Caption = "Recording Fleetplus values for AU01"
    
    
    With wsSE1b
        .Range("T1").Value = "Fleetplus Adjustment"
        .Range("T1").Font.Bold = True
        
        .Range("U3").Value = "Input"
        .Range("V3").Value = "Output"
        .Range("W3").Value = "Total Claimable"
        
        .Range("T4").Value = DateSerial(Year(Date), Month(Date) - 1, 1)
        .Range("T4").NumberFormat = "MMM-YY"
        .Range("T5").Value = "Adjustment BAS"
        
        .Range("U3:V3").Font.Bold = True
        .Range("U3:V3").Interior.Color = RGB(255, 255, 0)
        .Range("U5:V5").Font.Bold = True
        .Range("U5:V5").Font.Underline = xlUnderlineStyleSingle
        .Range("U5:V5").Interior.Color = RGB(146, 208, 80)
        
        .Range("U4").Value = Me.tb_TotalGST.Value
        .Range("V4").Value = Me.tb_EmpContribution.Value
        .Range("W4").Value = Me.tb_GSTPayable.Value
        
        .Range("U5").Formula = "=U4*11"
        .Range("V5").Formula = "=V4*11"
    End With
    
    With wsSE1b.Range("T3:W4")
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    
    
End Sub



Sub copySpool(myCount As Long)
    Dim tempWB As Workbook
    Dim tempWS As Worksheet
    Dim ws3 As Worksheet, ws4 As Worksheet
    Dim ws As Worksheet, ws2 As Worksheet
    Dim xRow As Long, xRow1 As Long, xCol As Long, xlastrow As Long
    Dim myFile As String
    
    
    Select Case myCount
    Case 0 'ACP41
        Set ws = wsSA1c
        Set ws2 = wsSA1b
        Set ws3 = wsSA1d
        Set ws4 = wsSA1e
        myFile = myPath(myCount) & Format(DateAdd("m", -1, Date), "YYYYMM") & " SPOOL " & myFileName(myCount) & ".xls"
    Case 1 'ACP52
        Set ws = wsSA2c
        Set ws2 = wsSA2b
        Set ws3 = wsSA2d
        Set ws4 = wsSA2e
        myFile = myPath(myCount) & Format(DateAdd("m", -1, Date), "YYYYMM") & " SPOOL " & myFileName(myCount) & ".xls"
    Case 2 'NTP48
        Set ws = wsSA3c
        Set ws2 = wsSA3b
        Set ws3 = wsSA3d
        Set ws4 = wsSA3e
        myFile = myPath(myCount) & Format(DateAdd("m", -1, Date), "YYYYMM") & " SPOOL " & myFileName(myCount) & ".xls"
    Case 3 'PALTA
        Set ws = wsSA4c
        Set ws2 = wsSA4b
        Set ws3 = wsSA4d
        Set ws4 = wsSA4e
        myFile = myPath(myCount) & Format(DateAdd("m", -1, Date), "YYYYMM") & " SPOOL " & myFileName(myCount) & ".xls"
    Case 4 'SEDNA CRUX
        Set ws = wsSA5c
        Set ws2 = wsSA5b
        Set ws3 = wsSA5d
        Set ws4 = wsSA5e
        myFile = myPath(myCount) & Format(DateAdd("m", -1, Date), "YYYYMM") & " SPOOL SEDNA.xls"
    Case 5 'PRELUDE
        Set ws = wsSA6c
        Set ws2 = wsSA6b
        Set ws3 = wsSA6d
        Set ws4 = wsSA6e
        myFile = myPath(myCount) & Format(DateAdd("m", -1, Date), "YYYYMM") & " SPOOL " & myFileName(myCount) & ".xls"
    Case 6 'AU01
        Set ws = wsSE1c
        Set ws2 = wsSE1b
        Set ws3 = wsSE1d
        Set ws4 = wsSE1e
        myFile = myPath(myCount) & Format(DateAdd("m", -1, Date), "YYYYMM") & " SPOOL " & myFileName(myCount) & ".xls"
    Case 7 'AU02
        Set ws = wsSE2c
        Set ws2 = wsSE2b
        Set ws3 = wsSE2d
        Set ws4 = wsSE2e
        myFile = myPath(myCount) & Format(DateAdd("m", -1, Date), "YYYYMM") & " SPOOL " & myFileName(myCount) & ".xls"
    Case 8 'AU10
        Set ws = wsSE3c
        Set ws2 = wsSE3b
        Set ws3 = wsSE3d
        Set ws4 = wsSE3e
        myFile = myPath(myCount) & Format(DateAdd("m", -1, Date), "YYYYMM") & " SPOOL " & myFileName(myCount) & ".xls"
    Case 9 'AU11
        Set ws = wsSE4c
        Set ws2 = wsSE4b
        Set ws3 = wsSE4d
        Set ws4 = wsSE4e
        myFile = myPath(myCount) & Format(DateAdd("m", -1, Date), "YYYYMM") & " SPOOL " & myFileName(myCount) & ".xls"
    End Select
    
    If Dir(myFile) = "" Then
        GoTo noFileFound
    End If
    
    Set tempWB = Workbooks.Open(myFile, ReadOnly:=True)
    Set tempWS = tempWB.Worksheets(1)
    
    tempWS.Cells.Copy ws.Range("A1")
    tempWS.Cells.Copy ws3.Range("A1")
    tempWS.Cells.Copy ws4.Range("A1")
    
    tempWB.Close False
    Set tempWB = Nothing
    Set tempWS = Nothing
    
    xRow = 1
    While ws.Range("A" & xRow + 1).Value = ""
        xRow = xRow + 1
    Wend
    ws.Range("A1:A" & xRow).EntireRow.Delete
    
    While Not ws.Range("B:B").Find(" List does not contain any data", lookat:=xlWhole) Is Nothing
        xRow = ws.Range("B:B").Find(" List does not contain any data", lookat:=xlWhole).Row
        ws.Range("A" & xRow - 8 & ":A" & xRow + 1).EntireRow.Delete
    Wend
    
    xRow = 3
    xlastrow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    While Not xRow >= xlastrow
        If ws.Range("B" & xRow).Value = "" And ws.Range("F" & xRow).Value = "" And ws.Range("H" & xRow).Value = "" And ws.Range("I" & xRow).Value = "" And ws.Range("AI" & xRow).Value = "" And ws.Range("AK" & xRow).Value = "" And ws.Range("AM" & xRow).Value = "" And ws.Range("A" & xRow).Value = "" Then
            ws.Range("B" & xRow).EntireRow.Delete
            xRow = xRow - 1
            xlastrow = xlastrow - 1
        ElseIf ws.Range("A" & xRow).Value <> "" And ws.Range("A" & xRow - 1).Value = "" And ws.Range("A" & xRow + 1).Value <> "" Then
            ws.Range("A" & xRow).EntireRow.Insert
            xRow = xRow + 1
            xlastrow = xlastrow + 1
        End If
        xRow = xRow + 1
    Wend
    
    
    
    ws2.Range("T13:T15").Font.Bold = True
    ws2.Range("T13").Value = "GL A2680008 & A3620001"
    
    ws2.Range("T14").Value = "Deductible (A2680008)"
    ws2.Range("T15").Value = "To be Paid Over (A3620001)"
    
    If Not ws.Range("H:H").Find("VST", lookat:=xlWhole) Is Nothing Then
        xRow = ws.Range("H:H").Find("VST", lookat:=xlWhole, SearchDirection:=xlPrevious).Row
        
        xRow1 = xRow
        While Not ws.Range("H" & xRow1 - 1).Value = ""
            xRow1 = xRow1 - 1
        Wend
        
        If Not ws.Range("A" & xRow1).EntireRow.Find("To be paid over", lookat:=xlPart) Is Nothing Then
            xCol = ws.Range("A" & xRow1).EntireRow.Find("To be paid over", lookat:=xlPart).Column
            ws2.Range("U15").Value = ws.Range(NumberToLetter(xCol) & xRow).Value
        End If
        
        If Not ws.Range("A" & xRow1).EntireRow.Find("Deductible", lookat:=xlPart) Is Nothing Then
            xCol = ws.Range("A" & xRow1).EntireRow.Find("Deductible", lookat:=xlPart).Column
            ws2.Range("U14").Value = ws.Range(NumberToLetter(xCol) & xRow).Value
        End If
        
    End If
    
    
    ws.Range("A1:A10").EntireRow.Insert
    
    ws.Range("A1").Value = "Rightful GST"
    ws.Range("A14").EntireRow.Copy ws.Range("A3")
    ws.Range("AT3").Value = "Effective Rate"
    ws.Range("AU3").Value = "Rightful GST"
    ws.Range("AV3").Value = "Variant"
    ws.Range("AW3").Value = "Remark"
    ws.Range("AX3").Value = "ADJUSTMENT BAS"
    
    ws.Range("AT3:AX3").Interior.Color = RGB(0, 255, 0)
    ws.Range("AT3:AX3").Font.Bold = True
    
    xRow = 4
    xRow1 = 15
    xlastrow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    While Not xRow1 > xlastrow
        If (ws.Range("AI" & xRow1).Value = "A1" Or ws.Range("AI" & xRow1).Value = "K1" Or ws.Range("AI" & xRow1).Value = "K7") And ws.Range("AC" & xRow1).Value <> "" Then
            If Abs((ws.Range("AK" & xRow1).Value * 0.1) - ws.Range("AN" & xRow1).Value) > 1 Then
                ws.Range("A" & xRow1).EntireRow.Copy
                ws.Range("A" & xRow).Insert
                Application.CutCopyMode = False
                xRow = xRow + 1
                xRow1 = xRow1 + 1
                xlastrow = xlastrow + 1
            End If
        End If
        
        xRow1 = xRow1 + 1
    Wend
    
    xRow = xRow - 1
    
    If xRow >= 4 Then
        ws.Range("AT4:AT" & xRow).Formula = "=AN4/AK4"
        ws.Range("AU4:AU" & xRow).Formula = "=IF(OR(AI4=""A1"",AI4=""K1"",AI4=""K7""),AK4*10%,0)"
        ws.Range("AV4:AV" & xRow).Formula = "=AU4-AN4"
        'ws.Range("AW4:AW" & xRow).Formula = ""
        ws.Range("AX4:AX" & xRow).Formula = "=AV4 *10"
        ws.Range("AX" & xRow + 1).Formula = "=SUBTOTAL(9,AX4:AX" & xRow & ")"
        ws.Range("AX" & xRow + 1).Interior.Color = RGB(255, 255, 0)
        With ws.Range("AX" & xRow + 1).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Range("AX" & xRow + 1).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        
        ws.Range("AT4:AT" & xRow).NumberFormat = "0%"
        ws.Range("AU4:AV" & xRow).NumberFormat = "#,##0.00"
        ws.Range("AX4:AX" & xRow + 1).NumberFormat = "#,##0.00"
    End If
    
    
    'Stop
    
noFileFound:
    
    
End Sub


Sub copyFBL3N(myCount As Long)
    Dim ws As Worksheet, ws1 As Worksheet
    Dim xRow As Long
    Dim xlastrow As Long
    Dim myTrigger As Boolean
    
    Select Case myCount
    Case 0 'ACP41
        Set ws = wsSA1d
        Set ws1 = wsSA1e
    Case 1 'ACP52
        Set ws = wsSA2d
        Set ws1 = wsSA2e
    Case 2 'NTP48
        Set ws = wsSA3d
        Set ws1 = wsSA3e
    Case 3 'PALTA
        Set ws = wsSA4d
        Set ws1 = wsSA4e
    Case 4 'SEDNA CRUX
        Set ws = wsSA5d
        Set ws1 = wsSA5e
    Case 5 'PRELUDE
        Set ws = wsSA6d
        Set ws1 = wsSA6e
    Case 6 'AU01
        Set ws = wsSE1d
        Set ws1 = wsSE1e
    Case 7 'AU02
        Set ws = wsSE2d
        Set ws1 = wsSE2e
    Case 8 'AU10
        Set ws = wsSE3d
        Set ws1 = wsSE3e
    Case 9 'AU11
        Set ws = wsSE4d
        Set ws1 = wsSE4e
    End Select
    
    If Not ws.Range("AI:AI").Find("Tx", lookat:=xlWhole) Is Nothing Then
        xRow = ws.Range("AI:AI").Find("Tx", lookat:=xlWhole).Row
        ws.Range("A1:A" & xRow - 1).EntireRow.Delete
        ws.Range("A2").EntireRow.Delete
        
        ws1.Range("A1:A" & xRow - 1).EntireRow.Delete
        ws1.Range("A2").EntireRow.Delete
        
        
        xRow = 2
        While Not xRow > Application.WorksheetFunction.Max(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, ws.Cells(ws.Rows.Count, "AI").End(xlUp).Row)
            If ws.Range("AI" & xRow).Value = "Tx" Or ws.Range("AI" & xRow).Value = "" Or ws.Range("Y" & xRow).Value = "" Then
                ws.Range("AI" & xRow).EntireRow.Delete
                xRow = xRow - 1
            ElseIf Not (ws.Range("AI" & xRow).Value = "K1" Or ws.Range("AI" & xRow).Value = "K7" Or ws.Range("AI" & xRow).Value = "A1") Then
                ws.Range("AI" & xRow).EntireRow.Delete
                xRow = xRow - 1
            End If
            
            xRow = xRow + 1
        Wend
        
        
        
        xRow = 2
        While Not xRow > Application.WorksheetFunction.Max(ws1.Cells(ws.Rows.Count, "B").End(xlUp).Row, ws1.Cells(ws.Rows.Count, "AI").End(xlUp).Row)
            If ws1.Range("AI" & xRow).Value = "Tx" Or ws1.Range("AI" & xRow).Value = "" Or ws1.Range("Y" & xRow).Value = "" Then
                ws1.Range("AI" & xRow).EntireRow.Delete
                xRow = xRow - 1
            ElseIf Not (ws1.Range("AI" & xRow).Value = "K1" Or ws1.Range("AI" & xRow).Value = "K7" Or ws1.Range("AI" & xRow).Value = "A1") Then
                ws1.Range("AI" & xRow).EntireRow.Delete
                xRow = xRow - 1
                
            ElseIf wsExclude.Range("O:O").Find(ws1.Range("Y" & xRow).Value) Is Nothing Then
                ws1.Range("AI" & xRow).EntireRow.Delete
                xRow = xRow - 1
            End If
            
            xRow = xRow + 1
        Wend
        
        If ws1.Range("A2").Value <> "" Then
            finalMessage = finalMessage & vbLf & " - " & myFileName(myCount) & ": GST Error"
        End If
        
        
        ws.Range("A:A, C:C, E:E, G:H, J:M, O:O, Q:R, T:V, X:X, Z:AB, AD:AD, AF:AG, AO:AZ").EntireColumn.Delete
        ws1.Range("A:A, C:C, E:E, G:H, J:M, O:O, Q:R, T:V, X:X, Z:AB, AD:AD, AF:AG, AO:AZ").EntireColumn.Delete
        
        'ws.Range("A:A, C:C, E:E, G:H, J:M, O:O, Q:R, T:V, X:X, Z:AB, AD:AD, AF:AH, AJ:AL, AP:AP, AR:AR, AT:AT, AV:AZ").EntireColumn.Delete
        
        ws.Range("S1").Value = "Doc. Currency"
        ws.Range("T1").Value = "Doc Exchange Rate"
        ws.Range("U1").Value = "Reporting Exchange Rate ("
        ws.Range("V1").Value = "Remarks"
        
        ws.Range("T:U").NumberFormat = "0.00000"
        
        ws.Range("A1:V1").Font.Bold = True
        ws.Range("S1:V1").Interior.Color = RGB(146, 208, 80)
        
        xRow = 2
        While Not ws.Range("A" & xRow).Value = "" Or Not ws.Range("B" & xRow).Value = "" Or Not ws.Range("C" & xRow).Value = "" Or Not ws.Range("D" & xRow).Value = ""
            If Not wsFBL3N.Range("A:A").Find(ws.Range("J" & xRow).Value) Is Nothing Then
                xlastrow = wsFBL3N.Range("A:A").Find(ws.Range("J" & xRow).Value).Row
                ws.Range("S" & xRow).Value = wsFBL3N.Range("B" & xlastrow).Value
                ws.Range("T" & xRow).Value = wsFBL3N.Range("C" & xlastrow).Value * 1
                
                Select Case ws.Range("S" & xRow).Value
                Case "AUD"
                    ws.Range("U" & xRow).Value = 1
                    ws.Range("V" & xRow).Value = "OK"
                Case "USD"
                    ws.Range("U" & xRow).Value = wsExch.Range("B2").Value
                    
                    If ws.Range("T" & xRow).Value > wsExch.Range("B36").Value Or ws.Range("T" & xRow).Value < wsExch.Range("B35").Value Then
                        ws.Range("V" & xRow).Value = "To check: value to be between " & wsExch.Range("B35").Value & " and " & wsExch.Range("B36").Value
                        ws.Range("V" & xRow).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws.Range("V" & xRow).Value = "OK"
                    End If
                    
                Case "CAD"
                    ws.Range("U" & xRow).Value = wsExch.Range("C2").Value
                    
                    If ws.Range("T" & xRow).Value > wsExch.Range("C36").Value Or ws.Range("T" & xRow).Value < wsExch.Range("C35").Value Then
                        ws.Range("V" & xRow).Value = "To check: value to be between " & wsExch.Range("C35").Value & " and " & wsExch.Range("C36").Value
                        ws.Range("V" & xRow).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws.Range("V" & xRow).Value = "OK"
                    End If
                    
                Case "EUR"
                    ws.Range("U" & xRow).Value = wsExch.Range("D2").Value
                    
                    If ws.Range("T" & xRow).Value > wsExch.Range("D36").Value Or ws.Range("T" & xRow).Value < wsExch.Range("D35").Value Then
                        ws.Range("V" & xRow).Value = "To check: value to be between " & wsExch.Range("D35").Value & " and " & wsExch.Range("D36").Value
                        ws.Range("V" & xRow).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws.Range("V" & xRow).Value = "OK"
                    End If
                    
                Case "KRW"
                    ws.Range("U" & xRow).Value = wsExch.Range("E2").Value
                    
                    If ws.Range("T" & xRow).Value > wsExch.Range("E36").Value Or ws.Range("T" & xRow).Value < wsExch.Range("E35").Value Then
                        ws.Range("V" & xRow).Value = "To check: value to be between " & wsExch.Range("E35").Value & " and " & wsExch.Range("E36").Value
                        ws.Range("V" & xRow).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws.Range("V" & xRow).Value = "OK"
                    End If
                    
                Case "NOK"
                    ws.Range("U" & xRow).Value = wsExch.Range("F2").Value
                    
                    If ws.Range("T" & xRow).Value > wsExch.Range("F36").Value Or ws.Range("T" & xRow).Value < wsExch.Range("F35").Value Then
                        ws.Range("V" & xRow).Value = "To check: value to be between " & wsExch.Range("F35").Value & " and " & wsExch.Range("F36").Value
                        ws.Range("V" & xRow).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws.Range("V" & xRow).Value = "OK"
                    End If
                    
                Case "GBP"
                    ws.Range("U" & xRow).Value = wsExch.Range("G2").Value
                    
                    If ws.Range("T" & xRow).Value > wsExch.Range("G36").Value Or ws.Range("T" & xRow).Value < wsExch.Range("G35").Value Then
                        ws.Range("V" & xRow).Value = "To check: value to be between " & wsExch.Range("G35").Value & " and " & wsExch.Range("G36").Value
                        ws.Range("V" & xRow).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws.Range("V" & xRow).Value = "OK"
                    End If
                    
                Case "JPY"
                    ws.Range("U" & xRow).Value = wsExch.Range("H2").Value
                    
                    If ws.Range("T" & xRow).Value > wsExch.Range("H36").Value Or ws.Range("T" & xRow).Value < wsExch.Range("H35").Value Then
                        ws.Range("V" & xRow).Value = "To check: value to be between " & wsExch.Range("H35").Value & " and " & wsExch.Range("H36").Value
                        ws.Range("V" & xRow).Interior.Color = RGB(255, 0, 0)
                    Else
                        ws.Range("V" & xRow).Value = "OK"
                    End If
                    
                End Select
                
                
                
            End If
            
            xRow = xRow + 1
        Wend
        
    Else
        ws.Cells.ClearContents
        
    End If
    
End Sub













Sub doUpdateData(ws_Data As Worksheet, ws_Ref As Worksheet, ws_Exc As Worksheet, my_Count As Long)
    Dim xlastrow As Long
    Dim myVal As Double
    myVal = 0
    
    
    
    With ws_Data
        '.Range("G1").Value = Month(DateAdd("m", -1, Date))
        '.Range("H1").Value = Year(DateAdd("m", -1, Date))
        '.Range("H3").Value = Date
        
        'if net position is 0 then Yes, otherwise put 0
        .Range("H6").Formula = "=IF(ROUND(J32,0)=0,""YES"",""NO"")"
    
        .Range("C9").Formula = "='Supporting Data'!B9"
        .Range("C10").Formula = "='Supporting Data'!B10"
        .Range("C11").Formula = "='Supporting Data'!B11"
        .Range("C12").Formula = "='Supporting Data'!B12"
                
        .Range("C15").Formula = "='Supporting Data'!B15"
        
        .Range("C19").Formula = "='Supporting Data'!B18"
        .Range("C20").Formula = "='Supporting Data'!B19"
        
        .Range("C22").Formula = "='Supporting Data'!B21"
        .Range("C23").Formula = "='Supporting Data'!B22"
        .Range("C24").Formula = "='Supporting Data'!B23"
        
        .Range("C27").Formula = "='Supporting Data'!B26"
        
        If my_Count >= 6 Then
            .Range("C38").Formula = "='Supporting Data'!K" & Month(DateAdd("m", -1, Date)) + 4
            .Range("C45").Formula = "='Supporting Data'!Q" & Month(DateAdd("m", -1, Date)) + 4
        Else
            .Range("C38").Value = "='Supporting Data'!U15"
            .Range("C45").Value = "='Supporting Data'!U14"
        End If
        
        .Range("G9").Value = ""
        .Range("G20").Value = ""
        
        
        'If my_Count = 0 Then
        '    .Range("G56").Value = .Range("J32").Value
        'ElseIf my_Count = 1 Then
        '    .Range("G61").Value = .Range("J32").Value
        'ElseIf my_Count = 2 Then
        '    .Range("G57").Value = .Range("J32").Value
        'Else
        If my_Count = 3 Then
            .Range("G20").Formula = "=-C20"
            .Range("F45").Formula = "=-F42"
        '    .Range("G58").Value = .Range("J32").Value
        'ElseIf my_Count = 4 Then
        '    .Range("G60").Value = .Range("J32").Value
        ElseIf my_Count = 5 Then
            xlastrow = 4
            If ws_Exc.Range("AX4").Value = "" Then
            Else
                While Not ws_Exc.Range("AX" & xlastrow + 1).Value = ""
                    xlastrow = xlastrow + 1
                Wend
                
                .Range("G20").Formula = "='Exception Rule'!AX" & xlastrow
            End If
        
        
        '    .Range("G59").Value = .Range("J32").Value
        ElseIf my_Count = 6 Then
            .Range("G9").Formula = "='Supporting Data'!V5"
            
            xlastrow = 4
            While Not ws_Exc.Range("AX" & xlastrow + 1).Value = ""
                xlastrow = xlastrow + 1
            Wend
            .Range("G20").Formula = "='Exception Rule'!AX" & xlastrow & " - 'Supporting Data'!U5 - " & wsSA4a.Range("G20").Value
            
            '.Range("G20").Formula = "= 0 -'Supporting Data'!U5 - " & wsSA4a.Range("G20").Value
            
            .Range("G56").Value = wsSA1a.Range("J32").Value
            .Range("G57").Value = wsSA3a.Range("J32").Value
            .Range("G58").Value = wsSA4a.Range("J32").Value
            .Range("G59").Value = wsSA6a.Range("J32").Value
            .Range("G60").Value = wsSA5a.Range("J32").Value
            .Range("G61").Value = wsSA2a.Range("J32").Value
            
            .Range("G86").Value = wsSA6a.Range("G49").Value * -1
            
            .Range("G91").Value = Me.tb_GSTPayable.Value
            
            '.Range("G62").Value = ws_Ref.Range("U4").Value 'wsSA1b.Range("").Value
        End If
        
        
        
        
        
        
        'SAPL each one
        'get the net position
        '.Range("G56").Value = ""
        '.Range("G57").Value = ""
        '.Range("G58").Value = ""
        '.Range("G59").Value = ""
        '.Range("G60").Value = ""
        '.Range("G61").Value = ""
        '.Range("G62").Value = ""
        
        
        'AU01 only
        'fleetplus output adjustment BAS
        '.Range("G9").Value = ""
        
        '.Range("G20").Value = ""
        
        
        '.Range("C9").Formula = "='Supporting Data'!B9"
        '.Range("C10").Formula = "='Supporting Data'!B10"
        '.Range("C11").Formula = "='Supporting Data'!B11"
        '.Range("C12").Formula = "='Supporting Data'!B12"
        '
        '.Range("C15").Formula = "='Supporting Data'!B15"
       '
        '.Range("C19").Formula = "='Supporting Data'!B18"
        '.Range("C20").Formula = "='Supporting Data'!B19"
       '
       ' .Range("C22").Formula = "='Supporting Data'!B21"
       ' .Range("C23").Formula = "='Supporting Data'!B22"
       ' .Range("C24").Formula = "='Supporting Data'!B23"
       '
       ' .Range("C27").Formula = "='Supporting Data'!B26"
        
       '
        
        
        'for PALTA, G20 = negative of C20
        '.Range("G20").Value = .Range("C20").Value * -1
        
        'only for AU01
        'need to calculate K1 (10% tax) if its not charged, the unpaid amount minus Fleetplus Input minus PALTA input form (normally positive number so it subtract)
        '.Range("G20").Value = ""
        
        'take from spool = to be paid over, take the value same row as the last VST
        'take value terus
        '.Range("C38").Value = ""
        
        'take from spool = deductible, take the value same row as the last VST
        'take value terus
        '.Range("C45").Value = ""
        
        'only for prelude only
        'get data from using
        'negative the value
        '.Range("G78").Value = "-"
        
        
        'SEHAL AU01 only
        'get value from email
        'total GST =  input
        'employee contribution = output
        'total gst payable = total claimable
        '.Range("G79").Value = ""
        
        
        
    End With
    
    
    
    
    
    
    
End Sub


Sub doCrossCheck()
    Dim myCount As Long
    
    
    With wsCrosscheck
        .Range("G17:G24").Value = 0
        .Range("G27:G37").Value = 0
    
        .Range("J17:J24").Value = 0
        .Range("J27:J37").Value = 0
    
        .Range("AE17:AI24").Value = 0
        .Range("AE27:AI37").Value = 0
        
        'ACP41
        'Debug.Print wsSA1a.Parent.Name & "    " & wsSA1a.Name
        
        wsSA1a.Calculate
        wsSA2a.Calculate
        wsSA3a.Calculate
        wsSA4a.Calculate
        wsSA5a.Calculate
        wsSA6a.Calculate
        wsSE1a.Calculate
        wsSE2a.Calculate
        wsSE3a.Calculate
        wsSE4a.Calculate
        
        wsSA1a.Range("J9:J29").Copy
        .Range("AQ17").PasteSpecial xlPasteValues
        'ACP52
        wsSA2a.Range("J9:J29").Copy
        .Range("AV17").PasteSpecial xlPasteValues
        'NTP48
        wsSA3a.Range("J9:J29").Copy
        .Range("AS17").PasteSpecial xlPasteValues
        'PALTA
        wsSA4a.Range("J9:J29").Copy
        .Range("AR17").PasteSpecial xlPasteValues
        'SEDNA CRUX
        wsSA5a.Range("J9:J29").Copy
        .Range("AT17").PasteSpecial xlPasteValues
        'PRELUDE
        wsSA6a.Range("J9:J29").Copy
        .Range("AU17").PasteSpecial xlPasteValues
        'AU01
        wsSE1a.Range("J9:J29").Copy
        .Range("Y17").PasteSpecial xlPasteValues
        'AU02
        wsSE2a.Range("J9:J29").Copy
        .Range("AB17").PasteSpecial xlPasteValues
        'AU10
        wsSE3a.Range("J9:J29").Copy
        .Range("AM17").PasteSpecial xlPasteValues
        'AU11
        wsSE4a.Range("J9:J29").Copy
        .Range("AJ17").PasteSpecial xlPasteValues
            
        Application.CutCopyMode = False
        
    End With
    
    
    wbCrosscheck.SaveAs Path_Main & "GST Crosscheck_" & Format(DateAdd("m", -1, Date), "yyyy_mm_mmmm") & "_SETL_SAPLJV_SEHAL.xlsb"
    wbCrosscheck.Close True
    
    
End Sub





Sub doSP()
    Dim myFolderFrom As String, myFrom As String
    Dim myFolderTo As String, myTo As String
    Dim myFileName As String
    Dim myCount As Long, myCol As Long
    
    
    Dim myWB As Workbook, myWS As Worksheet
    
    Me.LabelCaption.Caption = "Updating GST Crosscheck File in SP"
    
    myFolderFrom = "\\eu001-sp.shell.com\sites\AAFAA1747\GST Australia " & Year(DateAdd("m", -2, Date)) & "\"
    If Dir(myFolderFrom, vbDirectory) = "" Then
        myFolderFrom = "\\eu001-sp.shell.com\sites\AAFAA1747\GST Australia  " & Year(DateAdd("m", -2, Date)) & "\"
        If Dir(myFolderFrom, vbDirectory) = "" Then
            finalMessage = finalMessage & vbLf & " - SharePoint site (Y) unaccessible/ not found"
            Exit Sub
        End If
    End If
    
    myFrom = myFolderFrom & "AU_" & Format(DateAdd("m", -2, Date), "mmyyyy") & "_C6a.6.d.1_GST Crosscheck_" & Format(DateAdd("m", -2, Date), "mmmm yyyy") & "\"
    
    If Dir(myFrom, vbDirectory) = "" Then
        finalMessage = finalMessage & vbLf & " - SharePoint site (M) unaccessible/ not found"
        Exit Sub
    End If
    
    myFileName = "GST Crosscheck_" & Format(DateAdd("m", -2, Date), "YYYY_MM_MMMM") & "_SETL_SAPLJV_SEHAL.xlsb"
    
    If Dir(myFrom & myFileName) = "" Then
        finalMessage = finalMessage & vbLf & " - File in SP not found"
        Exit Sub
    End If
    
    
    If Month(Date) = 2 Then
        myFolderTo = "\\eu001-sp.shell.com\sites\AAFAA1747\GST Australia " & Year(DateAdd("m", -1, Date)) & "\"
    Else
        myFolderTo = myFolderFrom
    End If
    
    If Dir(myFolderTo, vbDirectory) = "" Then
        MkDir myFolderTo
    End If
    
    myTo = myFolderTo & "AU_" & Format(DateAdd("m", -1, Date), "mmyyyy") & "_C6a.6.d.1_GST Crosscheck_" & Format(DateAdd("m", -1, Date), "mmmm yyyy") & "\"
    If Dir(myTo, vbDirectory) = "" Then
        MkDir myTo
    End If
    
    
    
    Set myWB = Workbooks.Open(myFrom & myFileName)
    
    On Error GoTo noWS
    Set myWS = myWB.Worksheets("GST Cross Check")
    On Error GoTo 0
    
    myCount = 1
    myFileName = myTo & "GST Crosscheck_" & Format(DateAdd("m", -1, Date), "YYYY_MM_MMMM") & "_SETL_SAPLJV_SEHAL.xlsb"
    While Dir(myFileName) <> ""
        myFileName = myTo & "GST Crosscheck_" & Format(DateAdd("m", -1, Date), "YYYY_MM_MMMM") & "_SETL_SAPLJV_SEHAL (" & myCount & ").xlsb"
        myCount = myCount + 1
    Wend
    
    
    myWS.Range("G17:G24").Value = 0
    myWS.Range("G27:G37").Value = 0
    
    myWS.Range("J17:J24").Value = 0
    myWS.Range("J27:J37").Value = 0
    
    myWS.Range("AE17:AI24").Value = 0
    myWS.Range("AE27:AI37").Value = 0
    
    
    For myCount = 0 To 9
        Select Case myCount
        Case 0 'ACP41
            wsSA1a.Range("J9:J29").Copy
            myWS.Range("AQ17").PasteSpecial xlPasteValues
        Case 1 'ACP52
            wsSA2a.Range("J9:J29").Copy
            myWS.Range("AV17").PasteSpecial xlPasteValues
        Case 2 'NTP48
            wsSA3a.Range("J9:J29").Copy
            myWS.Range("AS17").PasteSpecial xlPasteValues
        Case 3 'PALTA
            wsSA4a.Range("J9:J29").Copy
            myWS.Range("AR17").PasteSpecial xlPasteValues
        Case 4 'SEDNA CRUX
            wsSA5a.Range("J9:J29").Copy
            myWS.Range("AT17").PasteSpecial xlPasteValues
        Case 5 'PRELUDE
            wsSA6a.Range("J9:J29").Copy
            myWS.Range("AU17").PasteSpecial xlPasteValues
        Case 6 'AU01
            wsSE1a.Range("J9:J29").Copy
            myWS.Range("Y17").PasteSpecial xlPasteValues
        Case 7 'AU02
            wsSE2a.Range("J9:J29").Copy
            myWS.Range("AB17").PasteSpecial xlPasteValues
        Case 8 'AU10
            wsSE3a.Range("J9:J29").Copy
            myWS.Range("AM17").PasteSpecial xlPasteValues
        Case 9 'AU11
            wsSE4a.Range("J9:J29").Copy
            myWS.Range("AJ17").PasteSpecial xlPasteValues
        End Select
        
        Application.CutCopyMode = False
    Next
    
    myWB.SaveCopyAs myFileName
    myWB.Close False
    
    Exit Sub
    
noWS:
    myWB.Close False
    finalMessage = finalMessage & vbLf & " - WS error for File in SP"
    
End Sub



















































Sub closeFile()
    Dim wb As Workbook
    
    Call ChangeProgress("Closing all working files", 0.99)
    
    wbSA1.Close True
    wbSA2.Close True
    wbSA3.Close True
    wbSA4.Close True
    wbSA5.Close True
    wbSA6.Close True
    wbSE1.Close True
    wbSE2.Close True
    wbSE3.Close True
    wbSE4.Close True
    
    wbFBL3N.Close True
    
End Sub


'**************************************************************************************************************
'SAP Scripts
'**************************************************************************************************************

Sub SAP_BP_Clear()
    Call ChangeProgress("Clearing SM37", 0.05)

    
    Dim xCount As Long
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "SM37"
    sessBP.findById("wnd[0]").sendVKey 0
    
    sessBP.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Debug.Print sessBP.findById("wnd[0]/usr/lbl[64,13]").Text
    
    '******************************
    'tested that the data goes from 13 to 23 (11 rows), might change, need to further check

doAgain:
    
    xCount = 13
    While Not sessBP.findById("wnd[0]/usr/lbl[64," & xCount & "]", False) Is Nothing
        sessBP.findById("wnd[0]/usr/chk[1," & xCount & "]").Selected = True
        xCount = xCount + 1
    Wend
    
    If xCount > 13 Then
        sessBP.findById("wnd[0]/tbar[1]/btn[14]").press
        sessBP.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        GoTo doAgain
    End If
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press

    
    
End Sub


Sub SAP_BP_1() '(AU10_Range As Range)
    
    Dim qCount As Long
    Dim spoolName As String
    Dim xRow As Long
    Dim setFormat As Long
    
    
    Call ChangeProgress("Running ZFR_GJVA", 0.05)
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
    
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "ZFR_GJVA"
    sessBP.findById("wnd[0]").sendVKey 0
    
    sessBP.findById("wnd[0]/usr/btnPUSHB_O1").press 'show further selection
    sessBP.findById("wnd[0]/usr/btnPUSHB_O2").press 'show tax payable posting
    sessBP.findById("wnd[0]/usr/btnPUSHB_O3").press 'output control
    sessBP.findById("wnd[0]/usr/btnPUSHB_O4").press 'output lists
    sessBP.findById("wnd[0]/usr/btnPUSHB_O5").press 'posting parameters
    
    sessBP.findById("wnd[0]/usr/txtBR_GJAHR-LOW").Text = Year(DateAdd("m", -1, Date))
    sessBP.findById("wnd[0]/usr/ctxtBR_BUDAT-LOW").Text = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "dd.mm.yyyy")
    sessBP.findById("wnd[0]/usr/ctxtBR_BUDAT-HIGH").Text = Format(DateSerial(Year(Date), Month(Date), 1 - 1), "dd.mm.yyyy")
    
    sessBP.findById("wnd[0]/usr/chkPAR_XADR").Selected = True
    sessBP.findById("wnd[0]/usr/chkPAR_UMSV").Selected = True
    sessBP.findById("wnd[0]/usr/ctxtPAR_LAUD").Text = Format(Date, "dd.mm.yyyy")
    
    
    For qCount = 1 To 11
        Select Case qCount
        Case 1
            'ACP41
            
            Call ChangeProgress("Running ZFR_GJVA - ACP41", 0.07)
            
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU01"
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "ACP41"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "ACP41" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            Wend
            
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "Y11001"
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "B").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("B4:B" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL ACP41"
            
            
        Case 2
            'ACP52
            
            Call ChangeProgress("Running ZFR_GJVA - ACP52", 0.09)
            
            
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU01"
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "ACP52"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "ACP52" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            Wend
            
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "Y11017"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "Y11018"
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "C").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("C4:C" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL ACP52"
        Case 3
            'NTP48
            
            Call ChangeProgress("Running ZFR_GJVA - NTP48", 0.011)
            
            
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU01"
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "NTP48"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "NTP48" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            Wend
            
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "Y11011"
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "D").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("D4:D" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL NTP48"
            
        Case 4
            'PALTA
            
            Call ChangeProgress("Running ZFR_GJVA - PALTA", 0.13)
            
            
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU01"
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "PALTA"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "PALTA" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            Wend
            
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "Y11010"
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "E").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("E4:E" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL PALTA"
            
        Case 5
            'SEDNA
            
            Call ChangeProgress("Running ZFR_GJVA - SEDNA", 0.15)
            
            
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU01"
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "SEDNA"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "SEDNA" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            Wend
            
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "Y11012"
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "F").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("F4:F" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL SEDNA"
            
        Case 6
            'PRELU
            
            Call ChangeProgress("Running ZFR_GJVA - PRELUDE", 0.17)
            
            
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU01"
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "PRELU"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "PRELU" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            Wend
            
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "Y11008"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "Y11009"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "Y11099"
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "G").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("G4:G" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL PRELUDE"
            
        Case 7
            'AU01
            
            Call ChangeProgress("Running ZFR_GJVA - AU01", 0.19)
            
            
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU01"
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "AU01"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "AU01" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            Wend
            
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "Y11001"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = "Y11011"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").Text = "Y11002"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").Text = "Y11008"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").Text = "Y11009"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,4]").Text = "Y11099"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,5]").Text = "Y11010"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,6]").Text = "Y11012"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,7]").Text = "Y11017"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E").verticalScrollbar.Position = 1
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,7]").Text = "Y11018"
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "H").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("H4:H" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL AU01"
            
            sessBP.findById("wnd[0]/tbar[1]/btn[16]").press
            sessBP.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
            
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "DZ"
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
                        
            
            
        Case 8
            'AU02
            
            Call ChangeProgress("Running ZFR_GJVA - AU02", 0.21)
            
            
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU02"
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "AU02"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "AU02" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            Wend
            
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/chkPAR_XADR").Selected = False
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "I").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("I4:I" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL AU02"
            
            'sessBP.findById("wnd[0]/tbar[1]/btn[16]").press
            sessBP.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN001-LOW").Text = ""
            
            
            
            
        Case 9
            'AU10
            
            Call ChangeProgress("Running ZFR_GJVA - AU10", 0.23)
            
            
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU10"
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "AU10"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "AU10" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            Wend
            
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            'sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            'AU10_Range.Copy
            'sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/chkPAR_XADR").Selected = False
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "J").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("J4:J" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL AU10"
            
        Case 10
            'AU11
            
            Call ChangeProgress("Running ZFR_GJVA - AU11", 0.25)
            
            
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU11"
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "AU11"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "AU11" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            Wend
            
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/chkPAR_XADR").Selected = False
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "K").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("K4:K" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL AU11"
            
        Case 11
            'AU0309
            
            Call ChangeProgress("Running ZFR_GJVA - AU03-AU09", 0.27)
            
            
            sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "AU0309"
            sessBP.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "AU03"
            myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BUKRS_%_APP_%-VALU_PUSH").press
            setFormat = 0
            While Left(sessBP.findById("wnd[0]/sbar").Text, 33) = "This run ID has already been used"
                setFormat = setFormat + 1
                sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = "AU039" & setFormat
                myType(qCount - 1) = sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text
                
                sessBP.findById("wnd[0]/usr/btn%_BR_BUKRS_%_APP_%-VALU_PUSH").press
            Wend
            
            
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "AU03"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "AU06"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "AU07"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "AU08"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "AU09"
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            sessBP.findById("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/tbar[0]/btn[16]").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "A").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("A4:A" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            xRow = wsExclude.Cells(wsExclude.Rows.Count, "L").End(xlUp).Row
            If xRow > 3 Then
                wsExclude.Range("L4:L" & xRow).Copy
                sessBP.findById("wnd[1]/tbar[0]/btn[24]").press
            End If
            
            
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            
            spoolName = "SPOOL AU0309"
        End Select
        
        
        
        
        If qCount = 1 Then
            For setFormat = 1 To 2
                
                If setFormat = 1 Then
                    sessBP.findById("wnd[0]/usr/btn%P028130_1000").press
                Else
                    sessBP.findById("wnd[0]/usr/btn%P028146_1000").press
                End If
            
                sessBP.findById("wnd[0]/tbar[1]/btn[32]").press
                sessBP.findById("wnd[1]/usr/btnAPP_FL_ALL").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(1).Selected = True
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(2).Selected = True
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(3).Selected = True
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(5).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(1).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 12
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(15).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(14).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 0
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(5).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 12
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(16).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 0
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(2).Selected = True
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(3).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 36
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(45).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 0
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(3).Selected = True
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(5).Selected = True
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(7).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 36
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(40).Selected = True
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(41).Selected = True
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(42).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 0
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(7).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 36
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(41).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 0
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(6).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(4).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 36
                sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(37).Selected = True
                sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
                sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
                
                sessBP.findById("wnd[0]/tbar[1]/btn[34]").press
                sessBP.findById("wnd[1]/usr/ctxtLTDX-VARIANT").Text = "NAS"
                sessBP.findById("wnd[1]/usr/txtLTDXT-TEXT").Text = "NAS Automation"
                sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
                
                'warning window: activewindow.name = wnd[2]
                'main SAP page: activewindow.name = wnd[0]
                If sessBP.ActiveWindow.Name = "wnd[2]" Then
                    sessBP.findById("wnd[2]/tbar[0]/btn[0]").press
                End If
                
                sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
            Next
            
            sessBP.findById("wnd[0]/usr/ctxtPAR_VAR1").Text = "NAS"
            sessBP.findById("wnd[0]/usr/ctxtPAR_VAR3").Text = "NAS"
        End If
        
        
        
        sessBP.findById("wnd[0]/mbar/menu[0]/menu[2]").Select 'set download schedule
        
        sessBP.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = spoolName
        sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
        sessBP.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCL"
        sessBP.findById("wnd[1]/tbar[0]/btn[13]").press
        sessBP.findById("wnd[1]/usr/btnSOFORT_PUSH").press
        sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
        sessBP.findById("wnd[1]/tbar[0]/btn[11]").press
        
        
    Next
    
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
    
End Sub


Sub SAP_BP_2()
    'Download BAS Report
    Dim myName As String
    
    Call ChangeProgress("Running S_ALR_87009913", 0.3)
    
            
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "S_ALR_87009913"
    sessBP.findById("wnd[0]").sendVKey 0
    
    sessBP.findById("wnd[0]/usr/ctxtPAR_LAUD").Text = Format(Date, "dd.mm.yyyy")
    sessBP.findById("wnd[0]/usr/chkPAR_NORO").Selected = True 'do not round
    
    For xCount = 0 To 10
        sessBP.findById("wnd[0]/usr/ctxtPAR_LAUI").Text = myType(xCount)
        
        Select Case xCount
        Case 0
            
            Call ChangeProgress("Running S_ALR_87009913 - ACP41", 0.3)
    
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU01"
            myName = "ACP41"
        Case 1
            
            Call ChangeProgress("Running S_ALR_87009913 - ACP52", 0.31)
            
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU01"
            myName = "ACP52"
        Case 2
            
            Call ChangeProgress("Running S_ALR_87009913 - NTP48", 0.32)
            
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU01"
            myName = "NTP48"
        Case 3
            
            Call ChangeProgress("Running S_ALR_87009913 - PALTA", 0.33)
            
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU01"
            myName = "PALTA"
        Case 4
            
            Call ChangeProgress("Running S_ALR_87009913 - SEDNA", 0.34)
            
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU01"
            myName = "SEDNA"
        Case 5
            
            Call ChangeProgress("Running S_ALR_87009913 - PRELUDE", 0.35)
            
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU01"
            myName = "PRELUDE"
        Case 6
            
            Call ChangeProgress("Running S_ALR_87009913 - AU01", 0.36)
            
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU01"
            myName = "AU01"
        Case 7
            
            Call ChangeProgress("Running S_ALR_87009913 - AU02", 0.37)
            
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU02"
            myName = "AU02"
        Case 8
            
            Call ChangeProgress("Running S_ALR_87009913 - AU10", 0.38)
            
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU10"
            myName = "AU10"
        Case 9
            
            Call ChangeProgress("Running S_ALR_87009913 - AU11", 0.39)
            
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU11"
            myName = "AU11"
        Case 10
            
            Call ChangeProgress("Running S_ALR_87009913 - AU03-AU09", 0.4)
            
            sessBP.findById("wnd[0]/usr/ctxtSEL_BKRS-LOW").Text = "AU03"
            sessBP.findById("wnd[0]/usr/btn%_SEL_BKRS_%_APP_%-VALU_PUSH").press
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "AU03"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "AU06"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "AU07"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "AU08"
            sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "AU09"
            sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
            myName = "AU0309"
        End Select
        
        sessBP.findById("wnd[0]/mbar/menu[0]/menu[2]").Select
        
        
        If Left(sessBP.findById("wnd[0]/sbar").Text, 21) = "No data found for run" Then
            finalMessage = finalMessage & vbLf & " - " & myName & ": no BAS data"
            GoTo nextItem
        ElseIf sessBP.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]", False) Is Nothing Then
            sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
        End If
        
        
        sessBP.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "BAS " & myName
        sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
        sessBP.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCL"
        sessBP.findById("wnd[1]/tbar[0]/btn[13]").press
        sessBP.findById("wnd[1]/usr/btnSOFORT_PUSH").press
        sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
        sessBP.findById("wnd[1]/tbar[0]/btn[11]").press
        
nextItem:
    Next
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
End Sub




Sub SAP_BP_3()
    'Extract in Excel form and PDF form
    Dim myFolder As String, myFileName As String
    
    Dim myPerc As Double
    myPerc = 0.4
    
    Dim isSpool As Boolean, spoolFormat As Boolean
    
    spoolFormat = False
    
    Call ChangeProgress("Running SM37", 0.41)
    
    
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "SM37"
    sessBP.findById("wnd[0]").sendVKey 0
    
    sessBP.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Debug.Print sessBP.findById("wnd[0]/usr/lbl[64,13]").Text
    
    '******************************
    'tested that the data goes from 13 to 23 (11 rows), might change, need to further check

    xCount = 13
    While Not sessBP.findById("wnd[0]/usr/lbl[64," & xCount & "]", False) Is Nothing
    
        'ensure it's completed
        While Not sessBP.findById("wnd[0]/usr/lbl[64," & xCount & "]").Text = "Finished"
            If sessBP.findById("wnd[0]/usr/lbl[64," & xCount & "]").Text = "Canceled" Then
             GoTo nextItem
            End If
            Application.Wait (Now + TimeValue("0:00:05"))
            sessBP.findById("wnd[0]").sendVKey 8
        Wend
        
        'Debug.Print sessBP.findById("wnd[0]/usr/lbl[64," & xCount & "]").Text
        myFileName = Format(DateAdd("m", -1, Date), "YYYYMM") & " " & sessBP.findById("wnd[0]/usr/lbl[4," & xCount & "]").Text & ".xls"
        
        
        myPerc = myPerc + 0.005
        Call ChangeProgress("Running SM37 - " & sessBP.findById("wnd[0]/usr/lbl[4," & xCount & "]").Text, myPerc)
        
        
        Select Case sessBP.findById("wnd[0]/usr/lbl[4," & xCount & "]").Text 'Right(sessBP.findById("wnd[0]/usr/lbl[4," & xCount & "]").Text, Len(sessBP.findById("wnd[0]/usr/lbl[4," & xCount & "]").Text) - 6)
        Case "SPOOL ACP41", "BAS ACP41"
            mySeq = 0
        Case "SPOOL ACP52", "BAS ACP52"
            mySeq = 1
        Case "SPOOL NTP48", "BAS NTP48"
            mySeq = 2
        Case "SPOOL PALTA", "BAS PALTA"
            mySeq = 3
        Case "SPOOL SEDNA", "BAS SEDNA"
            mySeq = 4
        Case "SPOOL PRELUDE", "BAS PRELUDE"
            mySeq = 5
        Case "SPOOL AU01", "BAS AU01"
            mySeq = 6
        Case "SPOOL AU02", "BAS AU02"
            mySeq = 7
        Case "SPOOL AU10", "BAS AU10"
            mySeq = 8
        Case "SPOOL AU11", "BAS AU11"
            mySeq = 9
        Case "SPOOL AU0309", "BAS AU0309"
            mySeq = 10
        End Select
        
        
        If Left(sessBP.findById("wnd[0]/usr/lbl[4," & xCount & "]").Text, 5) = "SPOOL" Then
            isSpool = True
        Else
            isSpool = False
        End If
        
        myFolder = myPath(mySeq) 'setting to specific folder
        
        
        If xCount > 13 Then
            sessBP.findById("wnd[0]/usr/chk[1," & xCount - 1 & "]").Selected = False 'unselecting previous spool
        End If
        
        
        sessBP.findById("wnd[0]/usr/chk[1," & xCount & "]").Selected = True 'selecting the particular spool
        sessBP.findById("wnd[0]/tbar[1]/btn[44]").press 'open spool
        
        
        
        If isSpool = True Then
            If sessBP.findById("wnd[0]/sbar").Text = "No list exists" Then
                finalMessage = finalMessage & vbLf & " - " & sessBP.findById("wnd[0]/usr/lbl[4," & xCount & "]").Text & ": no spool data"
                GoTo nextItem
            End If
            
            
            
            'spoolNo(mySeq) = sessBP.findById("wnd[0]/usr/lbl[3,3]").Text
            
            'If Not sessBP.findById("wnd[0]/usr/chk[1,4]", False) Is Nothing Then
            '    spoolNo(mySeq) = sessBP.findById("wnd[0]/usr/lbl[3,4]").Text
            '    sessBP.findById("wnd[0]/usr/chk[1,4]").Selected = True
            'Else
                spoolNo(mySeq) = sessBP.findById("wnd[0]/usr/lbl[3,3]").Text
                sessBP.findById("wnd[0]/usr/chk[1,3]").Selected = True
            'End If
            
            
            sessBP.findById("wnd[0]").sendVKey 6 'SELECT ABAP
            
            If spoolFormat = False Then
                sessBP.findById("wnd[0]/tbar[1]/btn[46]").press
                sessBP.findById("wnd[1]/usr/txtDIS_TO").Text = "999"
                sessBP.findById("wnd[1]/usr/txtDIS_TO").SetFocus
                sessBP.findById("wnd[1]/usr/txtDIS_TO").caretPosition = 3
                sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
                sessBP.findById("wnd[2]/tbar[0]/btn[0]").press
                spoolFormat = True
            End If
            
        Else
            If sessBP.findById("wnd[0]/sbar").Text = "No list exists" Then
                finalMessage = finalMessage & vbLf & " - " & sessBP.findById("wnd[0]/usr/lbl[4," & xCount & "]").Text & ": no BAS data"
                GoTo nextItem
            End If
            
            sessBP.findById("wnd[0]/usr/lbl[5,3]").SetFocus
            sessBP.findById("wnd[0]/usr/lbl[5,3]").caretPosition = 13
            sessBP.findById("wnd[0]/tbar[1]/btn[94]").press
            'sessBP.findById("wnd[0]/usr/chk[1,4]").Selected = True
            
            BacNo(mySeq) = sessBP.findById("wnd[0]/usr/lbl[3,4]").Text
            
            
            '****************************
            'got error, cannot save document
            
            sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
            sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
            GoTo nextItem
            
            'sessBP.findById("wnd[0]").sendVKey 6 'SELECT ABAP
            
            '****************************
            
            
            
            
            
            
        End If
        
        sessBP.findById("wnd[0]/mbar/menu[5]/menu[5]/menu[2]/menu[1]").Select 'system - list - save - local file
        sessBP.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select 'text with tab
        sessBP.findById("wnd[1]/tbar[0]/btn[0]").press 'enter
        sessBP.findById("wnd[1]/usr/ctxtDY_PATH").Text = myFolder
        sessBP.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = myFileName
        sessBP.findById("wnd[1]/tbar[0]/btn[11]").press 'replace file
        
        sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
        sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
        
        If isSpool = False Then
            sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
        End If
        
nextItem:
        xCount = xCount + 1
    Wend
    
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
    
End Sub



Sub SAP_BP_4()
    'download pdf
    Dim myPerc As Double
    myPerc = 0.53
    
    Call ChangeProgress("Running ZFU_SPL_TO_PDF", 0.53)
    
    
    'sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    'sessBP.findById("wnd[0]").sendVKey 0
    
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "ZFU_SPL_TO_PDF"
    sessBP.findById("wnd[0]").sendVKey 0
    
    For xCount = 0 To 10
        myPerc = myPerc + 0.05
        Call ChangeProgress("Running ZFU_SPL_TO_PDF - Spool " & myFileName(xCount), myPerc)
        
        If spoolNo(xCount) <> "" Then
            sessBP.findById("wnd[0]/usr/txtSPOOLNO").Text = spoolNo(xCount)
            'sessBP.findById("wnd[0]/usr/ctxtP_FILE").Text = myPath(xCount) & myFileName(xCount) & ".pdf"
            sessBP.findById("wnd[0]/tbar[1]/btn[8]").press
            sessBP.findById("wnd[1]/usr/ctxtDY_PATH").Text = myPath(xCount)
            sessBP.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "SPOOL " & myFileName(xCount) & ".pdf"
            sessBP.findById("wnd[1]/tbar[0]/btn[11]").press
            sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
        End If
    Next
    
    For xCount = 0 To 10
        myPerc = myPerc + 0.05
        Call ChangeProgress("Running ZFU_SPL_TO_PDF - BAS " & myFileName(xCount), myPerc)
        
        If BacNo(xCount) <> "" Then
            sessBP.findById("wnd[0]/usr/txtSPOOLNO").Text = BacNo(xCount)
            sessBP.findById("wnd[0]/tbar[1]/btn[8]").press
            sessBP.findById("wnd[1]/usr/ctxtDY_PATH").Text = myPath(xCount)
            sessBP.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "BAS " & myFileName(xCount) & ".pdf"
            sessBP.findById("wnd[1]/tbar[0]/btn[11]").press
            sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
        End If
    Next
    
    
    'sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    'sessBP.findById("wnd[0]/usr/btnSTARTBUTTON").press
    
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    sessBP.findById("wnd[0]").sendVKey 0
End Sub




Sub SAP_BP_5(myEnt As String, myWS As Worksheet)
    
    
    
    With myWS
        .Range("G1").Value = "FS10N"
        .Range("H1").Value = "A3620001"
        
        .Range("G3").Value = "PERIOD"
        .Range("H3").Value = "DEBIT"
        .Range("I3").Value = "CREDIT"
        .Range("J3").Value = "BALANCE"
        .Range("K3").Value = "CUMULATIVE BALANCE"
        
        .Range("M1").Value = "FS10N"
        .Range("N1").Value = "A2680008"
        .Range("M3").Value = "PERIOD"
        .Range("N3").Value = "DEBIT"
        .Range("O3").Value = "CREDIT"
        .Range("P3").Value = "BALANCE"
        .Range("Q3").Value = "CUMULATIVE BALANCE"
        
        .Range("G1:Q3").Font.Bold = True
        .Range("G3:K3").Interior.Color = RGB(255, 255, 0)
        .Range("M3:Q3").Interior.Color = RGB(255, 255, 0)
    End With
    
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    'sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    'sessBP.findById("wnd[0]").sendVKey 0
    
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "FS10N"
    sessBP.findById("wnd[0]").sendVKey 0
    sessBP.findById("wnd[0]/usr/ctxtSO_SAKNR-LOW").Text = "A3620001"
    sessBP.findById("wnd[0]/usr/ctxtSO_BUKRS-LOW").Text = myEnt
    sessBP.findById("wnd[0]/usr/txtGP_GJAHR").Text = Year(DateAdd("m", -1, Date))
    sessBP.findById("wnd[0]/tbar[1]/btn[8]").press
    
    If Not sessBP.findById("wnd[1]", False) Is Nothing Then
        sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
        GoTo nothingFound1
    End If
    Set myTable = sessBP.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell")
    
    For xCount = 0 To 17
        myWS.Range("G" & xCount + 4).Value = myTable.getcellvalue(xCount, "PERIOD")
        myWS.Range("H" & xCount + 4).Value = myTable.getcellvalue(xCount, "DEBIT")
        myWS.Range("I" & xCount + 4).Value = myTable.getcellvalue(xCount, "CREDIT")
        myWS.Range("J" & xCount + 4).Value = myTable.getcellvalue(xCount, "BALANCE")
        myWS.Range("K" & xCount + 4).Value = myTable.getcellvalue(xCount, "BALANCE_CUM")
        
        If Right(myWS.Range("H" & xCount + 4).Value, 1) = "-" Then myWS.Range("H" & xCount + 4).Value = "-" & Left(myWS.Range("H" & xCount + 4).Value, Len(myWS.Range("H" & xCount + 4).Value) - 1)
        If Right(myWS.Range("I" & xCount + 4).Value, 1) = "-" Then myWS.Range("I" & xCount + 4).Value = "-" & Left(myWS.Range("I" & xCount + 4).Value, Len(myWS.Range("I" & xCount + 4).Value) - 1)
        If Right(myWS.Range("J" & xCount + 4).Value, 1) = "-" Then myWS.Range("J" & xCount + 4).Value = "-" & Left(myWS.Range("J" & xCount + 4).Value, Len(myWS.Range("J" & xCount + 4).Value) - 1)
        If Right(myWS.Range("K" & xCount + 4).Value, 1) = "-" Then myWS.Range("K" & xCount + 4).Value = "-" & Left(myWS.Range("K" & xCount + 4).Value, Len(myWS.Range("K" & xCount + 4).Value) - 1)
        
    Next
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    

nothingFound1:
    
    sessBP.findById("wnd[0]/usr/ctxtSO_SAKNR-LOW").Text = "A2680008"
    sessBP.findById("wnd[0]/tbar[1]/btn[8]").press
    
    
    If Not sessBP.findById("wnd[1]", False) Is Nothing Then
        sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
        GoTo nothingFound2
    End If
    
    
    Set myTable = sessBP.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell")
    
    For xCount = 0 To 17
        myWS.Range("M" & xCount + 4).Value = myTable.getcellvalue(xCount, "PERIOD")
        myWS.Range("N" & xCount + 4).Value = myTable.getcellvalue(xCount, "DEBIT")
        myWS.Range("O" & xCount + 4).Value = myTable.getcellvalue(xCount, "CREDIT")
        myWS.Range("P" & xCount + 4).Value = myTable.getcellvalue(xCount, "BALANCE")
        myWS.Range("Q" & xCount + 4).Value = myTable.getcellvalue(xCount, "BALANCE_CUM")
        
        If Right(myWS.Range("N" & xCount + 4).Value, 1) = "-" Then myWS.Range("N" & xCount + 4).Value = "-" & Left(myWS.Range("N" & xCount + 4).Value, Len(myWS.Range("N" & xCount + 4).Value) - 1)
        If Right(myWS.Range("O" & xCount + 4).Value, 1) = "-" Then myWS.Range("O" & xCount + 4).Value = "-" & Left(myWS.Range("O" & xCount + 4).Value, Len(myWS.Range("O" & xCount + 4).Value) - 1)
        If Right(myWS.Range("P" & xCount + 4).Value, 1) = "-" Then myWS.Range("P" & xCount + 4).Value = "-" & Left(myWS.Range("P" & xCount + 4).Value, Len(myWS.Range("P" & xCount + 4).Value) - 1)
        If Right(myWS.Range("Q" & xCount + 4).Value, 1) = "-" Then myWS.Range("Q" & xCount + 4).Value = "-" & Left(myWS.Range("Q" & xCount + 4).Value, Len(myWS.Range("Q" & xCount + 4).Value) - 1)
        
    Next
    
nothingFound2:
    Set myTable = Nothing
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
End Sub


Sub SAP_BP_6()
    
    Call ChangeProgress("Running FBL3N", 0.66)
    
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "FBL3N"
    sessBP.findById("wnd[0]").sendVKey 0
    sessBP.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press
    sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "A2680008"
    sessBP.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "A3620001"
    sessBP.findById("wnd[1]/tbar[0]/btn[8]").press
    sessBP.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").Text = "AU01"
    sessBP.findById("wnd[0]/usr/ctxtPA_STIDA").Text = Format(DateSerial(Year(Date), Month(Date), 1 - 1), "dd.mm.yyyy")
    sessBP.findById("wnd[0]/tbar[1]/btn[8]").press
    
    sessBP.findById("wnd[0]/tbar[1]/btn[32]").press
    sessBP.findById("wnd[1]/usr/btnAPP_FL_ALL").press
    
    sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(2).Selected = True
    sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
    
    sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
    sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document currency"
    sessBP.findById("wnd[2]/tbar[0]/btn[0]").press
    sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
    
    sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
    sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Eff.exchange rate"
    sessBP.findById("wnd[2]/tbar[0]/btn[0]").press
    sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
    
    'sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 60
    'sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(63).Selected = True
    'sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
    'sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 72
    'sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(79).Selected = True
    'sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
    sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
    
    sessBP.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
    'sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
    sessBP.findById("wnd[1]/usr/ctxtDY_PATH").Text = Path_Main
    sessBP.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "FBL3N.xlsx"
    sessBP.findById("wnd[1]/tbar[0]/btn[11]").press
    
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
End Sub


Sub SAP_BP_7()
    Dim xRow As Long, xCount As Long, xRow1 As Long
    Dim xCol As Long
    
    
    Call ChangeProgress("Running ZFR_EXCHANGE_RATES", 0.64)
    
    
    wsExch.Range("A2:A32").EntireRow.ClearContents
    
    wsExch.Range("A1").Value = "Refreshed on " & Now
    
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "ZFR_EXCHANGE_RATES"
    sessBP.findById("wnd[0]").sendVKey 0
    sessBP.findById("wnd[0]/usr/ctxtS_KURST-LOW").Text = "M"
    sessBP.findById("wnd[0]/usr/ctxtS_TCURR-LOW").Text = "AUD"
    
    sessBP.findById("wnd[0]/usr/ctxtS_DATUM-LOW").Text = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "dd.mm.yyyy")
    sessBP.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").Text = Format(DateSerial(Year(Date), Month(Date), 1 - 1), "dd.mm.yyyy")
    
    
    For xCol = 2 To 8
        Select Case xCol
        Case 2
            Call ChangeProgress("Running ZFR_EXCHANGE_RATES - USD", 0.64)
            sessBP.findById("wnd[0]/usr/ctxtS_FCURR-LOW").Text = "USD"
        Case 3
            Call ChangeProgress("Running ZFR_EXCHANGE_RATES - CAD", 0.64)
            sessBP.findById("wnd[0]/usr/ctxtS_FCURR-LOW").Text = "CAD"
        Case 4
            Call ChangeProgress("Running ZFR_EXCHANGE_RATES - EUR", 0.65)
            sessBP.findById("wnd[0]/usr/ctxtS_FCURR-LOW").Text = "EUR"
        Case 5
            Call ChangeProgress("Running ZFR_EXCHANGE_RATES - KRW", 0.65)
            sessBP.findById("wnd[0]/usr/ctxtS_FCURR-LOW").Text = "KRW"
        Case 6
            Call ChangeProgress("Running ZFR_EXCHANGE_RATES - NOK", 0.65)
            sessBP.findById("wnd[0]/usr/ctxtS_FCURR-LOW").Text = "NOK"
        Case 7
            Call ChangeProgress("Running ZFR_EXCHANGE_RATES - GBP", 0.66)
            sessBP.findById("wnd[0]/usr/ctxtS_FCURR-LOW").Text = "GBP"
        Case 8
            Call ChangeProgress("Running ZFR_EXCHANGE_RATES - JPY", 0.66)
            sessBP.findById("wnd[0]/usr/ctxtS_FCURR-LOW").Text = "JPY"
        End Select
        
        sessBP.findById("wnd[0]/tbar[1]/btn[8]").press
        
        If xCol = 2 Then
            For xRow = 2 To Day(DateSerial(Year(Date), Month(Date), 1 - 1)) + 1
                wsExch.Range("A" & xRow).Value = sessBP.findById("wnd[0]/usr/lbl[18," & xRow + 1 & "]").Text
            Next
        End If
        
        For xRow = 2 To Day(DateSerial(Year(Date), Month(Date), 1 - 1)) + 1
            wsExch.Range(NumberToLetter(xCol) & xRow).Value = sessBP.findById("wnd[0]/usr/lbl[41," & xRow + 1 & "]").Text
        Next
        
        sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
        
    Next
        
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
End Sub






