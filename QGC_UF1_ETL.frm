VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf1 
   Caption         =   "Australia QGC Automation Tool Run 1"
   ClientHeight    =   4335
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   16560
   OleObjectBlob   =   "QGC_UF1_ETL.frx":0000
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



'screenshot
'https://stackoverflow.com/questions/43904385/using-excel-vba-macro-to-capture-save-screenshot-of-specific-area-in-same-file


'**************************************************************************************************************
'Requirements from this form
'**************************************************************************************************************
'1. Download GL and GST (Advance tax return and advance tax Report) from P42

'Const TestingRun As Boolean = True
'Const SkipPDF As Boolean = True


'**************************************************************************************************************
'Declaration of module variables
'**************************************************************************************************************
Private StartTime   As Double, SecondsElapsed As Double, MinutesElapsed As String
Private inputError  As String

Private wsSettings  As Worksheet, wsRef As Worksheet, wsData As Worksheet, wsCol As Worksheet
Private SesSion

Private wbMain      As Workbook
Private wsM_Data    As Worksheet, wsM_GST As Worksheet, wsM_Clearing As Worksheet, wsM_Input As Worksheet

Private wbCash      As Workbook
Private wsC_Control As Worksheet, wsC_Forcasting As Worksheet, wsC_Journal As Worksheet

Private failedReport As String
Private myDate As Date

Private myPath(15) As String, myPathMain As String, myFile(10) As String

Private TestPDF As Boolean, TestCashCall As Boolean

Private setFormat As Boolean

Private RefRange As Range

Private percProgress As Double


'**************************************************************************************************************
'Main Routine
'**************************************************************************************************************

Private Sub BtnRun_Click()
    
    StartTime = Timer
    inputError = ""
    failedReport = ""
    
    'TestPDF = Me.cbPDF.Value
    TestCashCall = Me.cbCash.Value
    'TestInput = Me.cbInput.Value
    
    checkFields
    
    If inputError <> "" Then
        MsgBox ("Please ensure the following fields are filled before proceeding:" & inputError)
        Exit Sub
    ElseIf Me.cbCash.Value = False And Me.cbInput_BGIA.Value = False And Me.cbInput_QCLNG.Value = False And Me.cbInput_QGC.Value = False And Me.cbInput_Single.Value = False And Me.cbSAP.Value = False Then
        MsgBox ("Please select one of the functions to run.")
        Exit Sub
    End If
    
    
    'check if SAP BP is open
    If Me.cbSAP.Value = True Then
        checkOpenSAP
        If openP42 = False Then
            MsgBox ("Please ensure you have SAP P42 open.")
            Exit Sub
        Else
            Set SesSion = sessP42
        End If
    End If
    
    If inputError <> "" Then
        MsgBox ("Error in your selected Input Form/ Cash Flow template. Potential worksheet naming issue. Please contact dev team for further assistance.")
        Exit Sub
    End If
    
    
    RunPauseAll
    
    
    Me.Frame1.Visible = True
    Me.btnClose.Visible = False
    Me.Height = 196
    Me.Width = 239.5
    
    
    
    'create save path folder
    createAllFolders
    
    'initiate downloading in queue
    StartDownloadQueue
    
    percProgress = 0.32
    Call ChangeProgress("Closing downloaded files")
    
    Call CloseFiles
    
    If TestCashCall = True Then
        'do cash call
        openFileCash
        doCash
    End If
    
    
    If Me.cbInput_BGIA.Value = True Or Me.cbInput_QCLNG.Value = True Or Me.cbInput_QGC.Value = True Or Me.cbInput_Single.Value = True Then
    'If TestInput = True Then
        'do input form
        doInput
    End If
    
    RunActivateAll
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    
    percProgress = 1
    Call ChangeProgress("Automation process completed for Australia QGC GST." & vbLf & "Time taken: " & MinutesElapsed)
    Me.btnClose.Visible = True
    
    
    'MsgBox ("Completed. Time taken: " & MinutesElapsed)
    'Unload Me
    
End Sub


Private Sub UserForm_Initialize()
    'ShowTitleBar Me
    Frame1.Visible = False
    
    
    Set wsSettings = ThisWorkbook.Worksheets("Settings")
    Set wsRef = ThisWorkbook.Worksheets("Ref")
    Set wsData = ThisWorkbook.Worksheets("Data Entry")
    Set wsCol = ThisWorkbook.Worksheets("ColE")
    
    
    If Not Environ("username") = wsSettings.Range("B1").Value Then
        wsSettings.Range("B:B").EntireColumn.Clear
        wsSettings.Range("B1").Value = Environ("username")
    Else
        Me.tbSaveLocation.Value = wsSettings.Range("B2").Value
    End If
    
    
    
End Sub


'**************************************************************************************************************
'Requirement Runs
'**************************************************************************************************************


Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnSaveLocation_Click()
    Me.tbSaveLocation.Value = SearchFolderLocation
    wsSettings.Range("B2").Value = Me.tbSaveLocation.Value
End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub

'**************************************************************************************************************
'Processing Runs
'**************************************************************************************************************

Sub ChangeProgress(VarTitle As String)
    
    Me.LabelCaption.Caption = VarTitle
    
    If percProgress > 1 Then
        LabelProgress.Width = percProgress / 100 * FrameProgress.Width
    Else
        LabelProgress.Width = percProgress * FrameProgress.Width
    End If
    
    Me.Repaint
End Sub



Sub checkFields()
    
    If Me.tbSaveLocation.Value = "" Then
        inputError = inputError & vbLf & " - Save Location (1)"
    ElseIf Dir(Me.tbSaveLocation.Value, vbDirectory) = "" Then
        inputError = inputError & vbLf & " - Save Location (2)"
    End If
    
    If Not Me.tbDate.Value = "" Then
        If Not IsDate(Me.tbDate.Value) Then
            If MsgBox("Payment due date not detected as date format. This will be left blank. Proceed?", vbYesNo) = vbNo Then
                End
            End If
        Else
            myDate = Me.tbDate.Value
        End If
    End If
    
    
End Sub



Sub createAllFolders()
    Dim tempPath As String, xCount As Long, myTempPath As String
    
    xCount = 1
    
    percProgress = 0.01
    Call ChangeProgress("Creating folders")
    
    If Me.cbSAP.Value = True Then
        'need downloading
        If InStr(1, Me.tbSaveLocation.Value, Year(DateAdd("m", -1, Date)) & "\", vbTextCompare) = 0 Then
            If Dir(Me.tbSaveLocation.Value & Year(DateAdd("m", -1, Date)) & "\", vbDirectory) = "" Then
                MkDir Me.tbSaveLocation.Value & Year(DateAdd("m", -1, Date)) & "\"
            End If
            
            tempPath = Me.tbSaveLocation.Value & Year(DateAdd("m", -1, Date)) & "\"
        Else
            tempPath = Me.tbSaveLocation.Value
        End If
        
        xCount = 1
        myPathMain = tempPath & Format(DateAdd("m", -1, Date), "mm mmm") & "\"
        While Dir(myPathMain, vbDirectory) <> ""
            myPathMain = tempPath & Format(DateAdd("m", -1, Date), "mm mmm") & " (" & xCount & ")\"
            xCount = xCount + 1
        Wend
        
        MkDir myPathMain
        
    Else
        'already downloaded report, just read
        If InStr(1, Me.tbSaveLocation.Value, Year(DateAdd("m", -1, Date)) & "\" & Format(DateAdd("m", -1, Date), "mm mmm") & "\", vbTextCompare) <> 0 Then
            myPathMain = Me.tbSaveLocation.Value
        ElseIf InStr(1, Me.tbSaveLocation.Value, Year(DateAdd("m", -1, Date)) & "\", vbTextCompare) <> 0 Then
            If Dir(Me.tbSaveLocation.Value & Format(DateAdd("m", -1, Date), "mm mmm") & "\", vbDirectory) <> "" Then
                myPathMain = Me.tbSaveLocation.Value & Format(DateAdd("m", -1, Date), "mm mmm") & "\"
            Else
                MsgBox ("Folder not found. Please ensure folder is available if reports has been downloaded (1)")
                End
            End If
        ElseIf Dir(Me.tbSaveLocation.Value & Year(DateAdd("m", -1, Date)) & "\" & Format(DateAdd("m", -1, Date), "mm mmm") & "\", vbDirectory) <> "" Then
            myPathMain = Me.tbSaveLocation.Value & Year(DateAdd("m", -1, Date)) & "\" & Format(DateAdd("m", -1, Date), "mm mmm") & "\"
        Else
            MsgBox ("Folder not found. Please ensure folder is available if reports has been downloaded (1)")
            End
        End If
    End If
    
    
    wsSettings.Range("B6").Value = myPathMain
    
    myPath(0) = myPathMain & "1100 - BGIA (QGC Upstream)\"
    myPath(1) = myPathMain & "5000 - QGC Group\"
    myPath(2) = myPathMain & "5000 - QGC JV\"
    myPath(3) = myPathMain & "5030 - Toll Co 2\"
    myPath(4) = myPathMain & "5031 - Toll Co 2 (2)\"
    myPath(5) = myPathMain & "5033 - Toll Co 1\"
    myPath(6) = myPathMain & "5036 - QCLNG (OpCo)\"
    myPath(7) = myPathMain & "5037 - Train 1\"
    myPath(8) = myPathMain & "5038 - Train 2\"
    myPath(9) = myPathMain & "5045 - T1 UJV\"
    myPath(10) = myPathMain & "5046 - T2 UJV\"
    
    myPath(11) = myPathMain & "BAS Control Checklist\"
    myPath(12) = myPathMain & "Cash Call\"
    myPath(13) = myPathMain & "Monthly IAS\"
    myPath(14) = myPathMain & "NIL BAS\"
    myPath(15) = myPathMain & "Uploaded to SharePoint\"
    
    
    If Me.cbSAP.Value = True Then
        For xCount = 0 To 15
            If Dir(myPath(xCount), vbDirectory) = "" Then MkDir myPath(xCount)
        Next
    Else
        For xCount = 0 To 15
            If Dir(myPath(xCount), vbDirectory) = "" Then
                'Debug.Print myPath(xCount)
                MsgBox ("Folders not found. Please ensure folder is available if reports has been downloaded (3)")
                End
            End If
        Next
    End If
    
    
End Sub

Sub StartDownloadQueue()
    
    If Me.cbSAP.Value = True Then
        SAP_Run1
        SAP_Run2
        SAP_Run3
    End If
    
    'resaveSAP
    
End Sub



Sub openFileCash()
    Dim xRow As Long
    Dim tempWs As Worksheet, tempWs1 As Worksheet
    
    If TestCashCall = True Then
        
        percProgress = 0.33
        Call ChangeProgress("Cash Calls - Open Files")
        'to amend so dont need to select for cash call
        
        Set wbCash = Workbooks.Add
        Set tempWs = wbCash.Worksheets(1)
        
        Set tempWs1 = ThisWorkbook.Worksheets("Control")
        tempWs1.Visible = xlSheetVisible
        tempWs1.Copy after:=tempWs
        tempWs1.Visible = xlSheetVeryHidden
        tempWs.Delete
        Set tempWs = Nothing
        
        Set wsC_Control = wbCash.Worksheets("Control")
        
        Set tempWs1 = ThisWorkbook.Worksheets("Cash forecasting")
        tempWs1.Visible = xlSheetVisible
        tempWs1.Copy after:=wsC_Control
        tempWs1.Visible = xlSheetVeryHidden
        
        Set wsC_Forcasting = wbCash.Worksheets("Cash forecasting")
        
        Set tempWs1 = ThisWorkbook.Worksheets("Journal Entries by BAS Group")
        tempWs1.Visible = xlSheetVisible
        tempWs1.Copy after:=wsC_Forcasting
        tempWs1.Visible = xlSheetVeryHidden
        
        Set wsC_Journal = wbCash.Worksheets("Journal Entries by BAS Group")
        
        wsC_Control.Range("C7").Value = Format(DateAdd("m", -1, Date), "MMM YYYY")
        
        For xRow = 4 To 63
            Select Case xRow
            Case 4, 5, 9, 10, 14, 18, 22, 26, 30, 34, 38, 42 To 50, 54 To 58, 62, 63
                wsC_Journal.Range("C" & xRow).Value = Format(DateAdd("m", -1, Date), "MMM YYYY") & Right(wsC_Journal.Range("C" & xRow).Value, Len(wsC_Journal.Range("C" & xRow).Value) - 8)
            End Select
        Next
        
        With wsC_Forcasting
            .Range("D6").Formula = "='Journal Entries by BAS Group'!M51"
            .Range("D7").Formula = "='Journal Entries by BAS Group'!I6"
            .Range("D8").Formula = "='Journal Entries by BAS Group'!I9"
            .Range("D12").Formula = "='Journal Entries by BAS Group'!L59"
            .Range("D13").Formula = "='Journal Entries by BAS Group'!I64"
            .Range("D14").Formula = "='Journal Entries by BAS Group'!I15"
            .Range("D15").Formula = "='Journal Entries by BAS Group'!I19"
            .Range("D16").Formula = "='Journal Entries by BAS Group'!I31"
            .Range("D17").Formula = "='Journal Entries by BAS Group'!I34"
            .Range("D18").Formula = "='Journal Entries by BAS Group'!I39"
            .Range("D19").Formula = "='Journal Entries by BAS Group'!I23"
            .Range("D20").Formula = "='Journal Entries by BAS Group'!I27"
            
        
            If Not myDate = 0 Then
                .Range("E6:E20").Value = myDate
                .Range("E27:E38").Value = myDate
                .Range("E6:E20").Value = myDate
                
                For xRow = 46 To 66
                    If xRow Mod 2 = 0 Then
                        .Range("R" & xRow).Value = myDate
                    End If
                Next
            Else
                .Range("E6:E20").Value = ""
                .Range("E27:E38").Value = ""
                .Range("E6:E20").Value = ""
                
                For xRow = 46 To 66
                    If xRow Mod 2 = 0 Then
                        .Range("R" & xRow).Value = ""
                    End If
                Next
            End If
            
            .Range("E46:G65").Value = 0
            
        End With
        
        With wsC_Journal
            .Range("C4").Formula = "=Control!$C$7 & "" - QGC Group JV"""
            .Range("C5").Formula = "=Control!$C$7 & "" - QGC Group JV"""
            
            .Range("C9").Formula = "=Control!$C$7 & "" - JV 171"""
            .Range("C10").Formula = "=Control!$C$7 & "" - JV 769"""
            
            .Range("C14").Formula = "=Control!$C$7 & "" - QCLNG T1 UJV BAS"""
            
            .Range("C18").Formula = "=Control!$C$7 & "" - QCLNG T2 UJV BAS"""
            
            .Range("C22").Formula = "=Control!$C$7 & "" - QGC T1 BAS"""
            
            .Range("C26").Formula = "=Control!$C$7 & "" - QGC T2 BAS"""
            
            .Range("C30").Formula = "=Control!$C$7 & "" - QGC T1 Toll Co BAS"""
            
            .Range("C34").Formula = "=Control!$C$7 & "" - QGC T2 Toll Co BAS"""
            
            .Range("C38").Formula = "=Control!$C$7 & "" - QGC T2 (2) Toll Co BAS"""
            
            .Range("C42").Formula = "=Control!$C$7 & "" - Diesel fuel rebate"""
            .Range("C43").Formula = "=Control!$C$7 & "" - Fleet plus a/c"""
            .Range("C44").Formula = "=Control!$C$7 & "" - FBT installment"""
            .Range("C45").Formula = "=Control!$C$7 & "" - QGC GST"""
            .Range("C46").Formula = "=Control!$C$7 & "" - QGC GST"""
            .Range("C47").Formula = "=Control!$C$7 & "" - QGC GST"""
            .Range("C48").Formula = "=Control!$C$7 & "" - QGC GST"""
            .Range("C49").Formula = "=Control!$C$7 & "" - QGC GST"""
            .Range("C50").Formula = "=Control!$C$7 & "" - QGC GST"""
            
            .Range("C54").Formula = "=Control!$C$7 & "" - BGIA Group BAS"""
            .Range("C55").Formula = "=Control!$C$7 & "" - BGIA Group BAS"""
            .Range("C56").Formula = "=Control!$C$7 & "" - BGIA Group BAS"""
            .Range("C57").Formula = "=Control!$C$7 & "" - BGIA Group BAS"""
            .Range("C58").Formula = "=Control!$C$7 & "" - BGIA Group BAS"""
            
            .Range("C62").Formula = "=Control!$C$7 & "" - QCLNG OpCo BAS"""
            .Range("C63").Formula = "=Control!$C$7 & "" - QCLNG OpCo BAS"""
        End With
        
    End If
    
End Sub





Sub doCash()
    Dim xCount As Long, xInOut As Long
    Dim tempWb As Workbook
    Dim tempWs As Worksheet
    Dim tempFile As String
    Dim refRow As Long
    Dim curFolder As String, curFileIn As String, curFileOut As String
    Dim xRow As Long
    
    'Set wsC_Control = wbCash.Worksheets("Control")
    'Set wsC_Forcasting = wbCash.Worksheets("Cash forecasting")
    'Set wsC_Journal = wbCash.Worksheets("Journal Entries by BAS Group")
    
    refRow = 46
    
    percProgress = 0.34
    Call ChangeProgress("Cash Calls - consolidating files")
    
    For xCount = 1 To 10
        
        Select Case xCount
        Case 1 'QGC Pty Ltd    - 2 is JV, not applicable
            curFileIn = "QGC Input " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileOut = "QGC Output " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFolder = myPath(1)
        Case 2 'QGC Upstream Holdings Pty Ltd (BGIA)
            curFileIn = "BGIA Input " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileOut = "BGIA Output " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFolder = myPath(0)
        Case 3 'QCLNG Operating Company Pty Ltd
            curFileIn = "QCLNG Input " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileOut = "QCLNG Output " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFolder = myPath(6)
        Case 4 'QCLNG – QGC / CNOOC T1 Joint Venture
            curFileIn = "5045 Input " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileOut = "5045 Output " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFolder = myPath(9)
        Case 5 'QCLNG – QGC / Tokyo Gas T2 Joint Venture
            curFileIn = "5046 Input " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileOut = "5046 Output " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFolder = myPath(10)
        Case 6 'QGC Train 1 Tolling Pty Ltd
            curFileIn = "5033 Input " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileOut = "5033 Output " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFolder = myPath(5)
        Case 7 'QGC Train 1 Pty Ltd
            curFileIn = "5037 Input " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileOut = "5037 Output " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFolder = myPath(7)
        Case 8 'QCLNG Train 2 Pty Ltd
            curFileIn = "5038 Input " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileOut = "5038 Output " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFolder = myPath(8)
        Case 9 'QGC Train 2 Tolling Pty Ltd
            curFileIn = "5030 Input " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileOut = "5030 Output " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFolder = myPath(3)
        Case 10 'QGC Train 2 Tolling No. 2 Pty Ltd
            curFileIn = "5031 Input " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileOut = "5031 Output " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFolder = myPath(4)
        End Select
        
        
        For xInOut = 1 To 2
            If xInOut = 1 Then
                tempFile = curFileIn
            Else
                tempFile = curFileOut
            End If
        
            If Dir(curFolder & tempFile & ".xlsx") <> "" Then
                Set tempWb = Workbooks.Open(curFolder & tempFile & ".xlsx", ReadOnly:=True)
                Set tempWs = tempWb.Worksheets(1)
                
                xRow = 2
                
                While Not tempWs.Range("A" & xRow).Value = ""
                    xRow = xRow + 1
                Wend
                
                While tempWs.Range("A" & xRow).Interior.Color = RGB(255, 255, 0)
                    tempWs.Range("A" & xRow).EntireRow.Delete
                Wend
                
                tempWs.Range("S2").Value = "Total Amount"
                tempWs.Range("S3").Value = "Total Valuation"
                tempWs.Range("S4").Value = "Total Clearing"
                
                If xCount >= 3 Then
                    tempWs.Range("T2").Formula = "=SUM(I:I)"
                    tempWs.Range("T3").Formula = "=SUMIF(Q:Q,""*valuation*"",I:I)"
                    tempWs.Range("T4").Formula = "=SUMIF(Q:Q,""*clearing*"",I:I)"
                Else
                    tempWs.Range("T2").Formula = "=SUM(E:E)"
                    tempWs.Range("T3").Formula = "=SUMIF(Q:Q,""*valuation*"",E:E)"
                    tempWs.Range("T4").Formula = "=SUMIF(Q:Q,""*clearing*"",E:E)"
                End If
                
                tempWs.Calculate
                
                wsC_Forcasting.Range("E" & refRow).Value = tempWs.Range("T2").Value
                wsC_Forcasting.Range("F" & refRow).Value = tempWs.Range("T3").Value
                wsC_Forcasting.Range("G" & refRow).Value = tempWs.Range("T4").Value
                
                tempWb.SaveAs curFolder & tempFile & " (1).xlsx"
                tempWb.Close False
                Set tempWb = Nothing
                Set tempWs = Nothing
            End If
            
            refRow = refRow + 1
        Next
        
    Next
    
    wsC_Forcasting.Calculate
    
    Call emailCash
    
    wsSettings.Range("B10").Value = myPath(12) & Format(DateAdd("m", -1, Date), "MMM YYYY") & "BAS - Cash Call Estimates.xlsx"
    
    wbCash.SaveAs myPath(12) & Format(DateAdd("m", -1, Date), "MMM YYYY") & " BAS - Cash Call Estimates.xlsx"
    wbCash.Close False
    
End Sub


Sub doInput()
    
    Dim xCount As Long
    Dim xRow As Long, xLastRow As Long
    Dim myWb As Workbook
    Dim myWs1 As Worksheet, myWs2 As Worksheet, myWs3 As Worksheet, ws As Worksheet
    
    Dim myFileUsed As String
    
    Dim refWb As Workbook
    Dim refWs As Worksheet
    
    Dim ImpFileName(1 To 10) As String
    
    Dim curFilePath As String
    Dim curPath As String
    Dim curFileName As String
    Dim tempStr As String
    
    wsData.Range("G1").Value = Month(DateAdd("m", -1, Date))
    wsData.Range("H1").Value = Year(DateAdd("m", -1, Date))
    wsData.Range("H3").Value = myDate
    
    Set myWb = Workbooks.Add
    Set myWs2 = Worksheets(1)
    wsData.Visible = xlSheetVisible
    wsData.Copy before:=myWs2
    wsData.Visible = xlSheetVeryHidden
    Set myWs1 = myWb.Worksheets(1)
    myWs1.Name = "Data Entry"
    myWs1.Tab.Color = RGB(146, 208, 80)
    
    myWs2.Delete
    Set myWs2 = Nothing
    
    wsCol.Visible = xlSheetVisible
    wsCol.Copy after:=myWs1
    Set myWs2 = myWb.Worksheets("ColE")
    myWs2.Name = "Col E Adj G11" 'Format(DateAdd("m", -1, Date), "mmyyyy") & " Col E Adj G11"
    myWs2.Tab.Color = RGB(146, 208, 80)
    wsCol.Visible = xlSheetVeryHidden
    
    percProgress = 0.35
    Call ChangeProgress("Preparing Input form file formats")
    
    Set myWs1 = myWb.Worksheets.Add(after:=myWs2)
    myWs1.Name = "Input GL (20610101)"
    
    Set myWs2 = myWb.Worksheets.Add(after:=myWs1)
    myWs2.Name = "Output GL (20610102)"
    
    Set myWs1 = myWb.Worksheets.Add(after:=myWs2)
    myWs1.Name = "Clearing GL (20610000)"
    
    Set myWs2 = myWb.Worksheets.Add(after:=myWs1)
    myWs2.Name = "GST Report"
    
    Set myWs1 = myWb.Worksheets.Add(after:=myWs2)
    myWs1.Name = "GST Report vs GL Sc"
    
    Set myWs2 = myWb.Worksheets.Add(after:=myWs1)
    myWs2.Name = "Exception List"
    
    Set myWs1 = myWb.Worksheets("Data Entry")
    
    
    Set myWs1 = myWb.Worksheets("GST Report vs GL Sc")
    With myWs1
        .Range("L:P").EntireColumn.Interior.Color = RGB(255, 255, 0)
        .Range("L2").Value = "Effective Rate"
        .Range("M2").Value = "Rightful GST"
        .Range("N2").Value = "Variant"
        .Range("O2").Value = "Remark"
        .Range("P2").Value = "ADJUSTMENT BAS"
    End With
    
    percProgress = 0.35
    Call ChangeProgress("Replicating Input form files")
    
    'added 31 May 2021
    If Me.cbInput_Single.Value = True Then
        '1
        
        xCount = 0
        ImpFileName(1) = "5046 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        While Dir(myPath(10) & ImpFileName(1)) <> ""
            xCount = xCount + 1
            ImpFileName(1) = "5046 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & "(" & xCount & ").xlsx"
        Wend
        myWb.SaveAs myPath(10) & ImpFileName(1)
        'myWb.SaveAs myPath(10) & "5046 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        
        '2
        xCount = 0
        ImpFileName(2) = "5038 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        While Dir(myPath(8) & ImpFileName(2)) <> ""
            xCount = xCount + 1
            ImpFileName(2) = "5038 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & "(" & xCount & ").xlsx"
        Wend
        myWb.SaveAs myPath(8) & ImpFileName(2)
        'myWb.SaveAs myPath(8) & "5038 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        
        '3
        xCount = 0
        ImpFileName(3) = "5045 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        While Dir(myPath(9) & ImpFileName(3)) <> ""
            xCount = xCount + 1
            ImpFileName(3) = "5045 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & "(" & xCount & ").xlsx"
        Wend
        myWb.SaveAs myPath(9) & ImpFileName(3)
        'myWb.SaveAs myPath(9) & "5045 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        
        '4
        xCount = 0
        ImpFileName(4) = "5037 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        While Dir(myPath(7) & ImpFileName(4)) <> ""
            xCount = xCount + 1
            ImpFileName(4) = "5037 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & "(" & xCount & ").xlsx"
        Wend
        myWb.SaveAs myPath(7) & ImpFileName(4)
        'myWb.SaveAs myPath(7) & "5037 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        
        '5
        xCount = 0
        ImpFileName(5) = "5031 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        While Dir(myPath(4) & ImpFileName(5)) <> ""
            xCount = xCount + 1
            ImpFileName(5) = "5031 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & "(" & xCount & ").xlsx"
        Wend
        myWb.SaveAs myPath(4) & ImpFileName(5)
        'myWb.SaveAs myPath(4) & "5031 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        
        '6
        xCount = 0
        ImpFileName(6) = "5030 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        While Dir(myPath(3) & ImpFileName(6)) <> ""
            xCount = xCount + 1
            ImpFileName(6) = "5030 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & "(" & xCount & ").xlsx"
        Wend
        myWb.SaveAs myPath(3) & ImpFileName(6)
        'myWb.SaveAs myPath(3) & "5030 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        
        '7
        xCount = 0
        ImpFileName(7) = "5033 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        While Dir(myPath(5) & ImpFileName(7)) <> ""
            xCount = xCount + 1
            ImpFileName(7) = "5033 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & "(" & xCount & ").xlsx"
        Wend
        myWb.SaveAs myPath(5) & ImpFileName(7)
        'myWb.SaveAs myPath(5) & "5033 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
    
    End If
    
    
    If Me.cbInput_QCLNG.Value = True Then
        '8
        xCount = 0
        ImpFileName(8) = "QCLNG" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        While Dir(myPath(6) & ImpFileName(8)) <> ""
            xCount = xCount + 1
            ImpFileName(8) = "QCLNG" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & "(" & xCount & ").xlsx"
        Wend
        myWb.SaveAs myPath(6) & ImpFileName(8)
        'myWb.SaveAs myPath(6) & "5039 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
    End If
    
    
    If Me.cbInput_BGIA.Value = True Then
        '9
        xCount = 0
        ImpFileName(9) = "BGIA Group" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        While Dir(myPath(0) & ImpFileName(9)) <> ""
            xCount = xCount + 1
            ImpFileName(9) = "BGIA Group" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & "(" & xCount & ").xlsx"
        Wend
        myWb.SaveAs myPath(0) & ImpFileName(9)
        'myWb.SaveAs myPath(0) & "BGIA Group" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
    End If
    
    If Me.cbInput_QGC.Value = True Then
        '10
        xCount = 0
        ImpFileName(10) = "QGC Group" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
        While Dir(myPath(1) & ImpFileName(10)) <> ""
            xCount = xCount + 1
            ImpFileName(10) = "QGC Group" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & "(" & xCount & ").xlsx"
        Wend
        myWb.SaveAs myPath(1) & ImpFileName(10)
        'myWb.SaveAs myPath(1) & "QGC Group" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
    End If
    
    
    myWb.Close False
    Set myWb = Nothing
    Set myWs1 = Nothing
    Set myWs2 = Nothing
    
    
    
    For xCount = 1 To 10
    
        If xCount < 8 Then
            If Me.cbInput_Single.Value = False Then GoTo skippedItem
        ElseIf xCount = 8 Then
            If Me.cbInput_QCLNG.Value = False Then GoTo skippedItem
        ElseIf xCount = 9 Then
            If Me.cbInput_BGIA.Value = False Then GoTo skippedItem
        ElseIf xCount = 10 Then
            If Me.cbInput_QGC.Value = False Then GoTo skippedItem
        End If
        
        Select Case xCount
        Case 1 '5046 GST
            
            percProgress = percProgress + 0.01
            myFileUsed = "5046"
            Call ChangeProgress("Preparing Input Form file: " & myFileUsed)
            
            curPath = myPath(10)
            'curFileName = "5046 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            curFilePath = curPath & ImpFileName(1)
            
            Set myWb = Workbooks.Open(curFilePath)
            Set myWs1 = myWb.Worksheets("Data Entry")
            myWs1.Name = "Data Entry - 5046"
            
            myWs1.Range("B3").Value = "QGC TRAIN 2 UJV MANAGER PTY LTD"
            myWs1.Range("B4").Value = "62 145 383 704"
            myWs1.Range("B5").Value = "5046"
            
            
        Case 2 '5038 GST
            
            percProgress = percProgress + 0.01
            myFileUsed = "5048"
            Call ChangeProgress("Preparing Input Form file: " & myFileUsed)
            
            curPath = myPath(8)
            'curFileName = "5038 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            curFilePath = curPath & ImpFileName(2)
            
            Set myWb = Workbooks.Open(curFilePath)
            Set myWs1 = myWb.Worksheets("Data Entry")
            myWs1.Name = "Data Entry - 5038"
            
            myWs1.Range("B3").Value = "QGC TRAIN 2 PTY LTD"
            myWs1.Range("B4").Value = "27 139 569 458"
            myWs1.Range("B5").Value = "5038"
            
            
        Case 3 '5045 GST
            
            percProgress = percProgress + 0.01
            myFileUsed = "5045"
            Call ChangeProgress("Preparing Input Form file: " & myFileUsed)
            
            curPath = myPath(9)
            'curFileName = "5045 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            curFilePath = curPath & ImpFileName(3)
            
            Set myWb = Workbooks.Open(curFilePath)
            Set myWs1 = myWb.Worksheets("Data Entry")
            myWs1.Name = "Data Entry - 5045"
            
            myWs1.Range("B3").Value = "QGC TRAIN 1 UJV MANAGER PTY LTD"
            myWs1.Range("B4").Value = "13 142 293 776"
            myWs1.Range("B5").Value = "5045"
            
            
        Case 4 '5037 GST
            
            percProgress = percProgress + 0.01
            myFileUsed = "5037"
            Call ChangeProgress("Preparing Input Form file: " & myFileUsed)
            
            curPath = myPath(7)
            'curFileName = "5037 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            curFilePath = curPath & ImpFileName(4)
            
            Set myWb = Workbooks.Open(curFilePath)
            Set myWs1 = myWb.Worksheets("Data Entry")
            myWs1.Name = "Data Entry - 5037"
            
            myWs1.Range("B3").Value = "QGC TRAIN 1 PTY LTD"
            myWs1.Range("B4").Value = "31 139 569 412"
            myWs1.Range("B5").Value = "5037"
            
            
            
            
        Case 5 '5031 GST
            
            percProgress = percProgress + 0.01
            myFileUsed = "5031"
            Call ChangeProgress("Preparing Input Form file: " & myFileUsed)
            
            curPath = myPath(4)
            'curFileName = "5031 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            curFilePath = curPath & ImpFileName(5)
            
            Set myWb = Workbooks.Open(curFilePath)
            Set myWs1 = myWb.Worksheets("Data Entry")
            myWs1.Name = "Data Entry - 5031"
            
            myWs1.Range("B3").Value = "QGC TRAIN 2 TOLLING NO.2 PTY LTD"
            myWs1.Range("B4").Value = "91 147 896 535"
            myWs1.Range("B5").Value = "5031"
            
            Set myWs1 = myWb.Worksheets("Clearing GL (20610000)")
            Set myWs2 = myWb.Worksheets.Add(after:=myWs1)
            myWs2.Name = "Sales GL (Toll Cos)"
            
        Case 6 '5030 GST
            
            percProgress = percProgress + 0.01
            myFileUsed = "5030"
            Call ChangeProgress("Preparing Input Form file: " & myFileUsed)
            
            curPath = myPath(3)
            'curFileName = "5030 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            curFilePath = curPath & ImpFileName(6)
            
            Set myWb = Workbooks.Open(curFilePath)
            Set myWs1 = myWb.Worksheets("Data Entry")
            myWs1.Name = "Data Entry - 5030"
            
            myWs1.Range("B3").Value = "QGC TRAIN 2 TOLLING PTY LTD"
            myWs1.Range("B4").Value = "81 142 293 687"
            myWs1.Range("B5").Value = "5030"
            
            Set myWs1 = myWb.Worksheets("Clearing GL (20610000)")
            Set myWs2 = myWb.Worksheets.Add(after:=myWs1)
            myWs2.Name = "Sales GL (Toll Cos)"
            
        Case 7 '5033 GST
            
            percProgress = percProgress + 0.01
            myFileUsed = "5033"
            Call ChangeProgress("Preparing Input Form file: " & myFileUsed)
            
            curPath = myPath(5)
            'curFileName = "5033 GST" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            curFilePath = curPath & ImpFileName(7)
            
            Set myWb = Workbooks.Open(curFilePath)
            Set myWs1 = myWb.Worksheets("Data Entry")
            myWs1.Name = "Data Entry - 5033"
            
            myWs1.Range("B3").Value = "QGC TRAIN 1 TOLLING PTY LTD"
            myWs1.Range("B4").Value = "87 142 293 650"
            myWs1.Range("B5").Value = "5033"
            
            Set myWs1 = myWb.Worksheets("Clearing GL (20610000)")
            Set myWs2 = myWb.Worksheets.Add(after:=myWs1)
            myWs2.Name = "Sales GL (Toll Cos)"
            
            
        Case 8 'QCLNG
            
            percProgress = percProgress + 0.01
            myFileUsed = "QCLNG"
            Call ChangeProgress("Preparing Input Form file: " & myFileUsed)
            
            curPath = myPath(6)
            'curFileName = "QCLNG" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            curFilePath = curPath & ImpFileName(8)
            
            Set myWb = Workbooks.Open(curFilePath)
            Set myWs1 = myWb.Worksheets("Data Entry")
            With myWs1
                .Name = "Data Entry - QCLNG Group"
                .Range("B3").Value = "QCLNG OPERATING COMPANY PTY LTD"
                .Range("B4").Value = "19 138 872 385"
                .Range("B5").Value = "5036"
                
                .Range("L19").Value = ""
                
                .Range("J31").Value = ""
                .Range("J35").Value = ""
                .Range("J36").Value = ""
                .Range("J37").Value = ""
                
                .Range("G65").Value = ""
                .Range("G72").Value = ""
                .Range("G79").Value = ""
                
            End With
            
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 1127"
            myWs2.Range("B3").Value = "QGC Midstream Investments"
            myWs2.Range("B4").Value = "77 130 857 215"
            myWs2.Range("B5").Value = "1127"
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 1113"
            myWs2.Range("B3").Value = "QGC Midstream Land Pty Ltd"
            myWs2.Range("B4").Value = "56 135 148 506"
            myWs2.Range("B5").Value = "1113"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5039"
            myWs2.Range("B3").Value = "QCLNG Common Facilities Company Pty Ltd"
            myWs2.Range("B4").Value = "33 139 569 485"
            myWs2.Range("B5").Value = "5039"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5036"
            myWs2.Range("B3").Value = "QCLNG Operating Company Pty Ltd"
            myWs2.Range("B4").Value = "19 138 872 385"
            myWs2.Range("B5").Value = "5036"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            'to be amended
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - MOT JV"
            myWs2.Range("B3").Value = "QCLNG xxxxx"
            myWs2.Range("B4").Value = "xx xxx xxx xxx"
            myWs2.Range("B5").Value = "xxxx"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            
            Set myWs1 = myWb.Worksheets("GST Report")
            myWs1.Name = "GST Report - QCLNG"
            
            
            
        Case 9 'BGIA  - 4 sub tabs
            
            percProgress = percProgress + 0.01
            myFileUsed = "BGIA"
            Call ChangeProgress("Preparing Input Form file: " & myFileUsed)
            
            curPath = myPath(0)
            'curFileName = "BGIA Group" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            curFilePath = curPath & ImpFileName(9)
            
            Set myWb = Workbooks.Open(curFilePath)
            Set myWs1 = myWb.Worksheets("Data Entry")
            myWs1.Name = "Data Entry - BGIA Group"
            myWs1.Range("B3").Value = "QGC Upstream Hold. Pty Ltd"
            myWs1.Range("B4").Value = "76 130 856 843"
            myWs1.Range("B5").Value = "1101"
            
            myWs1.Range("J31").Value = ""
            myWs1.Range("J35").Value = ""
            myWs1.Range("J36").Value = ""
            myWs1.Range("J37").Value = ""
            
            myWs1.Range("G65").Value = ""
            myWs1.Range("G72").Value = ""
            myWs1.Range("G79").Value = ""

            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 1122"
            myWs2.Range("B3").Value = "OME Resources Australi PL"
            myWs2.Range("B4").Value = "27 100 280 662"
            myWs2.Range("B5").Value = "1122"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 1116"
            myWs2.Range("B3").Value = "QGC Upstream Finance Pty"
            myWs2.Range("B4").Value = "29 131 154 642"
            myWs2.Range("B5").Value = "1116"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 1112"
            myWs2.Range("B3").Value = "Pure Energy Res PL"
            myWs2.Range("B4").Value = "48 115 514 880"
            myWs2.Range("B5").Value = "1112"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 1106"
            myWs2.Range("B3").Value = "QGC Upstream Limited Part"
            myWs2.Range("B4").Value = "83 715 246 894"
            myWs2.Range("B5").Value = "1106"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 1101"
            myWs2.Range("B3").Value = "QGC Upstream Hold. Pty Lt"
            myWs2.Range("B4").Value = "76 130 856 843"
            myWs2.Range("B5").Value = "1101"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 1100"
            myWs2.Range("B3").Value = "BG Int'l-Australia"
            myWs2.Range("B4").Value = "72 114 818 825"
            myWs2.Range("B5").Value = "1100"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            
            Set myWs1 = myWb.Worksheets("GST Report")
            myWs1.Name = "GST Report - BGIA"
            
            
            
        Case 10 'QGC
            
            percProgress = percProgress + 0.01
            myFileUsed = "QGC"
            Call ChangeProgress("Preparing Input Form file: " & myFileUsed)
            
            curPath = myPath(1)
            'curFileName = "QGC Group" & " Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx"
            curFilePath = curPath & ImpFileName(10)
            
            Set myWb = Workbooks.Open(curFilePath)
            Set myWs1 = myWb.Worksheets("Data Entry")
            myWs1.Name = "Data Entry - QGC Group"
            myWs1.Range("B3").Value = "QGC PTY LIMITED"
            myWs1.Range("B4").Value = "11 089 642 553"
            myWs1.Range("B5").Value = "5000"
            
            
            myWs1.Range("J31").Value = ""
            myWs1.Range("J35").Value = ""
            myWs1.Range("J36").Value = ""
            myWs1.Range("J37").Value = ""
            
            myWs1.Range("G65").Value = ""
            myWs1.Range("G72").Value = ""
            myWs1.Range("G79").Value = ""
            
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - JV 171"
            myWs2.Range("B3").Value = "Roma Petroleum Pty Ltd"
            myWs2.Range("B4").Value = "21 066 018 979"
            myWs2.Range("B5").Value = "5009"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - QGC JV"
            myWs2.Range("B3").Value = "QGC PTY LIMITED - JOINT VENTURE CONSOLIDATION"
            myWs2.Range("B4").Value = "11 089 642 553"
            myWs2.Range("B5").Value = "5000"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5044"
            myWs2.Range("B3").Value = "QGC Sales Queensland Pty"
            myWs2.Range("B4").Value = "76 121 868 273"
            myWs2.Range("B5").Value = "5044"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5043"
            myWs2.Range("B3").Value = "Condamine 4 Pty Ltd"
            myWs2.Range("B4").Value = "50 139 748 566"
            myWs2.Range("B5").Value = "5043"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5042"
            myWs2.Range("B3").Value = "Condamine 3 Pty Ltd"
            myWs2.Range("B4").Value = "46 139 748 548"
            myWs2.Range("B5").Value = "5042"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5041"
            myWs2.Range("B3").Value = "Condamine 2 Pty Ltd"
            myWs2.Range("B4").Value = "52 139 748 511"
            myWs2.Range("B5").Value = "5041"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5040"
            myWs2.Range("B3").Value = "Condamine 1 Pty Ltd"
            myWs2.Range("B4").Value = "31 139 748 486"
            myWs2.Range("B5").Value = "5040"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5032"
            myWs2.Range("B3").Value = "QGC (B7) Pty Ltd"
            myWs2.Range("B4").Value = "21 152 188 335"
            myWs2.Range("B5").Value = "5032"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5029"
            myWs2.Range("B3").Value = "QCLNG Pty Ltd"
            myWs2.Range("B4").Value = "46 150 538 024"
            myWs2.Range("B5").Value = "5029"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5028"
            myWs2.Range("B3").Value = "QGC Northern Forestry Pty"
            myWs2.Range("B4").Value = "45 145 383 697"
            myWs2.Range("B5").Value = "5028"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5027"
            myWs2.Range("B3").Value = "ACN 002 820 555 Pty Ltd"
            myWs2.Range("B4").Value = "92 002 820 555"
            myWs2.Range("B5").Value = "5027"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5026"
            myWs2.Range("B3").Value = "Sunshine Gas Operations P"
            myWs2.Range("B4").Value = "47 099 577 429"
            myWs2.Range("B5").Value = "5026"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5025"
            myWs2.Range("B3").Value = "Australian Oil & Gas Corp"
            myWs2.Range("B4").Value = "86 055 977 270"
            myWs2.Range("B5").Value = "5025"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5024"
            myWs2.Range("B3").Value = "QGC Midstream Srvcs Pty L"
            myWs2.Range("B4").Value = "98 123 756 034"
            myWs2.Range("B5").Value = "5024"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5023"
            myWs2.Range("B3").Value = "QGC (Exploration) Pty Ltd"
            myWs2.Range("B4").Value = "82 133 878 618"
            myWs2.Range("B5").Value = "5023"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5022"
            myWs2.Range("B3").Value = "SGAI Pty Ltd"
            myWs2.Range("B4").Value = "41 116 132 873"
            myWs2.Range("B5").Value = "5022"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5021"
            myWs2.Range("B3").Value = "Queensland Petroleum Comp"
            myWs2.Range("B4").Value = "47 114 654 661"
            myWs2.Range("B5").Value = "5021"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5020"
            myWs2.Range("B3").Value = "QGC IPT Pty Ltd"
            myWs2.Range("B4").Value = ""
            myWs2.Range("B5").Value = "5020"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5019"
            myWs2.Range("B3").Value = "Interstate Energy Pty Ltd"
            myWs2.Range("B4").Value = "94 002 820 564"
            myWs2.Range("B5").Value = "5019"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5018"
            myWs2.Range("B3").Value = "Interstate Pipelines P L"
            myWs2.Range("B4").Value = "50 004 335 013"
            myWs2.Range("B5").Value = "5018"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5017"
            myWs2.Range("B3").Value = "ACN 081 118 292 Pty Ltd"
            myWs2.Range("B4").Value = "63 081 118 292"
            myWs2.Range("B5").Value = "5017"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5016"
            myWs2.Range("B3").Value = "Sunshine Cooper Pty Ltd"
            myWs2.Range("B4").Value = "25 108 530 589"
            myWs2.Range("B5").Value = "5016"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5015"
            myWs2.Range("B3").Value = "Sunshine 685 Pty Limited"
            myWs2.Range("B4").Value = "54 103 512 241"
            myWs2.Range("B5").Value = "5015"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5014"
            myWs2.Range("B3").Value = "New South Oil Pty Ltd"
            myWs2.Range("B4").Value = "41 098 134 706"
            myWs2.Range("B5").Value = "5014"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5013"
            myWs2.Range("B3").Value = "Hamilbent Pty Ltd"
            myWs2.Range("B4").Value = "30 092 052 787"
            myWs2.Range("B5").Value = "5013"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5012"
            myWs2.Range("B3").Value = "BNG (Surat) Pty Ltd"
            myWs2.Range("B4").Value = "97 090 629 913"
            myWs2.Range("B5").Value = "5012"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5011"
            myWs2.Range("B3").Value = "Sunshine Gas Pty Limited"
            myWs2.Range("B4").Value = "44 098 563 663"
            myWs2.Range("B5").Value = "5011"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5009"
            myWs2.Range("B3").Value = "Roma Petroleum Pty Ltd"
            myWs2.Range("B4").Value = "21 066 018 979"
            myWs2.Range("B5").Value = "5009"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5008"
            myWs2.Range("B3").Value = "Petroleum Exploration Aus"
            myWs2.Range("B4").Value = "56 125 525 706"
            myWs2.Range("B5").Value = "5008"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5007"
            myWs2.Range("B3").Value = "Walloons CSG Co Pty Ltd"
            myWs2.Range("B4").Value = "53 130 344 366"
            myWs2.Range("B5").Value = "5007"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5006"
            myWs2.Range("B3").Value = "Gas Resources Limited"
            myWs2.Range("B4").Value = "41 247 728 983"
            myWs2.Range("B5").Value = "5006"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5005"
            myWs2.Range("B3").Value = "SGA (Queensland) Pty Ltd"
            myWs2.Range("B4").Value = "67 114 116 068"
            myWs2.Range("B5").Value = "5005"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5004"
            myWs2.Range("B3").Value = "Starzap Pty Ltd"
            myWs2.Range("B4").Value = "94 079 932 246"
            myWs2.Range("B5").Value = "5004"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5003"
            myWs2.Range("B3").Value = "Queensland Gas Company"
            myWs2.Range("B4").Value = "77 116 145 110"
            myWs2.Range("B5").Value = "5003"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5002"
            myWs2.Range("B3").Value = "QGC (Infrastruct) Pty Ltd"
            myWs2.Range("B4").Value = "77 116 145 174"
            myWs2.Range("B5").Value = "5002"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5001"
            myWs2.Range("B3").Value = "Condamine Power Station Pty Ltd"
            myWs2.Range("B4").Value = "80 120 323 588"
            myWs2.Range("B5").Value = "5001"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(2)
            myWs2.Name = "Data Entry - 5000"
            myWs2.Range("B3").Value = "QGC PTY LIMITED"
            myWs2.Range("B4").Value = "11 089 642 553"
            myWs2.Range("B5").Value = "5000"
            myWs2.Tab.ColorIndex = xlColorIndexNone
            
            
            Set myWs1 = myWb.Worksheets("Col E Adj G11")
            Set myWs2 = myWb.Worksheets.Add(after:=myWs1)
            myWs2.Name = "GST Report TW Recharge Adj_S1"
            
            
            Set myWs1 = myWb.Worksheets("GST Report")
            myWs1.Name = "GST Report - QGC"
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(myWs1.Index + 1)
            myWs2.Name = "GST Report - JV 171"
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(myWs1.Index + 1)
            myWs2.Name = "GST Report - QGC JV"
            
            
            Set myWs1 = myWb.Worksheets("GST Report vs GL Sc")
            myWs1.Name = "GST vs GL Sc - QGC"
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(myWs1.Index + 1)
            myWs2.Name = "GST vs GL Sc - JV 171"
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(myWs1.Index + 1)
            myWs2.Name = "GST vs GL Sc - QGC JV"
            
            
            Set myWs1 = myWb.Worksheets("Exception List")
            myWs1.Name = "Exception List - QGC"
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(myWs1.Index + 1)
            myWs2.Name = "Exception List - JV 171"
            
            myWs1.Copy after:=myWs1
            Set myWs2 = myWb.Worksheets(myWs1.Index + 1)
            myWs2.Name = "Exception List - QGC JV"
            
            
        End Select
        
        myWb.Save
        
        tempStr = Dir(curPath)
        
        While Not tempStr = ""
            If Right(tempStr, 20) = " Input " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx" Then 'InStr(1, tempStr, " Input " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx", vbTextCompare) <> 0 Then
                
                percProgress = percProgress + 0.01
                Call ChangeProgress("Input Form file: " & myFileUsed & " - Input")
                
                
                Set myWs1 = myWb.Worksheets("Input GL (20610101)")
                myWs1.Cells.Delete
                
                '##################################################
                Set refWb = Workbooks.Open(curPath & tempStr)
                Set refWs = refWb.Worksheets(1)
                
                xLastRow = refWs.Cells(refWs.Rows.Count, "B").End(xlUp).Row
                
                While refWs.Range("A" & xLastRow).Interior.Color = RGB(255, 255, 0)
                    refWs.Range("A" & xLastRow).EntireRow.Delete
                    xLastRow = xLastRow - 1
                Wend
                
                refWs.Cells.Copy myWs1.Range("A1")
                
                refWb.Close False
                Set refWb = Nothing
                Set refWs = Nothing
                
                'add summary items on the right
                
                With myWs1
                    .Range("U2:U20").Interior.Color = RGB(218, 238, 243)
                    .Range("R1").Value = "Remarks"
                    .Range("T2").Value = "Total Amt in Loc Curr"
                    .Range("T3").Value = "Valuation"
                    .Range("T4").Value = "Clearing"
                    .Range("T5").Value = "B1"
                    .Range("T6").Value = "B2"
                    .Range("T7").Value = "B6"
                    .Range("T8").Value = "B9"
                    .Range("T9").Value = "IQ"
                    .Range("T10").Value = "P0"
                    .Range("T11").Value = "Without Tax Code"
                    .Range("T12").Value = "ZI"
                    .Range("T13").Value = "GST on asset disposal proceeds"
                    .Range("T14").Value = "B9 (Energy Sales - Doc No: 1600000056)"
                    .Range("T15").Value = "VCA May-July 2020"
                    .Range("T16").Value = "Chinchilla Accm Tax"
                    .Range("T17").Value = "Contracted Gas Sales (CNOOC Receipt)"
                    .Range("T18").Value = "Contracted Gas Sales (TG Receipt)"
                    .Range("T19").Value = "Visa Control Account (Doc No: 901203879)"
                    .Range("T20").Value = "Recode transaction for Tax Code B9 (Doc No: 1600000253)"
                    
                    If xCount = 1 Or xCount = 3 Then
                        .Range("T21").Value = "Cash Call - AUD"
                        .Range("T22").Value = "Cash Call - USD"
                    End If
                    
                    
                    If xCount = 9 Or xCount = 10 Then 'BGIA or QGC
                        'take col E
                        .Range("U2").Formula = "=SUM(E:E)"
                        .Range("U3:U20").Formula = "=SUMIF(R:R,T3,E:E)"
                        .Range("U11").Formula = "=SUMIF(O:O,"""",E:E)"
                        .Range("U2:U20").NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                        
                    ElseIf xCount = 1 Or xCount = 3 Then
                        .Range("U2").Formula = "=SUM(I:I)"
                        .Range("U3:U22").Formula = "=SUMIF(R:R,T3,I:I)"
                        .Range("U11").Formula = "=SUMIF(O:O,"""",I:I)"
                        .Range("U2:U22").NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                        .Range("U2:U22").Interior.Color = RGB(218, 238, 243)
                        
                    Else
                        .Range("U2").Formula = "=SUM(I:I)"
                        .Range("U3:U20").Formula = "=SUMIF(R:R,T3,I:I)"
                        .Range("U11").Formula = "=SUMIF(O:O,"""",I:I)"
                        .Range("U2:U20").NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                    End If
                    
                End With
                
                Set myWs1 = Nothing
                
            ElseIf Right(tempStr, 21) = " Output " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx" Then 'InStr(1, tempStr, " Output " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx", vbTextCompare) <> 0 Then
                
                percProgress = percProgress + 0.01
                Call ChangeProgress("Input Form file: " & myFileUsed & " - Output")
                
                Set myWs1 = myWb.Worksheets("Output GL (20610102)")
                myWs1.Cells.Delete
                
                Set refWb = Workbooks.Open(curPath & tempStr)
                Set refWs = refWb.Worksheets(1)
                
                xLastRow = refWs.Cells(refWs.Rows.Count, "B").End(xlUp).Row
                
                While refWs.Range("A" & xLastRow).Interior.Color = RGB(255, 255, 0)
                    refWs.Range("A" & xLastRow).EntireRow.Delete
                    xLastRow = xLastRow - 1
                Wend
                
                refWs.Cells.Copy myWs1.Range("A1")
                
                refWb.Close False
                Set refWb = Nothing
                Set refWs = Nothing
                
                'add summary items on the right
                
                
                With myWs1
                    .Range("U2:U10").Interior.Color = RGB(218, 238, 243)
                    .Range("R1").Value = "Remarks"
                    .Range("T2").Value = "Total Amt in Loc Curr"
                    .Range("T3").Value = "Valuation"
                    .Range("T4").Value = "Clearing"
                    .Range("T5").Value = "Cost Recovery"
                    .Range("T6").Value = "APA"
                    .Range("T7").Value = "S1"
                    .Range("T8").Value = "GST on asset disposal proceeds"
                    .Range("T9").Value = "T1 UJV"
                    .Range("T10").Value = "T2 UJV"
                    
                    If xCount >= 9 Then
                        .Range("U2").Formula = "=SUM(E:E)"
                        .Range("U3:U10").Formula = "=SUMIF(R:R,T3,E:E)"
                        .Range("U2:U10").NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                    Else
                        .Range("U2").Formula = "=SUM(I:I)"
                        .Range("U3:U10").Formula = "=SUMIF(R:R,T3,I:I)"
                        .Range("U2:U10").NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                    End If
                    
                End With
                
                
                Set myWs1 = Nothing
                
            ElseIf Right(tempStr, 23) = " Clearing " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx" Then 'InStr(1, tempStr, " Clearing " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx", vbTextCompare) <> 0 Then
                
                percProgress = percProgress + 0.01
                Call ChangeProgress("Input Form file: " & myFileUsed & " - Clearing")
                
                
                Set myWs1 = myWb.Worksheets("Clearing GL (20610000)")
                myWs1.Cells.Delete
                
                Set refWb = Workbooks.Open(curPath & tempStr)
                Set refWs = refWb.Worksheets(1)
                
                refWs.Cells.Copy myWs1.Range("A1")
                
                refWb.Close False
                Set refWb = Nothing
                Set refWs = Nothing
                
                Set myWs1 = Nothing
            
            
            ElseIf Right(tempStr, 20) = " Sales " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx" Then
                
                percProgress = percProgress + 0.01
                Call ChangeProgress("Input Form file: " & myFileUsed & " - Sales")
                
                Set myWs1 = myWb.Worksheets("Sales GL (Toll Cos)")
                myWs1.Cells.Delete
                
                Set refWb = Workbooks.Open(curPath & tempStr)
                Set refWs = refWb.Worksheets(1)
                
                refWs.Cells.Copy myWs1.Range("A1")
                refWb.Close False
                Set refWb = Nothing
                Set refWs = Nothing
                
                myWs1.Range("T2").Value = "Total amount (DR Doc type only)"
                myWs1.Range("U2").Formula = "=SUMIF(K:K,""DR"",I:I)"
                
                Set myWs1 = Nothing
                
            
            ElseIf InStr(1, tempStr, " GST Report SAP Download " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xls", vbTextCompare) <> 0 And Right(tempStr, 4) = ".xls" Then
                
                percProgress = percProgress + 0.02
                Call ChangeProgress("Input Form file: " & myFileUsed & " - GST Report")
                
                If xCount = 10 Then
                    If tempStr = "JV171 GST Report SAP Download " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xls" Then
                        Set myWs1 = myWb.Worksheets("GST Report - JV 171")
                        Set myWs2 = myWb.Worksheets("GST vs GL Sc - JV 171")
                        Set myWs3 = myWb.Worksheets("Exception List - JV 171")
                        myWs1.Cells.Delete
                        myWs2.Cells.Delete
                        myWs3.Cells.Delete
                        
                        Set refWb = Workbooks.Open(curPath & tempStr)
                        Set refWs = refWb.Worksheets(1)
                        
                        
                        GSTReport refWb, myWs1, myWs2, myWs3
                        
                        'refWs.Cells.Copy myWs1.Range("A1")
                        
                        refWb.SaveAs curPath & tempStr & "x", FileFormat:=51
                        refWb.Close True
                        Set refWb = Nothing
                        Set refWs = Nothing
                        
                        Set myWs1 = Nothing
                        Set myWs2 = Nothing
                        
                    ElseIf tempStr = "QGC GST Report SAP Download " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xls" Then
                        Set myWs1 = myWb.Worksheets("GST Report - QGC")
                        Set myWs2 = myWb.Worksheets("GST vs GL Sc - QGC")
                        Set myWs3 = myWb.Worksheets("Exception List - QGC")
                        myWs1.Cells.Delete
                        myWs2.Cells.Delete
                        myWs3.Cells.Delete
                        
                        Set refWb = Workbooks.Open(curPath & tempStr)
                        Set refWs = refWb.Worksheets(1)
                        
                        GSTReport refWb, myWs1, myWs2, myWs3
                        
                        'refWs.Cells.Copy myWs1.Range("A1")
                
                        refWb.SaveAs curPath & tempStr & "x", FileFormat:=51
                        refWb.Close True
                        Set refWb = Nothing
                        Set refWs = Nothing
                        
                        Set myWs1 = Nothing
                        Set myWs2 = Nothing
                        
                    ElseIf tempStr = "QGC JV GST Report SAP Download " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xls" Then
                        Set myWs1 = myWb.Worksheets("GST Report - QGC JV")
                        Set myWs2 = myWb.Worksheets("GST vs GL Sc - QGC JV")
                        Set myWs3 = myWb.Worksheets("Exception List - QGC JV")
                        myWs1.Cells.Delete
                        myWs2.Cells.Delete
                        myWs3.Cells.Delete
                        
                        Set refWb = Workbooks.Open(curPath & tempStr)
                        Set refWs = refWb.Worksheets(1)
                        
                        GSTReport refWb, myWs1, myWs2, myWs3
                        
                        'refWs.Cells.Copy myWs1.Range("A1")
                
                        refWb.SaveAs curPath & tempStr & "x", FileFormat:=51
                        refWb.Close True
                        Set refWb = Nothing
                        Set refWs = Nothing
                        
                        Set myWs1 = Nothing
                        Set myWs2 = Nothing
                        
                    End If
                    
                ElseIf xCount = 9 Then '
                    Set myWs1 = myWb.Worksheets("GST Report - BGIA")
                    Set myWs2 = myWb.Worksheets("GST Report vs GL Sc")
                    Set myWs3 = myWb.Worksheets("Exception List")
                    myWs1.Cells.Delete
                    myWs2.Cells.Delete
                    myWs3.Cells.Delete
                    
                    Set refWb = Workbooks.Open(curPath & tempStr)
                    Set refWs = refWb.Worksheets(1)
                    
                    GSTReport refWb, myWs1, myWs2, myWs3
                        
                    'refWs.Cells.Copy myWs1.Range("A1")
                    
                    refWb.SaveAs curPath & tempStr & "x", FileFormat:=51
                    refWb.Close True
                    Set refWb = Nothing
                    Set refWs = Nothing
                    
                    
                    Set myWs1 = Nothing
                    Set myWs2 = Nothing
                    
                    
                    
                ElseIf xCount = 8 Then 'QCLNG
                    Set myWs1 = myWb.Worksheets("GST Report - QCLNG")
                    Set myWs2 = myWb.Worksheets("GST Report vs GL Sc")
                    Set myWs3 = myWb.Worksheets("Exception List")
                    myWs1.Cells.Delete
                    myWs2.Cells.Delete
                    myWs3.Cells.Delete
                    
                    Set refWb = Workbooks.Open(curPath & tempStr)
                    Set refWs = refWb.Worksheets(1)
                    
                     
                     refWb , myWs1, myWs2, myWs3
                        
                    'refWs.Cells.Copy myWs1.Range("A1")
                    
                    refWb.SaveAs curPath & tempStr & "x", FileFormat:=51
                    refWb.Close True
                    Set refWb = Nothing
                    Set refWs = Nothing
                    
                    
                    With myWs1
                        .Range("K1:N2").Font.Bold = True
                        .Range("K10:N11").Font.Bold = True
                        .Range("N1:N17").Font.Bold = True
                        .Range("K1").Value = "5036"
                        .Range("K10").Value = "5039"
                        
                        .Range("L:N").NumberFormat = "#,##0.00"
                        
                        .Range("K2").Value = "Tax Code"
                        .Range("K11").Value = "Tax Code"
                        .Range("L2").Value = "GST Report"
                        .Range("L11").Value = "GST Report"
                        .Range("M2").Value = "GL"
                        .Range("M11").Value = "GL"
                        .Range("N2").Value = "Difference"
                        .Range("N11").Value = "Difference"
                        
                        
                        .Range("K1:N1").Merge
                        .Range("K1").HorizontalAlignment = xlCenter
                        .Range("K10:N10").Merge
                        .Range("K10").HorizontalAlignment = xlCenter
                        
                        .Range("K3:N7").Interior.Color = RGB(255, 255, 0)
                        .Range("K12:N16").Interior.Color = RGB(255, 255, 0)
                        
                        .Range("K3").Value = "B1"
                        .Range("K4").Value = "B2"
                        .Range("K5").Value = "B9"
                        .Range("K6").Value = "IQ"
                        .Range("K7").Value = "S1"
                        
                        .Range("K12").Value = "B1"
                        .Range("K13").Value = "B2"
                        .Range("K14").Value = "B9"
                        .Range("K15").Value = "IQ"
                        .Range("K16").Value = "S1"
                        
                        .Range("L3:L7").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,K3,'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,""5036""))"
                        .Range("M3:M6").Formula = "=SUMIFS('Input GL (20610101)'!$I:$I,'Input GL (20610101)'!$A:$A,""5036"",'Input GL (20610101)'!$O:$O,K3)"
                        .Range("M7").Formula = "=SUMIFS('Output GL (20610102)'!$I:$I,'Output GL (20610102)'!$A:$A,""5036"",'Output GL (20610102)'!$O:$O,K7)"
                        .Range("N3:N7").Formula = "=M3-L3"
                        
                        .Range("L12:L16").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,K12,'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,""5039""))"
                        .Range("M12:M15").Formula = "=SUMIFS('Input GL (20610101)'!$I:$I,'Input GL (20610101)'!$A:$A,""5039"",'Input GL (20610101)'!$O:$O,K12)"
                        .Range("M16").Formula = "=SUMIFS('Output GL (20610102)'!$I:$I,'Output GL (20610102)'!$A:$A,""5036"",'Output GL (20610102)'!$O:$O,K7)"
                        .Range("N12:N16").Formula = "=M12-L12"
                        
                        
                    End With
                    
                    
                    Set myWs1 = Nothing
                    Set myWs2 = Nothing
                    
                Else
                    Set myWs1 = myWb.Worksheets("GST Report")
                    Set myWs2 = myWb.Worksheets("GST Report vs GL Sc")
                    Set myWs3 = myWb.Worksheets("Exception List")
                    myWs1.Cells.Delete
                    myWs2.Cells.Delete
                    myWs3.Cells.Delete
                    
                    Set refWb = Workbooks.Open(curPath & tempStr)
                    Set refWs = refWb.Worksheets(1)
                    
                    GSTReport refWb, myWs1, myWs2, myWs3
                        
                    'refWs.Cells.Copy myWs1.Range("A1")
                    
                    refWb.SaveAs curPath & tempStr & "x", FileFormat:=51
                    refWb.Close True
                    Set refWb = Nothing
                    Set refWs = Nothing
                    
                    Set myWs1 = Nothing
                    Set myWs2 = Nothing
                    
                End If
                
            Else
                
                
                
                
            End If
            tempStr = Dir
        Wend
        
        
        Set myWs1 = myWb.Worksheets("Col E Adj G11")
        
        percProgress = percProgress + 0.01
        Call ChangeProgress("Input Form file: " & myFileUsed & " - Set Col E")
        
        
        With myWs1
            
            If xCount = 1 Or xCount = 3 Then
                .Range("C73").Formula = "='Input GL (20610101)'!U21"
                .Range("C74").Formula = "='Input GL (20610101)'!U22"
                .Range("C76").Formula = "='Input GL (20610101)'!U5"
                .Range("C77").Formula = "='Output GL (20610102)'!U7"
                
            ElseIf xCount = 5 Or xCount = 6 Or xCount = 7 Then
                .Range("C68").Formula = "=-'Output GL (20610102)'!U7"
                .Range("C69").Formula = "=-C68"
                .Range("C71").Formula = "='Sales GL (Toll Cos)'!U2"
                
            ElseIf xCount = 8 Then
                'QCLNG
                
                myWb.Worksheets("GST Report - QCLNG").Range("L3:L7").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,K3,'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,""5036""))"
                myWb.Worksheets("GST Report - QCLNG").Range("M3:M7").Formula = "=SUMIFS('Input GL (20610101)'!$I:$I,'Input GL (20610101)'!$A:$A,""5036"",'Input GL (20610101)'!$O:$O,K3)"
                myWb.Worksheets("GST Report - QCLNG").Range("N3:N7").Formula = "=M3-L3"
                
                myWb.Worksheets("GST Report - QCLNG").Range("L12:L16").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,K12,'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,""5039""))"
                myWb.Worksheets("GST Report - QCLNG").Range("M12:M16").Formula = "=SUMIFS('Input GL (20610101)'!$I:$I,'Input GL (20610101)'!$A:$A,""5039"",'Input GL (20610101)'!$O:$O,K12)"
                myWb.Worksheets("GST Report - QCLNG").Range("N12:N16").Formula = "=M12-L12"
                
                
                
                
                .Range("C24").Formula = "='GST Report - QCLNG'!N3"
                .Range("C26").Formula = "='GST Report - QCLNG'!N4"
                .Range("C28").Formula = "='GST Report - QCLNG'!N5"
                .Range("C30").Formula = "='GST Report - QCLNG'!N6"
                .Range("C32").Formula = "='GST Report - QCLNG'!N12"
                .Range("C34").Formula = "='GST Report - QCLNG'!N13"
                .Range("C36").Formula = "='GST Report - QCLNG'!N14"
                .Range("C38").Formula = "='GST Report - QCLNG'!N15"
                .Range("C40").Formula = "='GST Report - QCLNG'!N7"
                .Range("C42").Formula = "='GST Report - QCLNG'!N16"
                .Range("C48").Formula = "='Output GL (20610102)'!U9" '"=VLOOKUP(""T1 UJV"",'Output GL (20610102)'!T:U,2,0)"
                .Range("C50").Formula = "='Output GL (20610102)'!U10" '"=VLOOKUP(""T2 UJV"",'Output GL (20610102)'!T:U,2,0)"
            
            ElseIf xCount = 9 Then
                'BGIA
                
                
                
                
            ElseIf xCount = 10 Then
                'QGC
                .Range("C10").Formula = "='Input GL (20610101)'!U12" '"=VLOOKUP(""ZI"",'Input GL (20610101)'!T:U,2,0)"
                .Range("C12").Formula = "='Input GL (20610101)'!U13" '"=VLOOKUP(""GST on asset disposal proceeds"",'Input GL (20610101)'!T:U,2,0)"
                .Range("C14").Formula = "='Input GL (20610101)'!U16" '"=VLOOKUP(""Chinchilla Accm Tax"",'Input GL (20610101)'!T:U,2,0)"
                .Range("C16").Formula = "='Input GL (20610101)'!U19" '"=VLOOKUP(""Visa Control Account (Doc No: 901203879)"",'Input GL (20610101)'!T:U,2,0)"
                .Range("C18").Formula = "='Input GL (20610101)'!U7" '"=VLOOKUP(""B6"",'Input GL (20610101)'!T:U,2,0)"
                .Range("C20").Formula = "='Input GL (20610101)'!U20" '"=VLOOKUP(""Recode transaction for Tax Code B9 (Doc No: 1600000253)"",'Input GL (20610101)'!T:U,2,0)"
                .Range("C22").Formula = "='Input GL (20610101)'!U10" '"=VLOOKUP(""P0"",'Input GL (20610101)'!T:U,2,0)"
                .Range("C46").Formula = "='Output GL (20610102)'!U8" '"=VLOOKUP(""GST on asset disposal proceeds"",'Output GL (20610102)'!T:U,2,0)"
                .Range("C52").Formula = "='Output GL (20610102)'!U6" '"=VLOOKUP(""APA"",'Output GL (20610102)'!T:U,2,0)"
            End If
            
        End With
        
        
        percProgress = percProgress + 0.03
        Call ChangeProgress("Input Form file: " & myFileUsed & " - Set Data Entry")
        
        Select Case xCount
        Case 1, 3
            For Each ws In myWb.Worksheets
                If Left(ws.Name, 10) = "Data Entry" Then
                    With ws
                        .Range("C19").Formula = "=SUM('Col E Adj G11'!G73,'Col E Adj G11'!G74,'Col E Adj G11'!G76,'Col E Adj G11'!G77)"
                        
                        .Range("G64").Formula = "=SUMIF('Input GL (20610101)'!$Q:$Q,B64,'Input GL (20610101)'!$E:$E)"
                        .Range("G71").Formula = "=SUMIF('Output GL (20610102)'!$Q:$Q,B71,'Output GL (20610102)'!$E:$E)"
                        
                        .Range("G92").Formula = "=SUMIF('Clearing GL (20610000)'!$Q:$Q,B92,'Clearing GL (20610000)'!$I:$I)"
                    End With
                    Exit For
                End If
            Next
            
        Case 5 To 7
            For Each ws In myWb.Worksheets
                If Left(ws.Name, 10) = "Data Entry" Then
                    With ws
                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,{""S1"",""S3"",""S4"",""S7""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$G:$G,'GST Report'!$C:$C,{""S1"",""S3"",""S4"",""S7""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,""S4"",'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$G:$G,'GST Report'!$C:$C,""S4"",'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,""S9"",'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$G:$G,'GST Report'!$C:$C,""S9"",'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5))"
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,{""B2"",""B3""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$H:$H,'GST Report'!$C:$C,{""B2"",""B3""},'GST Report'!$D:$D,""<>"",'GST Report'!B:B,B5))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$H:$H,'GST Report'!$C:$C,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,{""B3"",""B6"",""B9""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$H:$H,'GST Report'!$C:$C,{""B3"",""B6"",""B9""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5))"
                        
                        
                        .Range("G9").Formula = "='Col E Adj G11'!G68 + G10"
                        .Range("G10").Formula = "='Col E Adj G11'!C71"
                        .Range("G20").Formula = "='Col E Adj G11'!G69"
                        
                        .Range("L19").Formula = "=J19+'Col E Adj G11'!G4"
                        .Range("J31").Formula = "='Col E Adj G11'!F4"
                        .Range("J35").Formula = "=-'Col E Adj G11'!F4"
                        .Range("J36").Formula = "='Col E Adj G11'!C6"
                        .Range("J37").Formula = "='Col E Adj G11'!C38"
                        
                        .Range("G64").Formula = "=SUMIF('Input GL (20610101)'!$Q:$Q,B64,'Input GL (20610101)'!$E:$E)"
                        .Range("G65").Formula = "=-SUM('Col E Adj G11'!F29,'Col E Adj G11'!F34)"
                        
                        .Range("G71").Formula = "=SUMIF('Output GL (20610102)'!$Q:$Q,B71,'Output GL (20610102)'!$E:$E)"
                        .Range("G72").Formula = "=SUM('Col E Adj G11'!F30,'Col E Adj G11'!F35)"
                        
                        .Range("G92").Formula = "=SUMIF('Clearing GL (20610000)'!$Q:$Q,B92,'Clearing GL (20610000)'!$I:$I)"
                    End With
                    Exit For
                End If
            Next
                        
        Case 8
            For Each ws In myWb.Worksheets
                If ws.Name = "Data Entry - QCLNG Group" Then
                    With ws
                        .Range("C9").Formula = "=SUM('Data Entry - 1127'!C9,'Data Entry - 1113'!C9,'Data Entry - 5039'!C9,'Data Entry - 5036'!C9)"
                        .Range("G9").Formula = "=SUM('Data Entry - 1127'!G9,'Data Entry - 1113'!G9,'Data Entry - 5039'!G9,'Data Entry - 5036'!G9)"
                        
                        .Range("C10").Formula = "=SUM('Data Entry - 1127'!C10,'Data Entry - 1113'!C10,'Data Entry - 5039'!C10,'Data Entry - 5036'!C10)"
                        .Range("C11").Formula = "=SUM('Data Entry - 1127'!C11,'Data Entry - 1113'!C11,'Data Entry - 5039'!C11,'Data Entry - 5036'!C11)"
                        .Range("C12").Formula = "=SUM('Data Entry - 1127'!C12,'Data Entry - 1113'!C12,'Data Entry - 5039'!C12,'Data Entry - 5036'!C12)"
                        .Range("C15").Formula = "=SUM('Data Entry - 1127'!C15,'Data Entry - 1113'!C15,'Data Entry - 5039'!C15,'Data Entry - 5036'!C15)"
                        
                        .Range("C19").Formula = "=SUM('Data Entry - 1127'!C19,'Data Entry - 1113'!C19,'Data Entry - 5039'!C19,'Data Entry - 5036'!C19)"
                        .Range("G19").Formula = "=SUM('Data Entry - 1127'!G19,'Data Entry - 1113'!G19,'Data Entry - 5039'!C19,'Data Entry - 5036'!G19)"
                        
                        .Range("C20").Formula = "=SUM('Data Entry - 1127'!C20,'Data Entry - 1113'!C20,'Data Entry - 5039'!C20,'Data Entry - 5036'!C20)"
                        .Range("G20").Formula = "=SUM('Data Entry - 1127'!G20,'Data Entry - 1113'!G20,'Data Entry - 5039'!G20,'Data Entry - 5036'!G20)"
                        .Range("H20").Formula = "=SUM('Data Entry - 1127'!H20,'Data Entry - 1113'!H20,'Data Entry - 5039'!H20,'Data Entry - 5036'!H20)"
                        
                        .Range("C22").Formula = "=SUM('Data Entry - 1127'!C22,'Data Entry - 1113'!C22,'Data Entry - 5039'!C22,'Data Entry - 5036'!C22)"
                        .Range("C23").Formula = "=SUM('Data Entry - 1127'!C23,'Data Entry - 1113'!C23,'Data Entry - 5039'!C23,'Data Entry - 5036'!C23)"
                        .Range("C24").Formula = "=SUM('Data Entry - 1127'!C24,'Data Entry - 1113'!C24,'Data Entry - 5039'!C24,'Data Entry - 5036'!C24)"
                        
                        .Range("J31").Formula = "='Col E Adj G11'!F4"
                        .Range("J35").Formula = "=-'Col E Adj G11'!F4"
                        .Range("J36").Formula = "='Col E Adj G11'!C6"
                        .Range("J37").Formula = "='Col E Adj G11'!C54"
                        
                    End With
                    
                ElseIf ws.Name = "Data Entry - 5036" Then
                    With ws
                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""S1"",""S3"",""S4"",""S7""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!$C:$C,{""S1"",""S3"",""S4"",""S7""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,""S9"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!$C:$C,""S9"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""B2"",""B3""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,{""B2"",""B3""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""B3"",""B6"",""B9""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,{""B3"",""B6"",""B9""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        
                        .Range("L9").Formula = "=SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$A:$A,""5000"",'GST Report - QCLNG'!$C:$C,""S0"",'GST Report - QCLNG'!$C:$C,""S1"",'GST Report - QCLNG'!$C:$C,""S2"",'GST Report - QCLNG'!$C:$C,""S3"",'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$C:$C,""S7"")+SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!B:B,""5000"",'GST Report - QCLNG'!$G:$G,""S0"",'GST Report - QCLNG'!$C:$C,""S1"",'GST Report - QCLNG'!$C:$C,""S2"",'GST Report - QCLNG'!$C:$C,""S3"",'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$C:$C,""S7"")"
                        
                        
                        .Range("G9").Formula = "='Col E Adj G11'!G40 + 'Col E Adj G11'!G48 + 'Col E Adj G11'!G50"
                        .Range("G19").Formula = "='Col E Adj G11'!G4"
                        .Range("G20").Formula = "='Col E Adj G11'!G24"
                        
                        .Range("J35").Formula = "=-'Col E Adj G11'!F4"
                        .Range("J36").Formula = "='Col E Adj G11'!C6"
                        .Range("J37").Formula = "='Col E Adj G11'!C66"
                        
                    End With
                    
                ElseIf ws.Name = "Data Entry - 5039" Then
                    With ws
                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""S1"",""S3"",""S4"",""S7""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!$C:$C,{""S1"",""S3"",""S4"",""S7""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,""S9"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!$C:$C,""S9"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""B2"",""B3""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,{""B2"",""B3""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""B3"",""B6"",""B9""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,{""B3"",""B6"",""B9""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        
                        .Range("G9").Formula = "='Col E Adj G11'!G42"
                        .Range("G19").Formula = "='Col E Adj G11'!G32"
                        
                        .Range("L9").Formula = "=SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$A:$A,""5000"",'GST Report - QCLNG'!$C:$C,""S0"",'GST Report - QCLNG'!$C:$C,""S1"",'GST Report - QCLNG'!$C:$C,""S2"",'GST Report - QCLNG'!$C:$C,""S3"",'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$C:$C,""S7"")+SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!B:B,""5000"",'GST Report - QCLNG'!$G:$G,""S0"",'GST Report - QCLNG'!$C:$C,""S1"",'GST Report - QCLNG'!$C:$C,""S2"",'GST Report - QCLNG'!$C:$C,""S3"",'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$C:$C,""S7"")"
                        .Range("J37").Formula = "='Col E Adj G11'!C66"
                        
                    End With
                ElseIf Left(ws.Name, 10) = "Data Entry" Then
                    With ws
                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""S1"",""S3"",""S4"",""S7""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!$C:$C,{""S1"",""S3"",""S4"",""S7""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,""S9"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!$C:$C,""S9"",'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""B2"",""B3""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,{""B2"",""B3""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$C:$C,{""B3"",""B6"",""B9""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5),SUMIFS('GST Report - QCLNG'!$H:$H,'GST Report - QCLNG'!$C:$C,{""B3"",""B6"",""B9""},'GST Report - QCLNG'!$D:$D,""<>"",'GST Report - QCLNG'!$A:$A,B5))"
                         
                        .Range("L9").Formula = "=SUMIFS('GST Report - QCLNG'!$F:$F,'GST Report - QCLNG'!$A:$A,""5000"",'GST Report - QCLNG'!$C:$C,""S0"",'GST Report - QCLNG'!$C:$C,""S1"",'GST Report - QCLNG'!$C:$C,""S2"",'GST Report - QCLNG'!$C:$C,""S3"",'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$C:$C,""S7"")+SUMIFS('GST Report - QCLNG'!$G:$G,'GST Report - QCLNG'!B:B,""5000"",'GST Report - QCLNG'!$G:$G,""S0"",'GST Report - QCLNG'!$C:$C,""S1"",'GST Report - QCLNG'!$C:$C,""S2"",'GST Report - QCLNG'!$C:$C,""S3"",'GST Report - QCLNG'!$C:$C,""S4"",'GST Report - QCLNG'!$C:$C,""S7"")"
                        .Range("J37").Formula = "='Col E Adj G11'!C66"
                    End With
                    
                End If
                
            Next
            
        Case 9
            For Each ws In myWb.Worksheets
                If ws.Name = "Data Entry - BGIA Group" Then
                    With ws
                        .Range("C9").Formula = "=SUM('Data Entry - 1122'!C9,'Data Entry - 1116'!C9,'Data Entry - 1112'!C9,'Data Entry - 1106'!C9,'Data Entry - 1101'!C9,'Data Entry - 1100'!C9)"
                        .Range("C10").Formula = "=SUM('Data Entry - 1122'!C10,'Data Entry - 1116'!C10,'Data Entry - 1112'!C10,'Data Entry - 1106'!C10,'Data Entry - 1101'!C10,'Data Entry - 1100'!C10)"
                        .Range("C15").Formula = "=SUM('Data Entry - 1122'!C15,'Data Entry - 1116'!C15,'Data Entry - 1112'!C15,'Data Entry - 1106'!C15,'Data Entry - 1101'!C15,'Data Entry - 1100'!C15)"
                        
                        .Range("C19").Formula = "=SUM('Data Entry - 1122'!C19,'Data Entry - 1116'!C19,'Data Entry - 1112'!C19,'Data Entry - 1106'!C19,'Data Entry - 1101'!C19,'Data Entry - 1100'!C19)"
                        .Range("C20").Formula = "=SUM('Data Entry - 1122'!C20,'Data Entry - 1116'!C20,'Data Entry - 1112'!C20,'Data Entry - 1106'!C20,'Data Entry - 1101'!C20,'Data Entry - 1100'!C20)"
                        .Range("C23").Formula = "=SUM('Data Entry - 1122'!C23,'Data Entry - 1116'!C23,'Data Entry - 1112'!C23,'Data Entry - 1106'!C23,'Data Entry - 1101'!C23,'Data Entry - 1100'!C23)"
                        
                        .Range("J35").Formula = "=-'Col E Adj G11'!F4"
                        .Range("J36").Formula = "='Col E Adj G11'!C6"
                        .Range("J37").Formula = "='Col E Adj G11'!C54"
                        
                        
                        
                    End With
                    
                ElseIf ws.Name = "Data Entry - 1100" Then
                    With ws
                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$G:$G,'GST Report - BGIA'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,""S4"",'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$G:$G,'GST Report - BGIA'!$B:$B,""S4"",'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,""S9"",'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$G:$G,'GST Report - BGIA'!$B:$B,""S9"",'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,{""B2"",""B3""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$H:$H,'GST Report - BGIA'!$B:$B,{""B2"",""B3""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$H:$H,'GST Report - BGIA'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$H:$H,'GST Report - BGIA'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        
                        .Range("L19").Formula = "=SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$A:$A,""5000"",'GST Report - BGIA'!$B:$B,""S0"",'GST Report - BGIA'!$B:$B,""S1"",'GST Report - BGIA'!$B:$B,""S2"",'GST Report - BGIA'!$B:$B,""S3"",'GST Report - BGIA'!$B:$B,""S4"",'GST Report - BGIA'!$B:$B,""S7"")+SUMIFS('GST Report - BGIA'!$G:$G,'GST Report - BGIA'!B:B,""5000"",'GST Report - BGIA'!$G:$G,""S0"",'GST Report - BGIA'!$B:$B,""S1"",'GST Report - BGIA'!$B:$B,""S2"",'GST Report - BGIA'!$B:$B,""S3"",'GST Report - BGIA'!$B:$B,""S4"",'GST Report - BGIA'!$B:$B,""S7"")"
                        
                        .Range("J35").Formula = "=-'Col E Adj G11'!F4"
                        .Range("J36").Formula = "='Col E Adj G11'!C6"
                        .Range("J37").Formula = "='Col E Adj G11'!C54"
                        
                    End With
                ElseIf Left(ws.Name, 10) = "Data Entry" Then
                    With ws
                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$G:$G,'GST Report - BGIA'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,""S4"",'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$G:$G,'GST Report - BGIA'!$B:$B,""S4"",'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,""S9"",'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$G:$G,'GST Report - BGIA'!$B:$B,""S9"",'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,{""B2"",""B3""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$H:$H,'GST Report - BGIA'!$B:$B,{""B2"",""B3""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$H:$H,'GST Report - BGIA'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5),SUMIFS('GST Report - BGIA'!$H:$H,'GST Report - BGIA'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - BGIA'!$D:$D,""<>"",'GST Report - BGIA'!$A:$A,B5))"
                        
                        .Range("L19").Formula = "=SUMIFS('GST Report - BGIA'!$F:$F,'GST Report - BGIA'!$A:$A,""5000"",'GST Report - BGIA'!$B:$B,""S0"",'GST Report - BGIA'!$B:$B,""S1"",'GST Report - BGIA'!$B:$B,""S2"",'GST Report - BGIA'!$B:$B,""S3"",'GST Report - BGIA'!$B:$B,""S4"",'GST Report - BGIA'!$B:$B,""S7"")+SUMIFS('GST Report - BGIA'!$G:$G,'GST Report - BGIA'!B:B,""5000"",'GST Report - BGIA'!$G:$G,""S0"",'GST Report - BGIA'!$B:$B,""S1"",'GST Report - BGIA'!$B:$B,""S2"",'GST Report - BGIA'!$B:$B,""S3"",'GST Report - BGIA'!$B:$B,""S4"",'GST Report - BGIA'!$B:$B,""S7"")"
                        
                    End With
                End If
            Next
        Case 10
            For Each ws In myWb.Worksheets
                If ws.Name = "Data Entry - QGC Group" Then
                    With ws
                        .Range("C9").Formula = getName("C9")
                        .Range("C10").Formula = getName("C10")
                        .Range("C11").Formula = getName("C11")
                        .Range("C12").Formula = getName("C12")
                        .Range("C15").Formula = getName("C15")
                        
                        .Range("G23").Formula = getName("G23")
                        
                        .Range("C19").Formula = getName("C19")
                        .Range("G19").Formula = getName("G19")
                        
                        .Range("C20").Formula = getName("C20")
                        
                        .Range("C22").Formula = getName("C22")
                        .Range("C23").Formula = getName("C23")
                        .Range("C24").Formula = getName("C24")
                        
                        .Range("G9").Formula = "=SUM('Col E Adj G11'!G26,'Col E Adj G11'!G32,'Col E Adj G11'!G34,-SUM('Col E Adj G11'!$G$38,'Col E Adj G11'!G43))"
                        .Range("G20").Formula = "=SUM('Col E Adj G11'!G10,'Col E Adj G11'!G12,'Col E Adj G11'!G14,'Col E Adj G11'!G16,'Col E Adj G11'!G37,'Col E Adj G11'!G42)"
                        
                        .Range("J31").Formula = "='Col E Adj G11'!F4"
                        .Range("J35").Formula = "=-'Col E Adj G11'!F4"
                        .Range("J36").Formula = "='Col E Adj G11'!C6"
                        .Range("J37").Formula = "='Col E Adj G11'!C46"
                        
                        .Range("G64").Formula = "=SUMIF('Input GL (20610101)'!Q:Q,'Data Entry - QGC Group'!B64,'Input GL (20610101)'!$E:$E)"
                        .Range("G65").Formula = "=-SUM('Col E Adj G11'!F37,'Col E Adj G11'!F42)"
                        .Range("G71").Formula = "=SUMIF('Output GL (20610102)'!Q:Q,'Data Entry - QGC Group'!B71,'Output GL (20610102)'!$E:$E)"
                        .Range("G72").Formula = "=SUM('Col E Adj G11'!F38,'Col E Adj G11'!F43)"
                        
                        .Range("G78").Formula = "='Data Entry - JV 171'!J38"
                        .Range("G79").Formula = "='Data Entry - QGC JV'!J38"
                        
                        .Range("G92").Formula = "=SUMIF('Clearing GL (20610000)'!Q:Q,B92,'Clearing GL (20610000)'!$E:$E)"
                        
                    End With
                    
                ElseIf ws.Name = "Data Entry - 5000" Then
                    With ws

                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$G:$G,'GST Report - QGC'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,""S4"",'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$G:$G,'GST Report - QGC'!$B:$B,""S4"",'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,""S9"",'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$G:$G,'GST Report - QGC'!$B:$B,""S9"",'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,{""B2"",""B3""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$H:$H,'GST Report - QGC'!$B:$B,{""B2"",""B3""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$H:$H,'GST Report - QGC'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$H:$H,'GST Report - QGC'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))-('GST Report - QGC'!H6*11)"
                        
                        .Range("G9").Formula = "=SUM('Col E Adj G11'!G26,'Col E Adj G11'!G32,'Col E Adj G11'!G34,-SUM('Col E Adj G11'!$G$38,'Col E Adj G11'!G43))"
                        .Range("G20").Formula = "=SUM('Col E Adj G11'!G10,'Col E Adj G11'!G12,'Col E Adj G11'!G14,'Col E Adj G11'!G16,'Col E Adj G11'!G37,'Col E Adj G11'!G42)"
                        
                        .Range("G23").Formula = "=SUM(SUMIFS('GST Report - QGC'!H:H,'GST Report - QGC'!B:B,{""B3"",""B6"",""B9""},'GST Report - QGC'!A:A,B5))*-11"
                        
                        
                        .Range("L9").Formula = "=SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$A:$A,""5000"",'GST Report - QGC'!$B:$B,""S0"",'GST Report - QGC'!$B:$B,""S1"",'GST Report - QGC'!$B:$B,""S2"",'GST Report - QGC'!$B:$B,""S3"",'GST Report - QGC'!$B:$B,""S4"",'GST Report - QGC'!$B:$B,""S7"")+SUMIFS('GST Report - QGC'!$G:$G,'GST Report - QGC'!B:B,""5000"",'GST Report - QGC'!$G:$G,""S0"",'GST Report - QGC'!$B:$B,""S1"",'GST Report - QGC'!$B:$B,""S2"",'GST Report - QGC'!$B:$B,""S3"",'GST Report - QGC'!$B:$B,""S4"",'GST Report - QGC'!$B:$B,""S7"")"
                        .Range("L19").Formula = "=J19+'Col E Adj G11'!G4"
                        
                        .Range("J31").Formula = "='Col E Adj G11'!F4"
                        .Range("J35").Formula = "=-'Col E Adj G11'!F4"
                        .Range("J36").Formula = "='Col E Adj G11'!C6"
                        .Range("J37").Formula = "='Col E Adj G11'!C46"
                        
                    End With
                    
                ElseIf ws.Name = "Data Entry - QGC JV" Then
                    With ws
                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report - QGC JV'!$F:$F,'GST Report - QGC JV'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - QGC JV'!$D:$D,""<>""),SUMIFS('GST Report - QGC JV'!$G:$G,'GST Report - QGC JV'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - QGC JV'!$D:$D,""<>""))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report - QGC JV'!$F:$F,'GST Report - QGC JV'!$B:$B,""S4"",'GST Report - QGC JV'!$D:$D,""<>""),SUMIFS('GST Report - QGC JV'!$G:$G,'GST Report - QGC JV'!$B:$B,""S4"",'GST Report - QGC JV'!$D:$D,""<>""))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report - QGC JV'!$F:$F,'GST Report - QGC JV'!$B:$B,""S9"",'GST Report - QGC JV'!$D:$D,""<>""),SUMIFS('GST Report - QGC JV'!$G:$G,'GST Report - QGC JV'!$B:$B,""S9"",'GST Report - QGC JV'!$D:$D,""<>""))"
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report - QGC JV'!$F:$F,'GST Report - QGC JV'!$B:$B,{""B2"",""B3""},'GST Report - QGC JV'!$D:$D,""<>""),SUMIFS('GST Report - QGC JV'!$H:$H,'GST Report - QGC JV'!$B:$B,{""B2"",""B3""},'GST Report - QGC JV'!$D:$D,""<>""))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report - QGC JV'!$F:$F,'GST Report - QGC JV'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QGC JV'!$D:$D,""<>""),SUMIFS('GST Report - QGC JV'!$H:$H,'GST Report - QGC JV'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QGC JV'!$D:$D,""<>""))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report - QGC JV'!$F:$F,'GST Report - QGC JV'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - QGC JV'!$D:$D,""<>""),SUMIFS('GST Report - QGC JV'!$H:$H,'GST Report - QGC JV'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - QGC JV'!$D:$D,""<>""))-('GST Report - QGC JV'!H12*11)"
                        
                        .Range("L9").Formula = "=SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$A:$A,""5000"",'GST Report - QGC'!$B:$B,""S0"",'GST Report - QGC'!$B:$B,""S1"",'GST Report - QGC'!$B:$B,""S2"",'GST Report - QGC'!$B:$B,""S3"",'GST Report - QGC'!$B:$B,""S4"",'GST Report - QGC'!$B:$B,""S7"")+SUMIFS('GST Report - QGC'!$G:$G,'GST Report - QGC'!B:B,""5000"",'GST Report - QGC'!$G:$G,""S0"",'GST Report - QGC'!$B:$B,""S1"",'GST Report - QGC'!$B:$B,""S2"",'GST Report - QGC'!$B:$B,""S3"",'GST Report - QGC'!$B:$B,""S4"",'GST Report - QGC'!$B:$B,""S7"")"
                        
                        .Range("G20").Formula = "='Col E Adj G11'!G22"
                        .Range("J37").Formula = "='Col E Adj G11'!C46"
                        
                    End With
                
                ElseIf ws.Name = "Data Entry - JV 171" Then
                    With ws
                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report - JV 171'!$F:$F,'GST Report - JV 171'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - JV 171'!$D:$D,""<>""),SUMIFS('GST Report - JV 171'!$G:$G,'GST Report - JV 171'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - JV 171'!$D:$D,""<>""))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report - JV 171'!$F:$F,'GST Report - JV 171'!$B:$B,""S4"",'GST Report - JV 171'!$D:$D,""<>""),SUMIFS('GST Report - JV 171'!$G:$G,'GST Report - JV 171'!$B:$B,""S4"",'GST Report - JV 171'!$D:$D,""<>""))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report - JV 171'!$F:$F,'GST Report - JV 171'!$B:$B,""S9"",'GST Report - JV 171'!$D:$D,""<>""),SUMIFS('GST Report - JV 171'!$G:$G,'GST Report - JV 171'!$B:$B,""S9"",'GST Report - JV 171'!$D:$D,""<>""))"
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report - JV 171'!$F:$F,'GST Report - JV 171'!$B:$B,{""B2"",""B3""},'GST Report - JV 171'!$D:$D,""<>""),SUMIFS('GST Report - JV 171'!$H:$H,'GST Report - JV 171'!$B:$B,{""B2"",""B3""},'GST Report - JV 171'!$D:$D,""<>""))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report - JV 171'!$F:$F,'GST Report - JV 171'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - JV 171'!$D:$D,""<>""),SUMIFS('GST Report - JV 171'!$H:$H,'GST Report - JV 171'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - JV 171'!$D:$D,""<>""))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report - JV 171'!$F:$F,'GST Report - JV 171'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - JV 171'!$D:$D,""<>""),SUMIFS('GST Report - JV 171'!$H:$H,'GST Report - JV 171'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - JV 171'!$D:$D,""<>""))"
                        
                        .Range("J37").Formula = "='Col E Adj G11'!C46"
                        
                    End With
                
                ElseIf Left(ws.Name, 10) = "Data Entry" Then
                    With ws
                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$G:$G,'GST Report - QGC'!$B:$B,{""S1"",""S3"",""S4"",""S7""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,""S4"",'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$G:$G,'GST Report - QGC'!$B:$B,""S4"",'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,""S9"",'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$G:$G,'GST Report - QGC'!$B:$B,""S9"",'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        
                        .Range("G23").Formula = "=SUM(SUMIFS('GST Report - QGC'!H:H,'GST Report - QGC'!B:B,{""B3"",""B6"",""B9""},'GST Report - QGC'!A:A,B5))*-11"
                        
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,{""B2"",""B3""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$H:$H,'GST Report - QGC'!$B:$B,{""B2"",""B3""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$H:$H,'GST Report - QGC'!$B:$B,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5),SUMIFS('GST Report - QGC'!$H:$H,'GST Report - QGC'!$B:$B,{""B3"",""B6"",""B9""},'GST Report - QGC'!$D:$D,""<>"",'GST Report - QGC'!$A:$A,B5))"
                        
                        .Range("L19").Formula = "=J19+'Col E Adj G11'!G4"
                        .Range("L9").Formula = "=SUMIFS('GST Report - QGC'!$F:$F,'GST Report - QGC'!$A:$A,""5000"",'GST Report - QGC'!$B:$B,""S0"",'GST Report - QGC'!$B:$B,""S1"",'GST Report - QGC'!$B:$B,""S2"",'GST Report - QGC'!$B:$B,""S3"",'GST Report - QGC'!$B:$B,""S4"",'GST Report - QGC'!$B:$B,""S7"")+SUMIFS('GST Report - QGC'!$G:$G,'GST Report - QGC'!B:B,""5000"",'GST Report - QGC'!$G:$G,""S0"",'GST Report - QGC'!$B:$B,""S1"",'GST Report - QGC'!$B:$B,""S2"",'GST Report - QGC'!$B:$B,""S3"",'GST Report - QGC'!$B:$B,""S4"",'GST Report - QGC'!$B:$B,""S7"")"
                        
                    End With
                    
                End If
            Next
            
        Case Else
            For Each ws In myWb.Worksheets
                If Left(ws.Name, 10) = "Data Entry" Then
                    With ws
                        .Range("C9").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,{""S1"",""S3"",""S4"",""S7""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$G:$G,'GST Report'!$C:$C,{""S1"",""S3"",""S4"",""S7""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5))"
                        .Range("C10").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,""S4"",'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$G:$G,'GST Report'!$C:$C,""S4"",'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5))"
                        .Range("C15").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,""S9"",'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$G:$G,'GST Report'!$C:$C,""S9"",'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5))"
                        
                        .Range("C19").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,{""B2"",""B3""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$H:$H,'GST Report'!$C:$C,{""B2"",""B3""},'GST Report'!$D:$D,""<>"",'GST Report'!B:B,B5))"
                        .Range("C20").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$H:$H,'GST Report'!$C:$C,{""B0"",""B1"",""B6"",""B9"",""IQ""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5))"
                        .Range("C23").Formula = "=SUM(SUMIFS('GST Report'!$F:$F,'GST Report'!$C:$C,{""B3"",""B6"",""B9""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5),SUMIFS('GST Report'!$H:$H,'GST Report'!$C:$C,{""B3"",""B6"",""B9""},'GST Report'!$D:$D,""<>"",'GST Report'!$A:$A,B5))"
                        
                        
                        .Range("G64").Formula = "=SUMIF('Input GL (20610101)'!$Q:$Q,B64,'Input GL (20610101)'!$E:$E)"
                        .Range("G71").Formula = "=SUMIF('Output GL (20610102)'!$Q:$Q,B71,'Output GL (20610102)'!$E:$E)"
                        .Range("G92").Formula = "=SUMIF('Clearing GL (20610000)'!$Q:$Q,B92,'Clearing GL (20610000)'!$I:$I)"
                        
                        .Range("L19").Formula = "=J19"
                        
                    End With
                    Exit For
                End If
            Next
            
        End Select
    
        
        myWb.Save
        myWb.Close True
        
skippedItem:
    Next
    
    
End Sub









Function getName(myCell As String)
    getName = ""
    getName = "=SUM('Data Entry - 5044'!" & myCell
    getName = getName & ",'Data Entry - 5043'!" & myCell
    getName = getName & ",'Data Entry - 5042'!" & myCell
    getName = getName & ",'Data Entry - 5041'!" & myCell
    getName = getName & ",'Data Entry - 5040'!" & myCell
    getName = getName & ",'Data Entry - 5032'!" & myCell
    getName = getName & ",'Data Entry - 5029'!" & myCell
    getName = getName & ",'Data Entry - 5028'!" & myCell
    getName = getName & ",'Data Entry - 5027'!" & myCell
    getName = getName & ",'Data Entry - 5026'!" & myCell
    getName = getName & ",'Data Entry - 5025'!" & myCell
    getName = getName & ",'Data Entry - 5024'!" & myCell
    getName = getName & ",'Data Entry - 5023'!" & myCell
    getName = getName & ",'Data Entry - 5022'!" & myCell
    getName = getName & ",'Data Entry - 5021'!" & myCell
    getName = getName & ",'Data Entry - 5020'!" & myCell
    getName = getName & ",'Data Entry - 5019'!" & myCell
    getName = getName & ",'Data Entry - 5018'!" & myCell
    getName = getName & ",'Data Entry - 5017'!" & myCell
    getName = getName & ",'Data Entry - 5016'!" & myCell
    getName = getName & ",'Data Entry - 5015'!" & myCell
    getName = getName & ",'Data Entry - 5014'!" & myCell
    getName = getName & ",'Data Entry - 5013'!" & myCell
    getName = getName & ",'Data Entry - 5012'!" & myCell
    getName = getName & ",'Data Entry - 5011'!" & myCell
    getName = getName & ",'Data Entry - 5009'!" & myCell
    getName = getName & ",'Data Entry - 5008'!" & myCell
    getName = getName & ",'Data Entry - 5007'!" & myCell
    getName = getName & ",'Data Entry - 5006'!" & myCell
    getName = getName & ",'Data Entry - 5005'!" & myCell
    getName = getName & ",'Data Entry - 5004'!" & myCell
    getName = getName & ",'Data Entry - 5003'!" & myCell
    getName = getName & ",'Data Entry - 5002'!" & myCell
    getName = getName & ",'Data Entry - 5001'!" & myCell
    getName = getName & ",'Data Entry - 5000'!" & myCell
    getName = getName & ")"
            
End Function


Sub GSTReport(wb As Workbook, wsTo1 As Worksheet, wsTo2 As Worksheet, wsTo3 As Worksheet)
    Dim xRow1 As Long, xRow2 As Long, xLastRow As Long
    Dim xRowRef As Long, tempRow As Long, xCol As Long, xLastCol As Long
    Dim xCount As Long
    Dim tempRng1 As Range, tempRng2 As Range, tempRng As Range
    
    Dim tempWs1 As Worksheet, tempWs2 As Worksheet, ws As Worksheet
    
    For Each ws In wb.Worksheets
        If ws.Name = "GST Report" Or ws.Name = "Transaction" Then
            ws.Delete
        End If
    Next
    
    
    
    Set ws = wb.Worksheets(1)
    ws.Copy before:=ws
    ws.Copy before:=ws
    ws.Copy before:=ws
    Set tempWs1 = wb.Worksheets(1)
    Set tempWs2 = wb.Worksheets(2)
    Set ws = wb.Worksheets(3)
    tempWs1.Name = "GST Report"
    ws.Name = "Transaction"
    
    
    Select Case Left(wb.Name, 4)
    Case "5033", "5030", "5031", "5037", "5038", "5036", "QCLN"
        'If Left(wb.Name, 4) = "5038" Then Stop
        If Not tempWs1.Range("E:E").Find("curr.") Is Nothing Then
            xRow1 = tempWs1.Range("E:E").Find("curr.").Row - 1
            If xRow1 > 2 Then
                tempWs1.Range("A1:A" & xRow1).EntireRow.Delete
                tempWs2.Range("A" & xRow1 - 6 & ":A" & tempWs2.Cells(tempWs2.Rows.Count, "B").End(xlUp).Row).EntireRow.Delete
                ws.Range("A" & xRow1 - 6 & ":A" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row).EntireRow.Delete
            End If
        End If
        
        tempWs1.Range("A:F, H:I, K:K, M:M, O:V, X:X, Z:AE, AG:AI, AK:AM").EntireColumn.Delete
        
    Case "BGIA", "JV76", "JV68"
        If Not tempWs1.Range("B:B").Find("curr.") Is Nothing Then
            xRow1 = tempWs1.Range("B:B").Find("Curr.").Row - 1
            
            If xRow1 > 2 Then
                tempWs1.Range("A1:A" & xRow1).EntireRow.Delete
                tempWs2.Range("A" & xRow1 - 4 & ":A" & tempWs2.Cells(tempWs2.Rows.Count, "B").End(xlUp).Row).EntireRow.Delete
                ws.Range("A" & xRow1 - 4 & ":A" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row).EntireRow.Delete
            End If
        End If
        
        tempWs1.Range("A:C, E:E, G:G, I:I, K:Q, S:T, V:Y, AA:AC, AE:AF").EntireColumn.Delete
        
    Case "QGC ", "JV17"
        If Not tempWs1.Range("B:B").Find("curr.") Is Nothing Then
            xRow1 = tempWs1.Range("B:B").Find("Curr.").Row - 1
            
            If xRow1 > 2 Then
                tempWs1.Range("A1:A" & xRow1 - 1).EntireRow.Delete
                tempWs2.Range("A" & xRow1 - 4 & ":A" & tempWs2.Cells(tempWs2.Rows.Count, "B").End(xlUp).Row).EntireRow.Delete
                ws.Range("A" & xRow1 - 4 & ":A" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row).EntireRow.Delete
            End If
        End If
        
        If Left(wb.Name, 6) = "QGC JV" Then
            tempWs1.Range("A:C, E:E, G:G, I:I, K:Q, S:T, V:Y, AA:AC, AE:AF").EntireColumn.Delete
        Else
            tempWs1.Range("A:C, E:E, G:G, I:I, K:P, R:S, U:X, Z:AB, AD:AE").EntireColumn.Delete
        End If
        
        If tempWs1.Range("I2").Value = "" And tempWs1.Range("J2").Value = "" And tempWs1.Range("K2").Value <> "" Then
            tempWs1.Range("I:J").EntireColumn.Delete
        End If
    End Select
    
    tempWs1.Cells.Copy wsTo1.Range("A1")
    
    wsTo2.Range("A2").Value = "Company Code"
    wsTo2.Range("B2").Value = "Tax Code"
    wsTo2.Range("C2").Value = "Joint Venture"
    wsTo2.Range("D2").Value = "Eq. Grp"
    wsTo2.Range("E2").Value = "Post Per"
    wsTo2.Range("F2").Value = "PostDate"
    wsTo2.Range("G2").Value = "Document Number"
    wsTo2.Range("H2").Value = "Reference detail"
    wsTo2.Range("I2").Value = "SKy"
    wsTo2.Range("J2").Value = "Base amount in LC"
    wsTo2.Range("K2").Value = "Input/Output Tax in LC"
    
    wsTo2.Range("L2").Value = "Effective Rate"
    wsTo2.Range("M2").Value = "Rightful GST"
    wsTo2.Range("N2").Value = "Variant"
    wsTo2.Range("O2").Value = "Remark"
    wsTo2.Range("P2").Value = "ADJUSTMENT BAS"
    
    wsTo2.Range("A1:P2").Font.Bold = True
    wsTo2.Range("L:P").Interior.Color = RGB(255, 255, 0)
    
    wsTo2.Range("L:L").NumberFormat = "0%"
    
    
    Do While True
        For xRow1 = 1 To 7
            If tempWs2.Range(NumberToLetter(xRow1) & "1").Value = "CoCd" Or tempWs2.Range(NumberToLetter(xRow1) & "1").Value = "Comp" Then
                Exit Do
            End If
        Next
        tempWs2.Range("A1").EntireRow.Delete
    Loop
    
    
    Do While Application.WorksheetFunction.Max(tempWs2.Cells(tempWs2.Rows.Count, "B").End(xlUp).Row, tempWs2.Cells(tempWs2.Rows.Count, "D").End(xlUp).Row) > 1
        xRow2 = 1
        Do While tempWs2.Range("A" & xRow2 + 1).Value = ""
            xRow2 = xRow2 + 1
            If xRow2 = tempWs2.Cells(tempWs2.Rows.Count, "B").End(xlUp).Row Then Exit Do
        Loop
        
        'to find the last row in the destination sheet
        xRowRef = 0
        For xCount = 1 To 11
            xRowRef = Application.WorksheetFunction.Max(xRowRef, wsTo2.Cells(wsTo2.Rows.Count, "A").End(xlUp).Row)
        Next
        xRowRef = xRowRef + 1
        
        xRow1 = 1
        xLastCol = tempWs2.Cells(1, tempWs2.Columns.Count).End(xlToLeft).Column
        
        
        Set tempRng1 = Nothing
        Set tempRng2 = Nothing
        
        
        Set tempRng1 = tempWs2.Range("A1:G8").Find(what:="CoCd", LookIn:=xlValues, lookat:=xlWhole)
        Set tempRng2 = tempWs2.Range("A1:G8").Find(what:="Comp", LookIn:=xlValues, lookat:=xlWhole)
        
        If tempRng1 Is Nothing Then 'comp structure
            
            For xCol = 1 To xLastCol
                Select Case Trim(tempWs2.Range(NumberToLetter(xCol) & xRow1).Value)
                Case ""
                Case "Comp"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("A" & xRowRef)
                Case "Ta"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("B" & xRowRef)
                Case "Joint"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("C" & xRowRef)
                Case "Eq."
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("D" & xRowRef)
                Case "Post." 'M & Year
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("E" & xRowRef)
                Case "PostDate"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("F" & xRowRef)
                Case "Document"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("G" & xRowRef)
                Case "Reference detail"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("H" & xRowRef)
                Case "SKy"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("I" & xRowRef)
                Case "Base amount"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("J" & xRowRef)
                Case "Tax balance", "Input tax", "Output tax"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("K" & xRowRef)
                End Select
            Next
            
            tempWs2.Range("A1:A" & xRow2 + 3).EntireRow.Delete
            
        ElseIf tempRng2 Is Nothing Then 'cocd structure
            
            For xCol = 1 To xLastCol
                Select Case Trim(tempWs2.Range(NumberToLetter(xCol) & xRow1).Value)
                Case ""
                Case "CoCd"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("A" & xRowRef)
                Case "Tx"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("B" & xRowRef)
                Case "Joint"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("C" & xRowRef)
                Case "Eq."
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("D" & xRowRef)
                    
                Case "M"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("S1")
                Case "Year"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("T1")
                    
                Case "Pstng Date"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("F" & xRowRef)
                Case "DocumentNo"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("G" & xRowRef)
                Case "Reference"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("H" & xRowRef)
                Case "Trs"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("I" & xRowRef)
                Case "Tax base amount"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("J" & xRowRef)
                Case "Output tax", "Input tax", "Balance"
                    tempWs2.Range(NumberToLetter(xCol) & "3:" & NumberToLetter(xCol) & xRow2).Copy wsTo2.Range("K" & xRowRef)
                End Select
            Next
            
            For tempRow = 3 To xRow2
                wsTo2.Range("E" & xRowRef + tempRow - 3).Value = wsTo2.Range("S" & tempRow - 2).Value & wsTo2.Range("T" & tempRow - 2).Value
            Next
            
            wsTo2.Range("S:T").EntireColumn.Clear
            
            tempWs2.Range("A1:A" & xRow2 + 6).EntireRow.Delete
        End If
        
    Loop
    
    
    For xCount = 1 To 11
        xRowRef = Application.WorksheetFunction.Max(xRowRef, wsTo2.Cells(wsTo2.Rows.Count, NumberToLetter(xCount)).End(xlUp).Row)
    Next
    
    For xRow1 = 3 To xRowRef
        If wsTo2.Range("G" & xRow1).Value = "" Then
            wsTo2.Range("G" & xRow1).EntireRow.Delete
            xRow1 = xRow1 - 1
            xRowRef = xRowRef - 1
            If wsTo2.Range("A" & xRow1 + 1).Value = "" And wsTo2.Range("A" & xRow1 + 2).Value = "" And wsTo2.Range("A" & xRow1 + 3).Value = "" And wsTo2.Range("J" & xRow1).Value = "" And wsTo2.Range("J" & xRow1 + 1).Value = "" Then
                Exit For
            End If
            If xRow1 > xRowRef Then
                Exit For
            End If
        End If
    Next
    
    xRowRef = 0
    
    For xCount = 1 To 9
        xRowRef = Application.WorksheetFunction.Max(xRowRef, wsTo2.Cells(wsTo2.Rows.Count, xCount).End(xlUp).Row)
    Next
    
    If xRowRef >= 3 Then
        wsTo2.Range("L3:L" & xRowRef).Formula = "=IFERROR(K3/J3,0)"
        wsTo2.Range("M3:M" & xRowRef).Formula = "=IF(OR(B3=""B1"",B3=""B2"",B3=""B4"",B3=""B5"",B3=""B7"",B3=""B9"",B3=""IQ"",B3=""S1""),J3*10%,0)"
        wsTo2.Range("N3:N" & xRowRef).Formula = "=M3-K3"
        'wsTo2.Range("O3:O" & xRowRef).Value = ""
        wsTo2.Range("P3:P" & xRowRef).Formula = "=N3*10"
    End If
    
    tempWs2.Delete
    
    wsTo2.Range("A2:P2").Copy wsTo3.Range("A1")
    xRow1 = 3
    xRow2 = 2
    While Not wsTo2.Range("J" & xRow1).Value = ""
        Select Case wsTo2.Range("B" & xRow1).Value
        Case "B1", "B2", "B4", "B5", "B7", "B9", "IQ"
            'input, 10%, positive
            If Abs(Abs(wsTo2.Range("J" & xRow1).Value / 10) - Abs(wsTo2.Range("K" & xRow1).Value)) > 1 Or wsTo2.Range("K" & xRow1).Value < 0 Then
                wsTo2.Range("A" & xRow1 & ":P" & xRow1).Copy wsTo3.Range("A" & xRow2)
                xRow2 = xRow2 + 1
            End If
            
        Case "S1"
            'output, 10%, negative
            If Abs(Abs(wsTo2.Range("J" & xRow1).Value / 10) - Abs(wsTo2.Range("K" & xRow1).Value)) > 1 Or wsTo2.Range("K" & xRow1).Value > 0 Then
                wsTo2.Range("A" & xRow1 & ":P" & xRow1).Copy wsTo3.Range("A" & xRow2)
                xRow2 = xRow2 + 1
            End If
            
        'Case "B3", "B6", "B8", "P0"
            'input, 0%, positive
        'Case "S0", "S2", "S3", "S4", "S9"
            'output, 0%, negative
        Case Else
            If Abs(wsTo2.Range("K" & xRow1).Value) > 0 Then
                wsTo2.Range("A" & xRow1 & ":P" & xRow1).Copy wsTo3.Range("A" & xRow2)
                xRow2 = xRow2 + 1
            End If
            
        End Select
        
        xRow1 = xRow1 + 1
    Wend
    
    
    
    
    'need to clean ws sheet
    
    xRow1 = 1
    xRow2 = 0
    xLastRow = Application.WorksheetFunction.Max(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, ws.Cells(ws.Rows.Count, "D").End(xlUp).Row, ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)
    
    Set tempRng1 = ws.Range("A1:G10").Find("CoCd", LookIn:=xlValues, lookat:=xlWhole)
    Set tempRng2 = ws.Range("A1:G10").Find("Comp", LookIn:=xlValues, lookat:=xlWhole)
    Do While xRow1 <= xLastRow
        If Not tempRng1 Is Nothing Then 'CoCd
            Set tempRng1 = ws.Range("A" & xRow1 & ":G" & xRow1 + 10).Find("CoCd")
            xLastRow = tempRng1.Row
            If xLastRow > xRow2 Then
                If xRow2 > 1 Then
                    ws.Range("A" & xRow2 + 1 & ":A" & xLastRow - 2).EntireRow.Delete
                Else
                    ws.Range("A" & xRow2 + 1 & ":A" & xLastRow - 1).EntireRow.Delete
                End If
            End If
        ElseIf Not tempRng2 Is Nothing Then
            Set tempRng2 = ws.Range("A" & xRow1 & ":G" & xRow1 + 10).Find("Comp")
            xLastRow = tempRng2.Row
            If xLastRow > xRow2 Then
                If xRow2 > 1 Then
                    ws.Range("A" & xRow2 + 1 & ":A" & xLastRow - 2).EntireRow.Delete
                Else
                    ws.Range("A" & xRow2 + 1 & ":A" & xLastRow - 1).EntireRow.Delete
                End If
            End If
        Else
            Exit Do
        End If
        
        xRow2 = xRow1
        xLastRow = Application.WorksheetFunction.Max(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, ws.Cells(ws.Rows.Count, "C").End(xlUp).Row, ws.Cells(ws.Rows.Count, "D").End(xlUp).Row, ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)
        Do While ws.Range("A" & xRow2 + 1).Value = ""
            xRow2 = xRow2 + 1
            If xRow2 >= xLastRow Then
                Exit Do
            End If
        Loop
        
        xLastCol = Application.WorksheetFunction.Max(ws.Cells(xRow1, ws.Columns.Count).End(xlToLeft).Column, ws.Cells(xRow1 + 1, ws.Columns.Count).End(xlToLeft).Column)
        'xRow1 = 7
        For xCol = 1 To xLastCol
            If ws.Range(NumberToLetter(xCol) & xRow1).Value = "" And ws.Range(NumberToLetter(xCol) & xRow1 + 1).Value = "" Then
                
                ws.Range(NumberToLetter(xCol) & xRow1 & ":" & NumberToLetter(xCol) & xRow2).Delete shift:=xlToLeft
                xCol = xCol - 1
                xLastCol = xLastCol - 1
                If xLastCol < xCol Then Exit For
            End If
        Next
        
        xRow1 = xRow2 + 1
        xLastRow = Application.WorksheetFunction.Max(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, ws.Cells(ws.Rows.Count, "D").End(xlUp).Row, ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)
    Loop
End Sub






































'**************************************************************************************************************
'SAP Scripts
'**************************************************************************************************************


Sub SAP_Run1()
    percProgress = 0.02
    Call ChangeProgress("Start downloading process: FAGLL03")
    
    Dim xCount1 As Long, xCount2 As Long, xCount3 As Long
    
    Dim xCat As String, xFolder As String
    Dim myFileName As String
    
    setFormat = True
    
    
    SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    
    SesSion.findById("wnd[0]/tbar[0]/okcd").Text = "FAGLL03"
    SesSion.findById("wnd[0]").sendVKey 0
    
    
    For xCount1 = 1 To 5
        If xCount1 = 1 Then
            percProgress = 0.01
            Call ChangeProgress("Start downloading process: FAGLL03 - Input")
            SesSion.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").Text = "20610101"
            xCat = " Input "
            
        ElseIf xCount1 = 2 Then
            percProgress = 0.04
            Call ChangeProgress("Start downloading process: FAGLL03 - Output")
            SesSion.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").Text = "20610102"
            xCat = " Output "
            
        ElseIf xCount1 = 3 Then
            percProgress = 0.07
            Call ChangeProgress("Start downloading process: FAGLL03 - Clearing")
            SesSion.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").Text = "20610000"
            xCat = " Clearing "
            
        ElseIf xCount1 = 4 Then
            percProgress = 0.1
            Call ChangeProgress("Start downloading process: FAGLL03 - Sales")
            SesSion.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").Text = "63000102"
            xCat = " Sales "
            
        ElseIf xCount1 = 5 Then
            percProgress = 0.11
            Call ChangeProgress("Start downloading process: FAGLL03 - 107")
            SesSion.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").Text = "63000107"
            xCat = " 107 "
            
        End If
        
        
        
        If xCount1 <= 3 Then
            For xCount3 = 0 To 9
                'If Not xCount3 = 2 Then
                    xFolder = myPath(xCount3)
                    
                    SesSion.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").Text = "1113" 'set first company code
                    SesSion.findById("wnd[0]/usr/btn%_SD_BUKRS_%_APP_%-VALU_PUSH").press 'set item in company code
                    SesSion.findById("wnd[1]/tbar[0]/btn[16]").press 'clear existing items
                    
                    
                    Select Case xCount3
                    Case 0 'QGC Upstream Holdings Pty Ltd (BGIA)
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1100"
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "1106"
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "1116"
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "1101"
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "1112"
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = "1122"
                        myFileName = "BGIA" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                        xFolder = myPath(0)
                        
                    Case 1 'QGC Pty Ltd    - 2 is JV, not applicable
                        wsRef.Range("A2:A37").Copy
                        SesSion.findById("wnd[1]/tbar[0]/btn[24]").press 'paste
                        Application.CutCopyMode = False
                        myFileName = "QGC" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                        xFolder = myPath(1)
                        
                    Case 2 'QGC Train 2 Tolling Pty Ltd
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5030"
                        myFileName = "5030" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                        xFolder = myPath(3)
                        
                    Case 3 'QGC Train 2 Tolling No. 2 Pty Ltd
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5031"
                        myFileName = "5031" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                        xFolder = myPath(4)
                        
                    Case 4 'QGC Train 1 Tolling Pty Ltd
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5033"
                        myFileName = "5033" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                        xFolder = myPath(5)
                        
                    Case 5 'QCLNG Operating Company Pty Ltd
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1113"
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "5036"
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "5039"
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "1127"
                        myFileName = "QCLNG" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                        xFolder = myPath(6)
                        
                    Case 6 'QGC Train 1 Pty Ltd
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5037"
                        myFileName = "5037" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                        xFolder = myPath(7)
                        
                    Case 7 'QCLNG Train 2 Pty Ltd
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5038"
                        myFileName = "5038" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                        xFolder = myPath(8)
                        
                    Case 8 'QCLNG – QGC / CNOOC T1 Joint Venture
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5045"
                        myFileName = "5045" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                        xFolder = myPath(9)
                        
                    Case 9 'QCLNG – QGC / Tokyo Gas T2 Joint Venture
                        SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5046"
                        myFileName = "5046" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                        xFolder = myPath(10)
                        
                    End Select
                    'Debug.Print xFolder & "       " & myFileName
                    
                    
                    
                    
                    Call cont_SAP_1(xFolder, myFileName)
                    
                'End If
            Next
        Else
            
            For xCount2 = 1 To 3
                
                SesSion.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").Text = "1113" 'set first company code
                SesSion.findById("wnd[0]/usr/btn%_SD_BUKRS_%_APP_%-VALU_PUSH").press 'set item in company code
                SesSion.findById("wnd[1]/tbar[0]/btn[16]").press 'clear existing items
            
                
                Select Case xCount2
                Case 1
                    SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5033"
                    myFileName = "5033" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                    xFolder = myPath(5)
                Case 2
                    SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5030"
                    myFileName = "5030" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                    xFolder = myPath(3)
                Case 3
                    SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5031"
                    myFileName = "5031" & xCat & Format(DateAdd("m", -1, Date), "MMM YYYY")
                    xFolder = myPath(4)
                End Select
                
                Call cont_SAP_1(xFolder, myFileName)
            Next
        End If
        
    Next
    
    SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    
End Sub



Sub cont_SAP_1(curFilePath As String, curFileName As String)
    Dim tempCount As Long
    
    Dim wShell As Object, bWindowFound
    Set wShell = CreateObject("WScript.Shell")
    
    
    SesSion.findById("wnd[1]/tbar[0]/btn[8]").press
                
    SesSion.findById("wnd[0]/usr/radX_AISEL").Select
    SesSion.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "dd.mm.yyyy")
    SesSion.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = Format(DateSerial(Year(Date), Month(Date), 1 - 1), "dd.mm.yyyy")
    
    SesSion.findById("wnd[0]/tbar[1]/btn[8]").press
    
    
    'to check if there's no report to download
    
    
    If Left(SesSion.findById("wnd[0]/sbar").Text, 17) = "No items selected" Then
        failedReport = failedReport & vbLf & " - " & curFileName
        Exit Sub
    End If
    
    
    If setFormat = True Then
        setFormat = False
        
        SesSion.findById("wnd[0]/tbar[1]/btn[32]").press
        SesSion.findById("wnd[1]/usr/btnAPP_FL_ALL").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Company Code"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document currency"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Amount in doc. curr."
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Local Currency"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Amount in local currency"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Local currency 2"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Amount in loc.curr.2"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Local currency 3"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Amt in loc.curr. 3"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document Number"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document Type"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Posting Date"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document Date"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Posting Key"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Tax code"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Profit Center"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
        
        SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
        SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Text"
        SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
        SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
        
        SesSion.findById("wnd[1]/tbar[0]/btn[0]").press
        
    End If
    
    
    
    
    
    
    
    
    
    'download PDF version
    SesSion.findById("wnd[0]/tbar[0]/btn[86]").press
    SesSion.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "SAPWIN"
    SesSion.findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "Microsoft Print to PDF"
    SesSion.findById("wnd[1]/usr/ctxtPRI_PARAMS-PRCOP").Text = "1"
    SesSion.findById("wnd[1]/usr/radRADIO0500_1").SetFocus
    SesSion.findById("wnd[1]/tbar[0]/btn[13]").press
    
    
    wsSettings.Range("D1").Value = curFilePath
    
    
    Application.Wait (Now + TimeValue("0:00:03"))
    
    
    Do
        bWindowFound = wShell.AppActivate("Save Print Output As")
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop Until bWindowFound
    
    
    'bWindowFound = wShell.AppActivate("Save Print Output As")
    
    wShell.SendKeys curFileName & ".pdf"
    Application.Wait (Now + TimeValue("0:00:01"))
    
    'bWindowFound = wShell.AppActivate("Save Print Output As") 'new test
    wShell.SendKeys "{F4}"
    Application.Wait (Now + TimeValue("0:00:02"))
    
    'bWindowFound = wShell.AppActivate("Save Print Output As") 'new test
    wShell.SendKeys "^A"
    Application.Wait (Now + TimeValue("0:00:01"))
    
    'bWindowFound = wShell.AppActivate("Save Print Output As") 'new test
    wsSettings.Range("D1").Copy
    Application.Wait (Now + TimeValue("0:00:01"))
    wShell.SendKeys "^V"
    Application.Wait (Now + TimeValue("0:00:02"))
    
    For tempCount = 1 To 6
        
        wShell.SendKeys "{ENTER}"
        Application.Wait (Now + TimeValue("0:00:01"))
        'Sleep (200)
    Next
    
    
    Application.CutCopyMode = False
    wsSettings.Range("D1").Value = ""
    
    
    
    
    
    
    
    
    
    
    
    
    
    SesSion.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
    
    SesSion.findById("wnd[1]/usr/ctxtDY_PATH").Text = curFilePath
    SesSion.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = curFileName & ".xlsx"
    
    SesSion.findById("wnd[1]/tbar[0]/btn[0]").press
    
    
    'If TestPDF = False Then
    '    GoTo finishDyy
    'End If
    
    
finishDyy:
    
    SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    
End Sub


Sub SAP_Run2()
    Dim xCount1 As Long, curFileName As String, curFilePath As String
    Dim tempCount As Long
    Dim wShell As Object, bWindowFound
    Set wShell = CreateObject("WScript.Shell")
    
    SesSion.findById("wnd[0]/tbar[0]/okcd").Text = "S_ALR_87012357"
    SesSion.findById("wnd[0]").sendVKey 0
    
    SesSion.findById("wnd[0]/usr/txtBR_GJAHR-LOW").Text = Year(DateAdd("m", -1, Date))
    SesSion.findById("wnd[0]/usr/ctxtBR_BUDAT-LOW").Text = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "dd.mm.yyyy")
    SesSion.findById("wnd[0]/usr/ctxtBR_BUDAT-HIGH").Text = Format(DateSerial(Year(Date), Month(Date), 1 - 1), "dd.mm.yyyy")
    
    SesSion.findById("wnd[0]/usr/ctxtEXCDT").Text = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "dd.mm.yyyy")
    
    SesSion.findById("wnd[0]/usr/chkALCUR").Selected = True
    
    SesSion.findById("wnd[0]/usr/btnPUSHB_O4").press
    
    
    For xCount1 = 1 To 6
        
        SesSion.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "5036" 'set first company code
        SesSion.findById("wnd[0]/usr/btn%_BR_BUKRS_%_APP_%-VALU_PUSH").press 'set item in company code
        SesSion.findById("wnd[1]/tbar[0]/btn[16]").press 'clear existing items
    
    
        Select Case xCount1
        Case 1
            percProgress = 0.12
            Call ChangeProgress("Start downloading process: S_ALR_87012357 - 5033")
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5033"
            curFilePath = myPath(5)
            curFileName = "5033 GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            
        Case 2
            percProgress = 0.14
            Call ChangeProgress("Start downloading process: S_ALR_87012357 - 5030")
        
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5030"
            curFilePath = myPath(3)
            curFileName = "5030 GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            
        Case 3
            percProgress = 0.16
            Call ChangeProgress("Start downloading process: S_ALR_87012357 - 5031")
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5031"
            curFilePath = myPath(4)
            curFileName = "5031 GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            
        Case 4
            percProgress = 0.18
            Call ChangeProgress("Start downloading process: S_ALR_87012357 - 5037")
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5037"
            curFilePath = myPath(7)
            curFileName = "5037 GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            
        Case 5
            percProgress = 0.2
            Call ChangeProgress("Start downloading process: S_ALR_87012357 - 5038")
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5038"
            curFilePath = myPath(8)
            curFileName = "5038 GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            
        Case 6
            percProgress = 0.22
            Call ChangeProgress("Start downloading process: S_ALR_87012357 - 5036")
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5036"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "5039"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "1113"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "1127"
            curFilePath = myPath(6)
            'curFileName = "5036 GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFileName = "QCLNG GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            
        End Select
        
        SesSion.findById("wnd[1]/tbar[0]/btn[8]").press
        
        
        If xCount1 = 1 Then
            'set the format for GST OUTPUT and GST INPUT
            For tempCount = 1 To 2
                If tempCount = 1 Then
                    SesSion.findById("wnd[0]/usr/btn%P028143_1000").press
                    SesSion.findById("wnd[0]/tbar[1]/btn[32]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_FL_ALL").press
                    
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "COMPANY CODE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "FISCAL PERIOD"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "DOCUMENT TYPE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "TAX ACCOUNT"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "FISCAL YEAR"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "DOCUMENT NUMBER"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "POSTING DATE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "REFERENCE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "TAX CODE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "TRANSACTION"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "TAX BASE AMOUNT"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "DEDUCT. INPUT TAX"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "GROSS AMOUNT"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "INPUT TAX"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "ALT REFERENCE NUMBER"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "NON-DEDUCTIBLE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "VENDOR"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "G/L ACCOUNT"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[0]/tbar[1]/btn[34]").press
                    
                    
                    SesSion.findById("wnd[1]/usr/ctxtLTDX-VARIANT").Text = "GST_INPUT"
                    SesSion.findById("wnd[1]/usr/txtLTDXT-TEXT").Text = "INPUT TAX LINE ITEMS"
                    
                    
                Else
                    SesSion.findById("wnd[0]/usr/btn%P028127_1000").press
                    SesSion.findById("wnd[0]/tbar[1]/btn[32]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_FL_ALL").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "COMPANY CODE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "FISCAL PERIOD"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "FISCAL YEAR"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "POSTING DATE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "DOCUMENT DATE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "DOCUMENT NUMBER"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "REFERENCE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "TAX CODE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "TRANSACTION"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "TAX BASE AMOUNT"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "OUTPUT TAX PAYABLE"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "GROSS AMOUNT"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "OUTPUT TAX"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "ALT REFERENCE NUMBER"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/usr/btnB_SEARCH").press
                    SesSion.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "NOT TO BE PAID OVER"
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[1]/usr/btnAPP_WL_SING").press
                    
                    SesSion.findById("wnd[1]/tbar[0]/btn[0]").press
                    SesSion.findById("wnd[0]/tbar[1]/btn[34]").press
                    
                    SesSion.findById("wnd[1]/usr/ctxtLTDX-VARIANT").Text = "GST_OUTPUT"
                    SesSion.findById("wnd[1]/usr/txtLTDXT-TEXT").Text = "OUTPUT TAX LINE ITEMS"
                End If
                                
                SesSion.findById("wnd[1]/tbar[0]/btn[0]").press
                
                'warning window: activewindow.name = wnd[2]
                'main SAP page: activewindow.name = wnd[0]
                If SesSion.ActiveWindow.Name = "wnd[2]" Then
                    SesSion.findById("wnd[2]/tbar[0]/btn[0]").press
                End If
                
                SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
            Next
        End If
        
        
        SesSion.findById("wnd[0]/usr/ctxtPAR_VAR1").Text = "GST_OUTPUT"
        SesSion.findById("wnd[0]/usr/ctxtPAR_VAR3").Text = "GST_INPUT"
        
        SesSion.findById("wnd[0]/tbar[1]/btn[8]").press
        
        
        
        
        
        
        
        
        SesSion.findById("wnd[0]/tbar[0]/btn[86]").press
        SesSion.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "SAPWIN"
        SesSion.findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "Microsoft Print to PDF"
        SesSion.findById("wnd[1]/usr/ctxtPRI_PARAMS-PRCOP").Text = "1"
        SesSion.findById("wnd[1]/usr/radRADIO0500_1").SetFocus
        SesSion.findById("wnd[1]/tbar[0]/btn[13]").press
        
        
        
        
        wsSettings.Range("D1").Value = curFilePath
        
        Application.Wait (Now + TimeValue("0:00:02"))
        Do
            bWindowFound = wShell.AppActivate("Save Print Output As")
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop Until bWindowFound
        
        bWindowFound = wShell.AppActivate("Save Print Output As")
        
        wShell.SendKeys curFileName & ".pdf"
        Application.Wait (Now + TimeValue("0:00:01"))
        
        'bWindowFound = wShell.AppActivate("Save Print Output As") 'new test
        wShell.SendKeys "{F4}"
        Application.Wait (Now + TimeValue("0:00:02"))
        
        'bWindowFound = wShell.AppActivate("Save Print Output As") 'new test
        wShell.SendKeys "^A"
        Application.Wait (Now + TimeValue("0:00:01"))
        
        'bWindowFound = wShell.AppActivate("Save Print Output As") 'new test
        wsSettings.Range("D1").Copy
        Application.Wait (Now + TimeValue("0:00:01"))
        wShell.SendKeys "^V"
        Application.Wait (Now + TimeValue("0:00:02"))
        
        For tempCount = 1 To 6
            wShell.SendKeys "{ENTER}"
            Application.Wait (Now + TimeValue("0:00:01"))
            'Sleep (200)
        Next
        
        
        Application.CutCopyMode = False
        wsSettings.Range("D1").Value = ""
        
        
        
        
        SesSion.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
        SesSion.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        SesSion.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
        SesSion.findById("wnd[1]/tbar[0]/btn[0]").press
        
        
        SesSion.findById("wnd[1]/usr/ctxtDY_PATH").Text = curFilePath
        SesSion.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = curFileName & ".xls"
        SesSion.findById("wnd[1]/tbar[0]/btn[0]").press
        
        
        
        'If TestPDF = False Then
        '    GoTo finishDyy
        'End If
        
        
        
        
        
        
finishDyy:
        
        
        SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
        
    Next
    
    
    SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    
    
End Sub



Sub SAP_Run3()
    Dim xCount1 As Long, curFileName As String, curFilePath As String
    
    Dim tempCount As Long
    Dim wShell As Object, bWindowFound
    Set wShell = CreateObject("WScript.Shell")
    
    
    
    For xCount1 = 1 To 6
        SesSion.findById("wnd[0]/tbar[0]/okcd").Text = "ZBGFI_GJVA"
        SesSion.findById("wnd[0]").sendVKey 0
        
        
        SesSion.findById("wnd[0]/usr/txtBR_GJAHR-LOW").Text = Year(DateAdd("m", -1, Date))
        SesSion.findById("wnd[0]/usr/ctxtBR_BUDAT-LOW").Text = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "dd.mm.yyyy")
        SesSion.findById("wnd[0]/usr/ctxtBR_BUDAT-HIGH").Text = Format(DateSerial(Year(Date), Month(Date), 1 - 1), "dd.mm.yyyy")
    
        
        SesSion.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").Text = "1100" 'set first company code
        SesSion.findById("wnd[0]/usr/btn%_BR_BUKRS_%_APP_%-VALU_PUSH").press 'set item in company code
        SesSion.findById("wnd[1]/tbar[0]/btn[16]").press 'clear existing items
    
        Select Case xCount1
        
        
        Case 1
            percProgress = 0.24
            Call ChangeProgress("Start downloading process: ZBGFI_GJVA - 1100")
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1100"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "1101"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "1106"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "1112"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "1116"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = "1122"
            
            curFileName = "BGIA GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFilePath = myPath(0)
            
        Case 2
            percProgress = 0.26
            Call ChangeProgress("Start downloading process: ZBGFI_GJVA - 5012")
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5012"
            'surat
            curFileName = "JV769 GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFilePath = myPath(2)
            
        Case 3
            percProgress = 0.28
            Call ChangeProgress("Start downloading process: ZBGFI_GJVA - 5015")
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5015"
            'sunshine
            curFileName = "JV685 TARDRUM GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFilePath = myPath(2)
        
        Case 4
            percProgress = 0.3
            Call ChangeProgress("Start downloading process: ZBGFI_GJVA - QGC")
            
            wsRef.Range("A2:A37").Copy
            SesSion.findById("wnd[1]/tbar[0]/btn[24]").press 'paste
            Application.CutCopyMode = False
            
            curFileName = "QGC GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFilePath = myPath(1)
            
        Case 5
            percProgress = 0.32
            Call ChangeProgress("Start downloading process: ZBGFI_GJVA - QGC JV")
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5000"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "5004"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "5012"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "1106"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "1112"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = "5002"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").Text = "1122"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = "5014"
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = 1
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").Text = "5018"
            'SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,8]").Text = "5015"
            
            curFileName = "QGC JV GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFilePath = myPath(1)
            
        Case 6
            percProgress = 0.3
            Call ChangeProgress("Start downloading process: ZBGFI_GJVA - JV171")
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "5000"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "5009"
            
            curFileName = "JV171 GST Report SAP Download " & Format(DateAdd("m", -1, Date), "MMM YYYY")
            curFilePath = myPath(1)
            
        End Select
        
        
        SesSion.findById("wnd[1]/tbar[0]/btn[8]").press
        
        SesSion.findById("wnd[0]/usr/btn%_SEL_VNAM_%_APP_%-VALU_PUSH").press
        SesSion.findById("wnd[1]/tbar[0]/btn[16]").press 'clear existing items
        
        Select Case xCount1
        Case 1
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "O22000"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = "O22001"
        Case 2
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "O12008"
        Case 3
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "O15000"
        Case 4
            
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            'SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "O00010"
            wsRef.Range("C2:C72").Copy
            SesSion.findById("wnd[1]/tbar[0]/btn[24]").press 'paste
            Application.CutCopyMode = False
            
        Case 5
            wsRef.Range("E2:E70").Copy
            SesSion.findById("wnd[1]/tbar[0]/btn[24]").press 'paste
            Application.CutCopyMode = False
        Case 6
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "O09000"
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "O00039"
        End Select
        
        
        SesSion.findById("wnd[1]/tbar[0]/btn[8]").press
    
        Select Case xCount1
        Case Is >= 4
            SesSion.findById("wnd[0]/tbar[1]/btn[16]").press
            
            SesSion.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN001-LOW").Text = ""
            SesSion.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
            SesSion.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "ZI"
            SesSion.findById("wnd[1]/tbar[0]/btn[8]").press
            
        End Select
        
        
        SesSion.findById("wnd[0]/tbar[1]/btn[8]").press
        
        If Not SesSion.findById("wnd[1]", False) Is Nothing Then
            SesSion.findById("wnd[1]").HardCopy curFilePath & curFileName & ".jpg", 1
            SesSion.findById("wnd[1]/tbar[0]/btn[0]").press
            GoTo nothingFound
        End If

        
        
        
        
        'If TestPDF = False Then
        '    GoTo nothingFound
        'End If
        
        SesSion.findById("wnd[0]/tbar[0]/btn[86]").press
        SesSion.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "SAPWIN"
        SesSion.findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "Microsoft Print to PDF"
        SesSion.findById("wnd[1]/usr/ctxtPRI_PARAMS-PRCOP").Text = "1"
        SesSion.findById("wnd[1]/usr/radRADIO0500_1").SetFocus
        SesSion.findById("wnd[1]/tbar[0]/btn[13]").press
        
        
        wsSettings.Range("D1").Value = curFilePath
    
        Application.Wait (Now + TimeValue("0:00:03"))
        Do
            bWindowFound = wShell.AppActivate("Save Print Output As")
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop Until bWindowFound
        
        bWindowFound = wShell.AppActivate("Save Print Output As")
        
        wShell.SendKeys curFileName & ".pdf"
        Application.Wait (Now + TimeValue("0:00:01"))
        
        'bWindowFound = wShell.AppActivate("Save Print Output As") 'new test
        wShell.SendKeys "{F4}"
        Application.Wait (Now + TimeValue("0:00:02"))
        
        'bWindowFound = wShell.AppActivate("Save Print Output As") 'new test
        wShell.SendKeys "^A"
        Application.Wait (Now + TimeValue("0:00:01"))
        
        'bWindowFound = wShell.AppActivate("Save Print Output As") 'new test
        wsSettings.Range("D1").Copy
        Application.Wait (Now + TimeValue("0:00:01"))
        wShell.SendKeys "^V"
        Application.Wait (Now + TimeValue("0:00:02"))
        
        For tempCount = 1 To 6
            wShell.SendKeys "{ENTER}"
            Application.Wait (Now + TimeValue("0:00:01"))
            'Sleep (200)
        Next
        
        
        Application.CutCopyMode = False
        wsSettings.Range("D1").Value = ""
        
        
        
        
        
        
        SesSion.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
        SesSion.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        SesSion.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
        SesSion.findById("wnd[1]/tbar[0]/btn[0]").press
        
        SesSion.findById("wnd[1]/usr/ctxtDY_PATH").Text = curFilePath
        SesSion.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = curFileName & ".xls"

        SesSion.findById("wnd[1]/tbar[0]/btn[0]").press
        
        
        
        SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
        
        
        
        
        
nothingFound:
        SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
        SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    Next
    
    
    SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    SesSion.findById("wnd[0]/tbar[0]/btn[3]").press
    
    
End Sub




    


        



















'**************************************************************************************************************
'Email Creation
'**************************************************************************************************************

Sub emailCash()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strBody As String
    Dim sTo As String
    Dim sCC As String
    Dim myDateVal As String
    Dim xCount As Long, xCount2 As Long
    
    percProgress = 0.35
    Call ChangeProgress("Cash Calls - preparing email")
    
    
    sTo = "J.Jamaluddin@shell.com; SBSC-CMKL-GSS-TH-AU@shell.com; R.Paskaradass@shell.com"
    sCC = "Amir.Jamal@shell.com; Kay.Pfingst@shell.com; Jonathan.Soosay@shell.com; Irina.Ilushin@shell.com; Pey-Shy.Lee@shell.com; GXQGCJVFinance@shell.com; Kamini.Raja@shell.com; Zi-Yang.Puah@shell.com; M-S.Hanapi@shell.com; Norzeiny.M-Zain@shell.com"
    
    ' HTML before rows
    
    
    strBody = "<html><body>Hi all,<p>Please see details below for " & Format(DateAdd("m", -1, Date), "MMMM YYYY") & " BAS estimates. <br><br><br>"
    strBody = strBody & "<head><style>table, th, td {border: 1px solid black;}" & _
        "<table style=""width:40%""><tr>" & _
        "<b><th bgcolor=""#ff8500"">Coy</th></b>" & _
        "<b><th bgcolor=""#ff8500"">Entity</th></b>" & _
        "<b><th bgcolor=""#ff8500"">(Pay)/ Refund</th></b>" & _
        "<b><th bgcolor=""#ff8500"">Estimated Refund/ Payment Date</th></tr></b>"
    
    
    ' iterate collection
    For xCount = 27 To 38
        strBody = strBody & "<tr>"
        strBody = strBody & "<td ""col width=5%"" align=""center"">" & wsC_Forcasting.Range("B" & xCount).Value & "</td>"
        strBody = strBody & "<td ""col width=15%"" align=""center"">" & wsC_Forcasting.Range("C" & xCount).Value & "</td>"
        
        If wsC_Forcasting.Range("D" & xCount).Value < 0 Then
            strBody = strBody & "<td ""col width=10%"" align=""center""><font color=""red"">" & Format(wsC_Forcasting.Range("D" & xCount).Value, "#,##0;(#,##0)") & "</td>"
        Else
            strBody = strBody & "<td ""col width=10%"" align=""center"">" & Format(wsC_Forcasting.Range("D" & xCount).Value, "#,##0;(#,##0)") & "</td>"
        End If
        
        
        strBody = strBody & "<td ""col width=10%"" align=""center"">" & Format(wsC_Forcasting.Range("E" & xCount).Value, "dd-mmm-yyyy") & "</td>"
        strBody = strBody & "</tr>"
    Next
    
    strBody = strBody & "<tr>"
    strBody = strBody & "<td ""col width=5%"" align=""center""><b>" & wsC_Forcasting.Range("B39").Value & "</b></td>"
    strBody = strBody & "<td ""col width=15%"" align=""center""><b>" & wsC_Forcasting.Range("C39").Value & "</b></td>"
    
    If wsC_Forcasting.Range("D39").Value < 0 Then
        strBody = strBody & "<td ""col width=10%"" align=""center""><font color=""red""><b>" & Format(wsC_Forcasting.Range("D39").Value, "#,##0;(#,##0)") & "</b></td>"
    Else
        strBody = strBody & "<td ""col width=10%"" align=""center""><b>" & Format(wsC_Forcasting.Range("D39").Value, "#,##0;(#,##0)") & "</b></td>"
    End If
    
    strBody = strBody & "<td ""col width=10%""  align=""center""><b>" & Format(wsC_Forcasting.Range("E39").Value, "dd-mmm-yyyy") & "</b></td>"
    strBody = strBody & "</tr>"
    
    
    strBody = strBody & "</table><br><br>Regards,<br>" & Left(Application.UserName, InStr(1, Application.UserName, " ", vbTextCompare) - 1) & "</body></html>"
    'strbody = strbody & "</table><br><br>Regards,<br>" & Application.UserName & "</body></html>"
    
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = sTo
        .cc = sCC
        .Subject = Format(DateAdd("m", -1, Date), "MMMM YYYY") & " - Cash Flow Estimates"
        .HTMLBody = strBody 'strbody & OutMail.HTMLBody
        .Display
    End With
    
End Sub
    






