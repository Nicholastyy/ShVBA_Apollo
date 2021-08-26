VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf4_Clearing 
   Caption         =   "Australia SAPL & SEHAL Automation: Clearing"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14715
   OleObjectBlob   =   "NAS_UF3_Clearing.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "uf4_Clearing"
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

Private StartTime       As Double, SecondsElapsed As Double, MinutesElapsed As String
Private inputError      As String

Private wsSettings As Worksheet
Private wsMain As Worksheet

Private myLoc As String, myAlias As String

Private firstDownload As Boolean





'**************************************************************************************************************
'Main Routine
'**************************************************************************************************************
Private Sub BtnRun_Click()
    StartTime = Timer
    inputError = ""
    
    checkFields
    
    If inputError <> "" Then
        MsgBox ("Please ensure the following issues are corrected before proceeding:" & inputError)
        Exit Sub
    End If
    
    checkOpenSAP
    If openP16 = False Then
        MsgBox ("Please ensure you have SAP Blueprint open.")
        Exit Sub
    End If
    
    RunPauseAll
    
    makeFolder
    
    firstDownload = False
    downloadStuff
    
    closeNotThis
    'Call CloseAll(myLoc)
    
    cleanReports
    
    
    
    
    
    
    
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MinutesElapsed = VBA.Format((Timer - StartTime) / 86400, "hh:mm:ss")
    RunActivateAll
    
    MsgBox ("Automation process completed for Australia SAPL & SEHAL Cleansing process." & vbLf & "Time taken: " & MinutesElapsed)
    Unload Me
    
End Sub






'**************************************************************************************************************
'Requirement Runs
'**************************************************************************************************************


Private Sub UserForm_Initialize()
    Set wsSettings = ThisWorkbook.Worksheets("Settings")
    Set wsMain = ThisWorkbook.Worksheets("Main")
    
    If Not VBA.Environ("username") = wsSettings.Range("B1").Value Then
        wsSettings.Range("B:B").EntireColumn.Clear
        wsSettings.Range("B1").Value = VBA.Environ("username")
    Else
        Me.tbSaveLocation.Value = wsSettings.Range("B5").Value
        Me.tbClearing.Value = wsSettings.Range("B8").Value
        Me.tbAlias.Value = wsSettings.Range("B7").Value
    End If
    
    Me.tbInput.Value = wsMain.Range("P15").Value
    Me.tbOutput.Value = wsMain.Range("P17").Value
    
End Sub


Private Sub btnClearing_Click()
    Me.tbClearing.Value = SearchFileLocation
    wsSettings.Range("B8").Value = Me.tbClearing.Value
End Sub


Private Sub btnClear_Click()
    Me.tbClearing.Value = ""
    wsSettings.Range("B8").Value = ""
End Sub


Private Sub BtnCancel_Click()
    Unload Me
End Sub


Private Sub btnSaveLocation_Click()
    Me.tbSaveLocation.Value = SearchFolderLocation
    wsSettings.Range("B5").Value = Me.tbSaveLocation.Value
End Sub

Private Sub tbAlias_Change()
    wsSettings.Range("B7").Value = Me.tbAlias.Value
End Sub


'**************************************************************************************************************
'Processing Runs
'**************************************************************************************************************

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
    
    
    If Me.tbAlias.Value = "" Then
        inputError = inputError & vbLf & " - Alias (1)"
    ElseIf Len(Me.tbAlias.Value) <> 6 Then
        inputError = inputError & vbLf & " - Alias (2)"
    Else
        myAlias = UCase(Me.tbAlias.Value)
    End If
    
    If Me.tbInput.Value = "" Or Me.tbOutput.Value = "" Then
        If MsgBox("Input GL and/or Output GL value is blank. Proceed as is?", vbYesNo) = vbNo Then
            inputError = inputError & vbLf & " - Fleetplus Input and Output GL (1)"
        End If
    ElseIf Not IsNumeric(Me.tbInput.Value) Or Not IsNumeric(Me.tbOutput.Value) Then
        inputError = inputError & vbLf & " - Fleetplus Input and Output GL (2)"
    End If
    
    If Me.cb_01_27.Value = False And Me.cb_01_28.Value = False And Me.cb_01_30.Value = False And Me.cb_01_31.Value = False And Me.cb_01_37.Value = False Then
        If Me.cb_02_28.Value = False And Me.cb_02_31.Value = False Then
            If Me.cb_10_28.Value = False And Me.cb_10_31.Value = False Then
                If Me.cb_11_28.Value = False And Me.cb_11_31.Value = False Then
                    inputError = inputError & vbLf & " - No checkbox items selected"
                End If
            End If
        End If
    End If
    
    
    
    
    
End Sub

Sub makeFolder()
    Dim xCount As Long
    
    myLoc = Me.tbSaveLocation.Value & "FBB1 Journals\"
    xCount = 0
    
    While Dir(myLoc, vbDirectory) <> ""
        xCount = xCount + 1
        myLoc = Me.tbSaveLocation.Value & "FBB1 Journals (" & xCount & ")\"
    Wend
    
    MkDir myLoc
    
End Sub


Sub downloadStuff()
    Dim xCount As Long
    Dim myGL As String, myCoCd As String
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "FBL3N"
    sessBP.findById("wnd[0]").sendVKey 0
    
    If Me.cb_01_27.Value = True Then Call SAP_FBL3N("AU01", "A2685007")
    If Me.cb_01_28.Value = True Then Call SAP_FBL3N("AU01", "A2680008")
    If Me.cb_01_30.Value = True Then Call SAP_FBL3N("AU01", "A3499000")
    If Me.cb_01_31.Value = True Then Call SAP_FBL3N("AU01", "A3620001")
    If Me.cb_01_37.Value = True Then Call SAP_FBL3N("AU01", "A3780007")
    
    If Me.cb_02_28.Value = True Then Call SAP_FBL3N("AU02", "A2680008")
    If Me.cb_02_31.Value = True Then Call SAP_FBL3N("AU02", "A3620001")
    
    If Me.cb_10_28.Value = True Then Call SAP_FBL3N("AU10", "A2680008")
    If Me.cb_10_31.Value = True Then Call SAP_FBL3N("AU10", "A3620001")
    
    If Me.cb_11_28.Value = True Then Call SAP_FBL3N("AU11", "A2680008")
    If Me.cb_11_31.Value = True Then Call SAP_FBL3N("AU11", "A3620001")
    
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
    Exit Sub
    
    
    For xCount = 1 To 11
        Select Case xCount
        Case 1, 2, 9 To 11
            myCoCd = "AU01"
        Case 3, 4
            myCoCd = "AU11"
        Case 5, 6
            myCoCd = "AU10"
        Case 7, 8
            myCoCd = "AU02"
        End Select
        
        Select Case xCount
        Case 1
            myGL = "A3499000"
        Case 2
            myGL = "A2685007"
        Case 3, 6, 8, 10
            myGL = "A2680008"
        Case 4, 5, 7, 9
            myGL = "A3620001"
        Case 11
            myGL = "A3780007"
        End Select
        
        If xCount = 1 Then
            'Call SAP_FBL3N(myGL, myCoCd, False)
        Else
            'Call SAP_FBL3N(myGL, myCoCd, True)
        End If
        myGL = ""
        myCoCd = ""
    Next
    
    
End Sub



Sub cleanReports()
    Dim xCount As Long
    Dim myGL As String, myCoCd As String
    Dim ws As Worksheet, thisFileName As String
    Dim doItem As Boolean
    
    Dim myTabName As String, myFileName As String, altFileName As String
    Dim myPath As Long
    
    Dim mainWB As Workbook
    Dim mainWS As Worksheet, tempWS As Worksheet
    Dim curWb As Workbook, curWs As Worksheet
    
    
    Dim xRow As Long, xlastrow As Long
    Dim rowDIE As Long, docRow As Long, nextRowDie As Long
    
    Dim dieAmountPM As Double, dieAmountCM As Double, InOutVal As Double
    Dim F03Range As Range
    
    
    
    If Me.tbClearing.Value = "" Then
        Set mainWB = Workbooks.Add
        Set mainWS = mainWB.Worksheets(1)
        
        With mainWS
            .Name = "F-03"
            .Range("A1").Value = "AU01"
            .Range("A2").Value = "A3499000"
            
            .Range("D1").Value = "AU01"
            .Range("D2").Value = "A2685007"
            
            .Range("G1").Value = "AU11"
            .Range("G2").Value = "A2680008"
            
            .Range("J1").Value = "AU11"
            .Range("J2").Value = "A3620001"
            
            .Range("M1").Value = "AU10"
            .Range("M2").Value = "A3620001"
            
            .Range("P1").Value = "AU10"
            .Range("P2").Value = "A2680008"
            
            .Range("S1").Value = "AU02"
            .Range("S2").Value = "A3620001"
            
            .Range("V1").Value = "AU02"
            .Range("V2").Value = "A2680008"
            
            .Range("Y1").Value = "AU01"
            .Range("Y2").Value = "A3620001"
            
            .Range("AB1").Value = "AU01"
            .Range("AB2").Value = "A2680008"
            
            .Range("AE1").Value = "AU01"
            .Range("AE2").Value = "A3780007"
        End With
        
    Else
        Set mainWB = Workbooks.Open(Me.tbClearing.Value, ReadOnly:=True)
        For Each ws In mainWB.Worksheets
            If ws.Name = "F-03" Then
                Set mainWS = ws
            End If
        Next
        
        If mainWS Is Nothing Then
            Set mainWS = mainWB.Worksheets.Add(before:=mainWB.Worksheets(1))
            With mainWS
                .Name = "F-03"
                .Range("A1").Value = "AU01"
                .Range("A2").Value = "A3499000"
                
                .Range("D1").Value = "AU01"
                .Range("D2").Value = "A2685007"
                
                .Range("G1").Value = "AU11"
                .Range("G2").Value = "A2680008"
                
                .Range("J1").Value = "AU11"
                .Range("J2").Value = "A3620001"
                
                .Range("M1").Value = "AU10"
                .Range("M2").Value = "A3620001"
                
                .Range("P1").Value = "AU10"
                .Range("P2").Value = "A2680008"
                
                .Range("S1").Value = "AU02"
                .Range("S2").Value = "A3620001"
                
                .Range("V1").Value = "AU02"
                .Range("V2").Value = "A2680008"
                
                .Range("Y1").Value = "AU01"
                .Range("Y2").Value = "A3620001"
                
                .Range("AB1").Value = "AU01"
                .Range("AB2").Value = "A2680008"
                
                .Range("AE1").Value = "AU01"
                .Range("AE2").Value = "A3780007"
            End With
        End If
    End If
        
    xCount = 0
    thisFileName = "AU_" & Format(DateAdd("M", -1, Date), "MMYYYY") & "_BAS_GST_Upstream Clearing.xlsx"
    While Dir(myLoc & thisFileName) <> ""
        xCount = xCount + 1
        thisFileName = "AU_" & Format(DateAdd("M", -1, Date), "MMYYYY") & "_BAS_GST_Upstream Clearing (" & xCount & ").xlsx"
    Wend
    
    mainWB.SaveAs myLoc & "AU_" & Format(DateAdd("M", -1, Date), "MMYYYY") & "_BAS_GST_Upstream Clearing.xlsx"
    
    
    For xCount = 1 To 11
        doItem = False
        Select Case xCount
        Case 1
            If Me.cb_01_27.Value = True Then 'Call SAP_FBL3N("AU01", "A2685007")
                doItem = True
                myCoCd = "AU01"
                myGL = "A2685007"
                myTabName = "A2685007 AU01"
                Set F03Range = mainWS.Range("E1")
                F03Range.EntireColumn.Clear
            End If
            
        Case 2
            If Me.cb_01_28.Value = True Then 'Call SAP_FBL3N("AU01", "A2680008")
                doItem = True
                myCoCd = "AU01"
                myGL = "A2680008"
                myTabName = "AU01_A2680008 Input Tax"
                Set F03Range = mainWS.Range("AC1")
                F03Range.EntireColumn.Clear
            End If
            
        Case 3
            If Me.cb_01_30.Value = True Then 'Call SAP_FBL3N("AU01", "A3499000")
                doItem = True
                myCoCd = "AU01"
                myGL = "A3499000"
                myTabName = "A3499000 AU01"
                Set F03Range = mainWS.Range("B1")
                F03Range.EntireColumn.Clear
            End If
            
        Case 4
            If Me.cb_01_31.Value = True Then 'Call SAP_FBL3N("AU01", "A3620001")
                doItem = True
                myCoCd = "AU01"
                myGL = "A3620001"
                myTabName = "AU01_A3620001_Output Tax Rcvble"
                Set F03Range = mainWS.Range("Z1")
                F03Range.EntireColumn.Clear
            End If
            
        Case 5
            If Me.cb_01_37.Value = True Then 'Call SAP_FBL3N("AU01", "A3780007")
                doItem = True
                myCoCd = "AU01"
                myGL = "A3780007"
                myTabName = "AU01_A3780007_FBT GL"
                Set F03Range = mainWS.Range("F1")
                F03Range.EntireColumn.Clear
            End If
            
        Case 6
            If Me.cb_02_28.Value = True Then 'Call SAP_FBL3N("AU02", "A2680008")
                doItem = True
                myCoCd = "AU02"
                myGL = "A2680008"
                myTabName = "AU02_A2680008_Input Tax Rcvble"
                Set F03Range = mainWS.Range("W1")
                F03Range.EntireColumn.Clear
            End If
            
        Case 7
            If Me.cb_02_31.Value = True Then 'Call SAP_FBL3N("AU02", "A3620001")
                doItem = True
                myCoCd = "AU02"
                myGL = "A3620001"
                myTabName = "AU02_A3620001_Output Tax Pyble"
                Set F03Range = mainWS.Range("T1")
                F03Range.EntireColumn.Clear
            End If
            
        Case 8
            If Me.cb_10_28.Value = True Then 'Call SAP_FBL3N("AU10", "A2680008")
                doItem = True
                myCoCd = "AU10"
                myGL = "A2680008"
                myTabName = "AU10_A2680008_Input Tax"
                Set F03Range = mainWS.Range("Q1")
                F03Range.EntireColumn.Clear
            End If
            
        Case 9
            If Me.cb_10_31.Value = True Then 'Call SAP_FBL3N("AU10", "A3620001")
                doItem = True
                myCoCd = "AU10"
                myGL = "A3620001"
                myTabName = "AU10_A3620001_Output Tax"
                Set F03Range = mainWS.Range("N1")
                F03Range.EntireColumn.Clear
            End If
            
        Case 10
            If Me.cb_11_28.Value = True Then 'Call SAP_FBL3N("AU11", "A2680008")
                doItem = True
                myCoCd = "AU11"
                myGL = "A2680008"
                myTabName = "AU11_A2680008_Input Tax Rcvble"
                Set F03Range = mainWS.Range("H1")
                F03Range.EntireColumn.Clear
            End If
            
        Case 11
            If Me.cb_11_31.Value = True Then 'Call SAP_FBL3N("AU11", "A3620001")
                doItem = True
                myCoCd = "AU11"
                myGL = "A3620001"
                myTabName = "AU11_A3620001_Output Tax"
                Set F03Range = mainWS.Range("K1")
                F03Range.EntireColumn.Clear
            End If
            
        End Select
            
        If doItem = True Then
            myFileName = myGL & " " & myCoCd & " " & Format(DateAdd("m", -1, Date), "MMM YYYY") & ".xlsx"
            altFileName = myGL & " " & myCoCd & " - No FBL3N Report.jpg"
            Set tempWS = Nothing
            
            For Each ws In mainWB.Worksheets
                If ws.Name = myTabName Then
                    Set tempWS = ws
                    tempWS.Cells.Delete
                End If
            Next
            
            If tempWS Is Nothing Then
                Set tempWS = mainWB.Worksheets.Add(after:=mainWB.Worksheets(mainWB.Worksheets.Count))
                tempWS.Name = myTabName
            End If
            
            
            
            If Dir(myLoc & altFileName) <> "" Then
                tempWS.Pictures.Insert(myLoc & altFileName).Select
                Selection.ShapeRange.IncrementLeft 60
                Selection.ShapeRange.IncrementTop 20
                Selection.ShapeRange.ScaleWidth 1.6, msoFalse, msoScaleFromTopLeft
                Selection.ShapeRange.ScaleHeight 1.6, msoFalse, msoScaleFromTopLeft
                
            ElseIf Dir(myLoc & myFileName) <> "" Then
                nextRowDie = 0
                dieAmountCM = 0
                dieAmountPM = 0
                InOutVal = 0
                
                Set curWb = Workbooks.Open(myLoc & myFileName, ReadOnly:=True)
                Set curWs = curWb.Worksheets(1)
                
                xlastrow = curWs.Cells(curWs.Rows.Count, "I").End(xlUp).Row
                
                curWs.Range("S1").Value = "DONE"
                curWs.Range("S1").Interior.Color = RGB(0, 255, 0)
                
                If curWs.Shapes.Count >= 1 Then
                    curWs.Shapes.SelectAll
                    ActiveSheet.Shapes.SelectAll
                    Selection.Delete
                End If
                
                While curWs.Range("A2").Rows.OutlineLevel > 1
                    curWs.Range("A1:A" & xlastrow).Rows.Ungroup
                Wend
                
                While Not curWs.Range("A" & xlastrow).Interior.Color = RGB(255, 255, 255)
                    While curWs.Range("A" & xlastrow).Interior.Color = RGB(255, 255, 153)
                        curWs.Range("A" & xlastrow).EntireRow.Delete
                    Wend
                    xlastrow = xlastrow - 1
                Wend
                
                curWs.Cells.Copy tempWS.Range("A1")
                curWb.Close False
                Set curWb = Nothing
                Set curWs = Nothing
                
                xlastrow = tempWS.Cells(tempWS.Rows.Count, "A").End(xlUp).Row
                rowDIE = xlastrow + 5
                docRow = xlastrow + 7
                
                'tempWS.Range("E" & rowDie).Value = "DIE Diff"
                
                tempWS.Range("E" & docRow).Font.Bold = True
                tempWS.Range("E" & docRow).Font.Size = 14
                tempWS.Range("E" & docRow).Font.Color = RGB(0, 0, 128)
                
                'tempWS.Range("F" & rowDie).Formula = "=IF(I" & rowDie & "<0,50,40)"
                
                For xRow = 2 To xlastrow
                    If tempWS.Range("O" & xRow).Value = Format(DateAdd("m", -1, Date), "YYYY/MM") Then
                        dieAmountPM = dieAmountPM + tempWS.Range("I" & xRow).Value
                        
                        F03Range.Offset(nextRowDie, 0).Value = tempWS.Range("C" & xRow).Value
                        nextRowDie = nextRowDie + 1
                    
                        tempWS.Range("S" & xRow).Value = "Clear"
                        
                    ElseIf tempWS.Range("O" & xRow).Value = Format(Date, "YYYY/MM") Then
                        
                        If (tempWS.Range("R" & xRow).Value = myAlias And tempWS.Range("D" & xRow).Value = "SA" And (Left(tempWS.Range("F" & xRow).Value, 10) = "SETTLEMENT" Or Left(tempWS.Range("F" & xRow).Value, 4) = "SETL")) Or tempWS.Range("D" & xRow).Value = "DZ" Then  'Or tempWS.Range("D" & xRow).Value = "KR" Then
                            
                            dieAmountCM = dieAmountCM + tempWS.Range("I" & xRow).Value
                            
                            F03Range.Offset(nextRowDie, 0).Value = tempWS.Range("C" & xRow).Value
                            nextRowDie = nextRowDie + 1
                            
                            tempWS.Range("S" & xRow).Value = "Clear"
                        End If
                    End If
                Next
                
                
                
                If myCoCd = "AU01" And myGL = "A2680008" Then
                    tempWS.Range("E" & rowDIE - 1).Value = Format(DateAdd("m", -1, Date), "mmmm yyyy") & " Fleetplus"
                    tempWS.Range("F" & rowDIE - 1).Value = 50
                    tempWS.Range("G" & rowDIE - 1).Value = "A3165001"
                    tempWS.Range("H" & rowDIE - 1).Value = "120378"
                    tempWS.Range("I" & rowDIE - 1).Value = Me.tbInput.Value * -1
                    'tempWS.Range("I" & rowDie).Formula = "=I" & rowDie - 1
                    InOutVal = Me.tbInput.Value
                    
                ElseIf myCoCd = "AU01" And myGL = "A3620001" Then
                    tempWS.Range("E" & rowDIE - 1).Value = Format(DateAdd("m", -1, Date), "mmmm yyyy") & " Fleetplus"
                    tempWS.Range("F" & rowDIE - 1).Value = 40
                    tempWS.Range("G" & rowDIE - 1).Value = "A3165001"
                    tempWS.Range("H" & rowDIE - 1).Value = "120378"
                    tempWS.Range("I" & rowDIE - 1).Value = Me.tbOutput.Value * -1
                    'tempWS.Range("I" & rowDie).Formula = "=I" & rowDie - 1
                    InOutVal = Me.tbOutput.Value
                    
                End If
                    
                If Abs(dieAmountPM + dieAmountCM + InOutVal) < 1 Then
                    If Not Round(dieAmountPM + dieAmountCM + InOutVal, 2) = 0 Then
                        tempWS.Range("D" & rowDIE).Value = "DIE Diff"
                        tempWS.Range("E" & rowDIE).Value = "DIE"
                        tempWS.Range("I" & rowDIE).Formula = "=ABS(" & dieAmountPM & "+" & dieAmountCM & "+" & InOutVal & ")"
                        tempWS.Range("H" & rowDIE).Formula = "=IF(" & dieAmountPM & "+" & dieAmountCM & "+" & InOutVal & "<0,50,40)"
                        
                        tempWS.Range("F" & rowDIE).Formula = "=IF(H" & rowDIE & "=50,""A8381000"",""A8382000"")" ' dieAmountPM & "+" & dieAmountCM & "+" & InOutVal & "<0,50,40)"
                        
                        Select Case myCoCd
                        Case "AU01"
                            tempWS.Range("G" & rowDIE).Value = "120378"
                        Case "AU10"
                            tempWS.Range("G" & rowDIE).Value = "120998"
                        Case "AU11"
                            tempWS.Range("G" & rowDIE).Value = "120388"
                        End Select
                        
                    Else
                        tempWS.Range("A" & rowDIE).EntireRow.Delete
                    End If
                    
                    F03Range.EntireColumn.RemoveDuplicates Columns:=1, Header:=xlNo
                    Call SAP_F03(myGL, myCoCd, tempWS, rowDIE, docRow, F03Range)
                    
                Else
                    inputError = inputError & vbLf & " - " & myCoCd & "-" & myGL & ": DIE = " & Round(dieAmountPM + dieAmountCM + InOutVal, 2)
                End If
                
                
            End If
        End If
        
        
    Next
    
    
    mainWB.Save
    mainWB.Close True
    
End Sub



Sub SAP_F03(myGL As String, myCoCd As String, ws As Worksheet, rowDIE As Long, docRow As Long, refCell As Range)
    Dim xRow As Long, xCount As Long, xlastrow As Long
    Dim amtAssigned As Double, amtNotAssigned As Double
    
    
    With sessBP
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/okcd").Text = "F-03"
        .findById("wnd[0]").sendVKey 0
        
        .findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[2,0]").Select 'select by doc no
        
        .findById("wnd[0]/usr/ctxtRF05A-AGKON").Text = myGL
        .findById("wnd[0]/usr/ctxtBKPF-BUDAT").Text = Format(Date, "dd.mm.yyyy")
        .findById("wnd[0]/usr/txtBKPF-MONAT").Text = Month(Date)
        .findById("wnd[0]/usr/ctxtBKPF-BUKRS").Text = myCoCd
        .findById("wnd[0]/usr/ctxtBKPF-WAERS").Text = "AUD"
        .findById("wnd[0]/tbar[1]/btn[16]").press
        
        If Left(.findById("wnd[0]/sbar").Text, 53) = "You have no authorization for posting in company code" Then
            inputError = inputError & vbLf & " - " & myCoCd & "-" & myGL & ": " & .findById("wnd[0]/sbar").Text
            ws.Range("E" & docRow).Value = .findById("wnd[0]/sbar").Text & ". Item not cleared."
            Exit Sub
        End If
        
        xRow = 0
        xCount = 0
        While Not refCell.Offset(xRow, 0).Value = ""
            .findById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[" & xCount & ",0]").Text = refCell.Offset(xRow, 0).Value
            
            xRow = xRow + 1
            xCount = xCount + 1
            If xCount = 26 And Not refCell.Offset(xRow, 0).Value = "" Then
                xCount = 0
                .findById("wnd[0]").sendVKey 0
            End If
        Wend

        .findById("wnd[0]/tbar[1]/btn[16]").press
        
        If ws.Range("I" & rowDIE - 1).Value <> "" Or ws.Range("I" & rowDIE).Value <> "" Then 'if there's DIE and input/output
            .findById("wnd[0]").sendVKey 7
            
            If ws.Range("I" & rowDIE - 1).Value <> "" And ws.Range("I" & rowDIE).Value <> "" Then
                'do DIE first
                .findById("wnd[0]/usr/txtBKPF-XBLNR").Text = "GST CLR " & UCase(Format(DateAdd("m", -1, Date), "MMMYYYY"))
                .findById("wnd[0]/usr/txtBKPF-BKTXT").Text = "GST_" & UCase(Format(DateAdd("m", -1, Date), "MMMYYYY")) & "CLEARING"
                
                .findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = ws.Range("H" & rowDIE).Value
                .findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = ws.Range("F" & rowDIE).Value
                .findById("wnd[0]").sendVKey 0
                
                .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Round(ws.Range("I" & rowDIE).Value, 2) 'amount
                .findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = "DIE"
                .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = ws.Range("G" & rowDIE).Value 'cost center
                
                
                'do input output
                .findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = ws.Range("H" & rowDIE - 1).Value
                .findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = ws.Range("F" & rowDIE - 1).Value
                .findById("wnd[0]").sendVKey 0
                
                .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Abs(Round(ws.Range("I" & rowDIE - 1).Value, 2))
                .findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = "FLEETPLUS"
                .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = 120378
                
                .findById("wnd[0]/tbar[1]/btn[16]").press

            ElseIf ws.Range("I" & rowDIE - 1).Value <> "" Then 'got input or output
                .findById("wnd[0]/usr/txtBKPF-XBLNR").Text = "GST CLR " & UCase(Format(DateAdd("m", -1, Date), "MMMYYYY"))
                .findById("wnd[0]/usr/txtBKPF-BKTXT").Text = "GST_" & UCase(Format(DateAdd("m", -1, Date), "MMMYYYY")) & "CLEARING"
                
                .findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = ws.Range("H" & rowDIE - 1).Value
                .findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = ws.Range("F" & rowDIE - 1).Value
                .findById("wnd[0]").sendVKey 0
                
                .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Abs(Round(ws.Range("I" & rowDIE - 1).Value, 2))
                .findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = "FLEETPLUS"
                .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = 120378
                
                .findById("wnd[0]/tbar[1]/btn[16]").press
            
            ElseIf ws.Range("I" & rowDIE).Value <> "" Then 'got DIE
                .findById("wnd[0]/usr/txtBKPF-XBLNR").Text = "GST CLR " & UCase(Format(DateAdd("m", -1, Date), "MMMYYYY"))
                .findById("wnd[0]/usr/txtBKPF-BKTXT").Text = "GST_" & UCase(Format(DateAdd("m", -1, Date), "MMMYYYY")) & "CLEARING"
                
                .findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = ws.Range("H" & rowDIE).Value
                .findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = ws.Range("F" & rowDIE).Value
                .findById("wnd[0]").sendVKey 0
                
                .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Round(ws.Range("I" & rowDIE).Value, 2) 'amount
                .findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = "DIE"
                .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = ws.Range("G" & rowDIE).Value 'cost center
                
                .findById("wnd[0]/tbar[1]/btn[16]").press
            
            End If

            'check value
            If Right(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-AKTIV").Text, 1) = "-" Then
                amtAssigned = 1 * Left(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-AKTIV").Text, Len(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-AKTIV").Text) - 1)
            Else
                amtAssigned = 1 * .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-AKTIV").Text
            End If
            
            If Right(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text, 1) = "-" Then
                amtNotAssigned = 1 * Left(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text, Len(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text) - 1)
            Else
                amtNotAssigned = 1 * .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text
            End If
            
            If Not 1 * .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text = 0 Then
                
                
                inputError = inputError & vbLf & " - " & myCoCd & "-" & myGL & ": Unassigned amount = " & .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text
                ws.Range("E" & docRow).Value = "Unassigned amount at " & .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text & ". Item not cleared."
            
            Else
                
                'to ammend this button
                .findById("wnd[0]/tbar[0]/btn[11]").press
                ws.Range("E" & docRow).Value = .findById("wnd[0]/sbar").Text
            End If

        Else
            'need to click on blue item
            If Right(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-AKTIV").Text, 1) = "-" Then
                amtAssigned = 1 * Left(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-AKTIV").Text, Len(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-AKTIV").Text) - 1)
            Else
                amtAssigned = 1 * .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-AKTIV").Text
            End If
        
            If Right(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text, 1) = "-" Then
                amtNotAssigned = 1 * Left(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text, Len(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text) - 1)
            Else
                amtNotAssigned = 1 * .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text
            End If
        
            If (amtNotAssigned = 0 And amtAssigned = 0) Or amtAssigned = amtNotAssigned Then
                If amtAssigned = amtNotAssigned And Not amtAssigned = 0 Then
                    xRow = 0
                    'Debug.Print .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/tblSAPDF05XTC_6103/ctxtRFOPS_DK-BUDAT[4,2]").Text
                    While Mid(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/tblSAPDF05XTC_6103/ctxtRFOPS_DK-BUDAT[4," & xRow & "]").Text, 3, 1) = "."
                    'While Not .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/tblSAPDF05XTC_6103/txtDF05B-PSBET[1," & xRow & "]").Text = ""
                        
                        If Right(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/tblSAPDF05XTC_6103/ctxtRFOPS_DK-BUDAT[4," & xRow & "]").Text, 7) = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "mm.yyyy") Then
                        'If .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/tblSAPDF05XTC_6103/txtDF05B-PSBET[1," & xRow & "]").Text = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "dd.mm.yyyy") Then
                            
                            .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/tblSAPDF05XTC_6103/txtDF05B-PSBET[6," & xRow & "]").SetFocus
                            .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/tblSAPDF05XTC_6103/txtDF05B-PSBET[6," & xRow & "]").SetFocus
                            .findById("wnd[0]").sendVKey 2
                            .findById("wnd[0]").sendVKey 2
                            
                            
                            '.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/tblSAPDF05XTC_6103/txtDF05B-PSBET[6," & xRow & "]").caretPosition = 2
                            '.findById("wnd[0]").SENDVKEY 0
                            '.findById("wnd[0]").SENDVKEY 2
                            
                        End If
            
                        'to scroll ???
                        If xRow = 25 Then
                            .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/tblSAPDF05XTC_6103").verticalScrollbar.Position = .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/tblSAPDF05XTC_6103").verticalScrollbar.Position + 25
                            xRow = 0
                        Else
                            xRow = xRow + 1
                        End If
                    Wend
                End If
    
                If Not .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text * 1 = 0 Then ' Or .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text = "0.00") Then
                    
                    inputError = inputError & vbLf & " - " & myCoCd & "-" & myGL & ": Unassigned amount = " & .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text
                    ws.Range("E" & docRow).Value = "Unassigned amount at " & .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text & ". Item not cleared"
                Else
                    .findById("wnd[0]/tbar[0]/btn[11]").press
                    ws.Range("E" & docRow).Value = .findById("wnd[0]/sbar").Text
                End If
                
                Else
                    'error
                    inputError = inputError & vbLf & " - " & myCoCd & "-" & myGL & ": Unassigned amount = " & .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text
                    ws.Range("E" & docRow).Value = "Unassigned amount at " & .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text & ". Item not cleared."
                
                End If
            
            End If
        
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
    End With


End Sub





Sub SAP_Clearing(myGL As String, myCoCd As String)
    
    Dim xRow As Long, xCount As Long, xlastrow As Long
    Dim amtAssigned As String, amtNotAssigned As String
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "F-03"
    sessBP.findById("wnd[0]").sendVKey 0
    
    sessBP.findById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[2,0]").Select 'SELECT DOCUMENT NUMBER
    sessBP.findById("wnd[0]/usr/ctxtRF05A-AGKON").Text = myGL
    sessBP.findById("wnd[0]/usr/ctxtBKPF-BUDAT").Text = Format(Date, "dd.mm.yyyy")
    sessBP.findById("wnd[0]/usr/txtBKPF-MONAT").Text = Month(Date)
    sessBP.findById("wnd[0]/usr/ctxtBKPF-BUKRS").Text = myCoCd
    sessBP.findById("wnd[0]/usr/ctxtBKPF-WAERS").Text = "AUD"
    sessBP.findById("wnd[0]/tbar[1]/btn[16]").press
    
    
    
    
    
    
    
    
    
    
    
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    
End Sub























































'**************************************************************************************************************
'SAP Scripts
'**************************************************************************************************************


Sub SAP_FBL3N(myCoCd As String, myGL As String) ', resumeItem As Boolean)
    Dim myFileName As String
    myFileName = myGL & " " & myCoCd & " " & Format(DateAdd("m", -1, Date), "MMM YYYY") & ".xlsx"
    
    'If resumeItem = False Then
    '    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    '    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    '    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    '
    '    sessBP.findById("wnd[0]/tbar[0]/okcd").Text = "FBL3N"
    '    sessBP.findById("wnd[0]").sendVKey 0
    'End If
    
    sessBP.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").Text = myGL
    sessBP.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").Text = myCoCd
    sessBP.findById("wnd[0]/usr/ctxtPA_STIDA").Text = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "dd.mm.yyyy")
    
    If Not (myGL = "A3499000" And myCoCd = "AU01") Then
        sessBP.findById("wnd[0]/usr/ctxtPA_VARI").Text = "NAS Clearing"
        sessBP.findById("wnd[0]/usr/txtPA_NMAX").Text = "99999999"
    End If
    
    
    sessBP.findById("wnd[0]/tbar[1]/btn[8]").press
    
    
    If Left(sessBP.findById("wnd[0]/sbar").Text, 17) = "No items selected" Then
        'no report
        'need to screenshot
        sessBP.findById("wnd[0]").HardCopy myLoc & myGL & " " & myCoCd & " - No FBL3N Report.jpg", 1
        
    Else
        'download a report and save into the report
        'If myGL = "A3499000" And myCoCd = "AU01" Then
        If firstDownload = False Then
            firstDownload = True
            
            sessBP.findById("wnd[0]/tbar[1]/btn[32]").press
            sessBP.findById("wnd[1]/usr/btnAPP_FL_ALL").press
            
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "COMPANY CODE"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 12
            sessBP.findById("wnd[2]/tbar[0]/btn[0]").press
            
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "GL ACCOUNT"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 3
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "G/L ACCOUNT"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 2
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "document number"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 15
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "document type"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 13
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "posting date"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 12
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "text"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 4
            sessBP.findById("wnd[2]/tbar[0]/btn[0]").press
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "profit center"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 13
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "amount in loc.curr.2"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 20
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "amount in local currency"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 24
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "business area"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 13
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "clearing document"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 17
            sessBP.findById("wnd[2]/tbar[0]/btn[0]").press
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "clearing date"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 13
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "status"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 6
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 108
            sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").getAbsoluteRow(115).Selected = True
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/tblSAPLSKBHTC_FIELD_LIST").verticalScrollbar.Position = 0
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "year/month"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 10
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Eff.exchange rate"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 17
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document currency"
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/usr/btnB_SEARCH").press
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "user name"
            sessBP.findById("wnd[2]/usr/txtGD_SEARCHSTR").caretPosition = 9
            sessBP.findById("wnd[2]").sendVKey 0
            sessBP.findById("wnd[1]/usr/btnAPP_WL_SING").press
            sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
            
            'save template
            sessBP.findById("wnd[0]/tbar[1]/btn[36]").press
            sessBP.findById("wnd[1]/usr/ctxtLTDX-VARIANT").Text = "NAS Clearing"
            sessBP.findById("wnd[1]/usr/txtLTDXT-TEXT").Text = "Standard NAS Clearing"
            sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
            
            If Not sessBP.findById("wnd[2]", False) Is Nothing Then
                sessBP.findById("wnd[2]/tbar[0]/btn[0]").press
            End If
        End If
        
        
        sessBP.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
        If Not sessBP.findById("wnd[1]/usr/radRB_OTHERS", False) Is Nothing Then
            sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
        End If
        sessBP.findById("wnd[1]/usr/ctxtDY_PATH").Text = myLoc
        sessBP.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = myFileName
        sessBP.findById("wnd[1]/tbar[0]/btn[0]").press
        
        sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
        sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    End If
    
    'If myCoCd = "AU01" And myGL = "A3780007" Then
    '    sessBP.findById("wnd[0]/tbar[0]/btn[3]").press
    'End If
    
End Sub



