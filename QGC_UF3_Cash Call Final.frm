VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf3 
   Caption         =   "Australia QGC Automation Tool Run 3"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16605
   OleObjectBlob   =   "QGC_UF3_Cash Call Final.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf3"
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
Private wsRef As Worksheet


Private wbCash As Workbook
Private wsC_Journal As Worksheet, wsC_Forecasting As Worksheet

Private myPath As String, myPathCC As String
Private myFileName As String

Private myFolderNames(0 To 10) As String

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
    
    
    doCashFinal
    
    emailCash
    
    wbCash.Close True
    
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
    
    If Not Environ("username") = wsSettings.Range("B1").Value Then
        wsSettings.Range("B:B").EntireColumn.Clear
        wsSettings.Range("B1").Value = Environ("username")
    Else
        Me.tbSaveLocation.Value = wsSettings.Range("B6").Value
    End If
    
End Sub


Private Sub BtnCancel_Click()
    Unload Me
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
    
End Sub


Sub openFiles()
    Dim xRow As Long

    If Dir(Me.tbSaveLocation, vbDirectory) = "" Then
        MsgBox ("Source folder not found. Please ensure the correct file path is selected.")
        End
    End If
    
    myFileName = Format(DateAdd("m", -1, Date), "MMM YYYY") & " BAS - Cash Call Estimates.xlsx"
    
    
    If Not Dir(Me.tbSaveLocation.Value & myFileName) = "" Then
        myPath = Left(Left(myPath, InStrRev(myPath, "\") - 1), InStrRev(Left(myPath, InStrRev(myPath, "\") - 1), "\"))
        'mypath = Me.tbSaveLocation.Value
    ElseIf Not Dir(Me.tbSaveLocation.Value & "Cash Call\" & myFileName) = "" Then
        myPath = Me.tbSaveLocation.Value '& "Cash Call\"
    ElseIf Not Dir(Me.tbSaveLocation.Value & Format(DateAdd("m", -1, Date), "mm mmm") & "\Cash Call\" & myFileName) = "" Then
        myPath = Me.tbSaveLocation.Value & Format(DateAdd("m", -1, Date), "mm mmm") & "\" 'Cash Call\"
    ElseIf Not Dir(Me.tbSaveLocation.Value & Format(DateAdd("m", -1, Date), "yyyy") & "\" & Format(DateAdd("m", -1, Date), "mm mmm") & "\Cash Call\" & myFileName) = "" Then
        myPath = Me.tbSaveLocation.Value & Format(DateAdd("m", -1, Date), "yyyy") & "\" & Format(DateAdd("m", -1, Date), "mm mmm") & "\" 'Cash Call\"
    Else
        MsgBox ("Source file not found. Please ensure the correct file path is selected.")
        End
    End If
    
    
    Set wbCash = Workbooks.Open(myPath & "Cash Call\" & myFileName, ReadOnly:=True)
    
    On Error GoTo ErrorFound
    Set wsC_Journal = wbCash.Worksheets("Journal Entries by BAS Group")
    Set wsC_Forecasting = wbCash.Worksheets("Cash forecasting")
    On Error GoTo 0
    
    
    wbCash.SaveAs myPath & "Cash Call\" & Format(DateAdd("m", -1, Date), "MMM YYYY") & " BAS - Cash Call Final.xlsx"
    
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
    
    
    
    
    Exit Sub
ErrorFound:
    
    wbCash.Close False
    MsgBox ("Error detected in Cash Call file. Please ensure the correct file path is selected.")
    End
    
End Sub


Sub doCashFinal()
    Dim thisFolderName As String, curFileName As String
    Dim myFile As String
    Dim wb As Workbook, ws As Worksheet, ws2 As Worksheet
    Dim myWs As Worksheet
    Dim myCell As Range
    Dim xCount As Long
    
    
    
    
    
    For xCount = 0 To 10
        If Not xCount = 2 Then
            thisFolderName = Dir(myPath & myFolderNames(xCount))
            Do While Not thisFolderName = ""
                If Right(thisFolderName, 24) = "Input Form " & Format(DateAdd("m", -1, Date), "mmm yyyy") & ".xlsx" Then
                    curFileName = thisFolderName
                    Exit Do
                End If
                thisFolderName = Dir
            Loop
            
            If curFileName <> "" Then
                Set wb = Workbooks.Open(myPath & myFolderNames(xCount) & curFileName, ReadOnly:=True)
                Select Case xCount
                Case 0
                    '1100 - BGIA (QGC Upstream)
                    Set ws = wb.Worksheets("Data Entry - BGIA Group")
                    
                    Set ws = wb.Worksheets("Data Entry - 1100")
                    wsC_Journal.Range("G54").Value = ws.Range("J38").Value
                    
                    Set ws = wb.Worksheets("Data Entry - 1106")
                    wsC_Journal.Range("I56").Value = ws.Range("J38").Value
                    
                    Set ws = wb.Worksheets("Data Entry - 1122")
                    wsC_Journal.Range("K58").Value = ws.Range("J38").Value
                    
                Case 1
                    '5000 - QGC Group
                    Set ws = wb.Worksheets("Data Entry - 5000")
                    Set ws2 = wb.Worksheets("Col E Adj G11")
                    
                    wsC_Journal.Range("G42").Value = ws2.Range("C6").Value
                    wsC_Journal.Range("G43").Formula = "=SUMIF('[" & wb.Name & "]" & ws2.Name & "'!C:C,""Total GST Payable to Fleetplus"",'[" & wb.Name & "]" & ws2.Name & "'!F:F)"
                    wsC_Journal.Range("G44").Value = ws.Range("J37").Value
                    wsC_Journal.Range("G45").Formula = "= " & ws.Range("J38").Value & " - G43"
                    
                    
                    Set ws = wb.Worksheets("Data Entry - 5001")
                    wsC_Journal.Range("H46").Value = ws.Range("J38").Value
                    
                    Set ws = wb.Worksheets("Data Entry - 5002")
                    wsC_Journal.Range("I47").Value = ws.Range("J38").Value
                    
                    Set ws = wb.Worksheets("Data Entry - 5007")
                    wsC_Journal.Range("J48").Value = ws.Range("J38").Value
                    
                    Set ws = wb.Worksheets("Data Entry - QGC JV")
                    wsC_Journal.Range("G4").Value = ws.Range("J38").Value
                    
                    Set ws = wb.Worksheets("Data Entry - JV 171")
                    wsC_Journal.Range("G9").Value = ws.Range("J38").Value
                    
                    
                Case 3
                    '5030 - Toll Co 2
                    Set ws = wb.Worksheets("Data Entry - 5030")
                    wsC_Journal.Range("G34").Value = ws.Range("J38").Value
                    
                Case 4
                    '5031 - Toll Co 2 (2)
                    Set ws = wb.Worksheets("Data Entry - 5031")
                    wsC_Journal.Range("G38").Value = ws.Range("J38").Value
                    
                Case 5
                    '5033 - Toll Co 1
                    Set ws = wb.Worksheets("Data Entry - 5033")
                    wsC_Journal.Range("G30").Value = ws.Range("J38").Value
                    
                Case 6
                    '5036 - QCLNG (OpCo)
                    Set ws = wb.Worksheets("Data Entry - QCLNG Group")
                    
                    Set ws = wb.Worksheets("Data Entry - 5036")
                    wsC_Journal.Range("G62").Value = ws.Range("J38").Value
                    
                    Set ws = wb.Worksheets("Data Entry - 5039")
                    wsC_Journal.Range("H63").Value = ws.Range("J38").Value
                    
                Case 7
                    '5037 - Train 1
                    Set ws = wb.Worksheets("Data Entry - 5037")
                    wsC_Journal.Range("G22").Value = ws.Range("J38").Value
                    
                Case 8
                    '5038 - Train 2
                    Set ws = wb.Worksheets("Data Entry - 5038")
                    wsC_Journal.Range("G26").Value = ws.Range("J38").Value
                    
                Case 9
                    '5045 - T1 UJV
                    Set ws = wb.Worksheets("Data Entry - 5045")
                    wsC_Journal.Range("G14").Value = ws.Range("J38").Value
                    
                Case 10
                    '5046 - T2 UJV
                    Set ws = wb.Worksheets("Data Entry - 5046")
                    wsC_Journal.Range("G18").Value = ws.Range("J38").Value
                    
                End Select
                
                
                wb.Close False
                Set wb = Nothing
                Set ws = Nothing
                Set ws2 = Nothing
            End If
        End If
        curFileName = ""
    Next
    
    
    
End Sub





'**************************************************************************************************************
'Email Creation
'**************************************************************************************************************


Sub emailCash()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strBody As String, strBody2 As String
    Dim sTo As String
    Dim sCC As String
    Dim myDateVal As String
    Dim xCount As Long, xCount2 As Long
    
    sTo = "J.Jamaluddin@shell.com; SBSC-CMKL-GSS-TH-AU@shell.com; R.Paskaradass@shell.com"
    sCC = "Stephanie.KuahMeiYen@shell.com; M-S.Hanapi@shell.com; Amir.Jamal@shell.com; Kay.Pfingst@shell.com; Jonathan.Soosay@shell.com; Irina.Ilushin@shell.com; Pey-Shy.Lee@shell.com; GXQGCJVFinance@shell.com; Kamini.Raja@shell.com; Zi-Yang.Puah@shell.com; Norzeiny.M-Zain@shell.com"
    
    ' HTML before rows
    
    
    strBody = "<html><body><b>Hi all,</b><p>Please refer below for actual.<br><br><br>"
    
    strBody = strBody & "<b><u>BAS - Final</u></b><br><br>"
    
    strBody = strBody & "<head><style>table, th, td {border: 1px solid black;}" & _
        "<table style=""width:42%""><tr>" & _
        "<b><th bgcolor=""#ff8500"">Coy</th></b>" & _
        "<b><th bgcolor=""#ff8500"">Entity</th></b>" & _
        "<b><th bgcolor=""#ff8500"">(Pay)/ Refund</th></b>" & _
        "<b><th bgcolor=""#ff8500"">Confirmed</th></tr></b>"
    
    
    ' iterate collection
    For xCount = 6 To 20
        strBody = strBody & "<tr>"
        strBody = strBody & "<td ""col width=5%"" align=""center"">" & wsC_Forecasting.Range("B" & xCount).Value & "</td>"
        strBody = strBody & "<td ""col width=17%"" align=""center"">" & wsC_Forecasting.Range("C" & xCount).Value & "</td>"
        
        If wsC_Forecasting.Range("D" & xCount).Value < 0 Then
            strBody = strBody & "<td ""col width=10%"" align=""center""><font color=""red"">" & Format(wsC_Forecasting.Range("D" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
        Else
            strBody = strBody & "<td ""col width=10%"" align=""center"">" & Format(wsC_Forecasting.Range("D" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
        End If
        
        
        strBody = strBody & "<td ""col width=10%"" align=""center"">" & Format(wsC_Forecasting.Range("E" & xCount).Value, "dd-mmm-yyyy") & "</td>"
        strBody = strBody & "</tr>"
    Next
    
    strBody = strBody & "<tr>"
    strBody = strBody & "<td ""col width=5%"" align=""center""><b>" & wsC_Forecasting.Range("B21").Value & "</b></td>"
    strBody = strBody & "<td ""col width=17%"" align=""center""><b>" & wsC_Forecasting.Range("C21").Value & "</b></td>"
    
    If wsC_Forecasting.Range("D21").Value < 0 Then
        strBody = strBody & "<td ""col width=10%"" align=""center""><font color=""red""><b>" & Format(wsC_Forecasting.Range("D21").Value, "#,##0.00;(#,##0.00)") & "</b></td>"
    Else
        strBody = strBody & "<td ""col width=10%"" align=""center""><b>" & Format(wsC_Forecasting.Range("D21").Value, "#,##0.00;(#,##0.00)") & "</b></td>"
    End If
    
    strBody = strBody & "<td ""col width=10%""  align=""center""><b>" & Format(wsC_Forecasting.Range("E21").Value, "dd-mmm-yyyy") & "</b></td>"
    strBody = strBody & "</tr>"
    
    
    strBody = strBody & "</table><br><br>"
    
    strBody2 = "<b><u>Journal Entries by BAS Group</u></b><br><br>"
    
    
    For xCount = 3 To 64
        Select Case xCount
        Case 3, 8, 13, 17, 21, 25, 29, 33, 37, 41, 53, 61
            If xCount < 41 Or xCount = 61 Then
                strBody2 = strBody2 & "<head><style>table, th, td {border: 1px solid black;}" & _
                "<table style=""width:61%""><tr>" & _
                "<b><th colspan=""2"" bgcolor=""#ff8500"">" & wsC_Journal.Range("B" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">GL</th></b>" & _
                "<b><th bgcolor=""#ff8500"">PC</th></b>" & _
                "<b><th bgcolor=""#ff8500"">Pkey</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("G" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("H" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">Total</th></tr></b>"
                
            ElseIf xCount = 41 Then
                strBody2 = strBody2 & "<head><style>table, th, td {border: 1px solid black;}" & _
                "<table style=""width:89%""><tr>" & _
                "<b><th colspan=""2"" bgcolor=""#ff8500"">" & wsC_Journal.Range("B" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">GL</th></b>" & _
                "<b><th bgcolor=""#ff8500"">PC</th></b>" & _
                "<b><th bgcolor=""#ff8500"">Pkey</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("G" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("H" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("I" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("J" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("K" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("L" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">Total</th></tr></b>"
                
            ElseIf xCount = 53 Then
                strBody2 = strBody2 & "<head><style>table, th, td {border: 1px solid black;}" & _
                "<table style=""width:86%""><tr>" & _
                "<b><th colspan=""2"" bgcolor=""#ff8500"">" & wsC_Journal.Range("B" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">GL</th></b>" & _
                "<b><th bgcolor=""#ff8500"">PC</th></b>" & _
                "<b><th bgcolor=""#ff8500"">Pkey</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("G" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("H" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("I" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("J" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">" & wsC_Journal.Range("K" & xCount).Value & "</th></b>" & _
                "<b><th bgcolor=""#ff8500"">Total</th></tr></b>"
                
            End If
            
        Case 4 To 5, 9 To 10, 14, 18, 22, 26, 30, 34, 38, 42 To 50, 54 To 58, 62 To 63
            'content
            strBody2 = strBody2 & "<tr>"
            strBody2 = strBody2 & "<td ""col width=5%"" align=""center"">" & wsC_Journal.Range("B" & xCount).Value & "</td>"
            strBody2 = strBody2 & "<td ""col width=17%"" align=""center"">" & wsC_Journal.Range("C" & xCount).Value & "</td>"
            strBody2 = strBody2 & "<td ""col width=8%"" align=""center"">" & wsC_Journal.Range("D" & xCount).Value & "</td>"
            strBody2 = strBody2 & "<td ""col width=5%"" align=""center"">" & wsC_Journal.Range("E" & xCount).Value & "</td>"
            
            If wsC_Journal.Range("F" & xCount).Value = 40 Then
                strBody2 = strBody2 & "<td ""col width=5%"" align=""center""><font color=""red"">" & wsC_Journal.Range("F" & xCount).Value & "</td>"
            Else
                strBody2 = strBody2 & "<td ""col width=5%"" align=""center"">" & wsC_Journal.Range("F" & xCount).Value & "</td>"
            End If
            
            If wsC_Journal.Range("G" & xCount).Value < 0 Then
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("G" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            Else
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("G" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            End If
            
            If wsC_Journal.Range("H" & xCount).Value < 0 Then
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("H" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            Else
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("H" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            End If
            
            If wsC_Journal.Range("I" & xCount).Value < 0 Then
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("I" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            Else
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("I" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            End If
            
            
            
            If xCount >= 42 And xCount <= 58 Then
                
                If wsC_Journal.Range("J" & xCount).Value < 0 Then
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("J" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                Else
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("J" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                End If
                
                
                If wsC_Journal.Range("K" & xCount).Value < 0 Then
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("K" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                Else
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("K" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                End If
                
                
                If wsC_Journal.Range("L" & xCount).Value < 0 Then
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("L" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                Else
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("L" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                End If
            
                
                If xCount >= 42 And xCount <= 50 Then
                    
                    If wsC_Journal.Range("M" & xCount).Value < 0 Then
                        strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("M" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                    Else
                        strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("M" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                    End If
                    
                End If
            End If
            
            strBody2 = strBody2 & "</tr>"
            
        Case 6, 11, 15, 19, 23, 27, 31, 35, 39, 51, 59, 64
            strBody2 = strBody2 & "<tr>"
            strBody2 = strBody2 & "<td ""col width=35%"" align=""left"" colspan=""5""><b>Total</b></td>"
            
            If wsC_Journal.Range("G" & xCount).Value < 0 Then
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("G" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            Else
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("G" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            End If
            
            If wsC_Journal.Range("H" & xCount).Value < 0 Then
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("H" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            Else
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("H" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            End If
            
            If wsC_Journal.Range("I" & xCount).Value < 0 Then
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("I" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            Else
                strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("I" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
            End If
            
            
            If xCount = 51 Then
                If wsC_Journal.Range("J" & xCount).Value < 0 Then
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("J" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                Else
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("J" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                End If
                
                
                If wsC_Journal.Range("K" & xCount).Value < 0 Then
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("K" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                Else
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("K" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                End If
                
                
                If wsC_Journal.Range("L" & xCount).Value < 0 Then
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("L" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                Else
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("L" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                End If
                
                If wsC_Journal.Range("M" & xCount).Value < 0 Then
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("M" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                Else
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("M" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                End If
                
            ElseIf xCount = 59 Then
                If wsC_Journal.Range("J" & xCount).Value < 0 Then
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("J" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                Else
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("J" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                End If
                
                
                If wsC_Journal.Range("K" & xCount).Value < 0 Then
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("K" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                Else
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("K" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                End If
                
                
                If wsC_Journal.Range("L" & xCount).Value < 0 Then
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center""><font color=""red"">" & Format(wsC_Journal.Range("L" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                Else
                    strBody2 = strBody2 & "<td ""col width=7%"" align=""center"">" & Format(wsC_Journal.Range("L" & xCount).Value, "#,##0.00;(#,##0.00)") & "</td>"
                End If
                
            End If
            
            strBody2 = strBody2 & "</table><br>"
            
        End Select
    Next
    
    
    
    strBody2 = strBody2 & "<br>Regards,<br>" & Replace(Left(Application.UserName, InStr(1, Application.UserName, " ", vbTextCompare) - 1), ",", "") & "</body></html>"
    'strbody = strbody & "</table><br><br>Regards,<br>" & Application.UserName & "</body></html>"
    
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = sTo
        .cc = sCC
        .Subject = Format(DateAdd("m", -1, Date), "MMMM YYYY") & " - Cash Flow Actual"
        .HTMLBody = strBody & strBody2 'strbody & OutMail.HTMLBody
        .Display
    End With
    
End Sub
    





