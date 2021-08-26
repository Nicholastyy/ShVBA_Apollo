VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf3_Posting 
   Caption         =   "Australia SAPL & SEHAL Automation: Posting"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14610
   OleObjectBlob   =   "NAS_UF2_Posting.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "uf3_Posting"
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

Private wbCross As Workbook
Private wsSettlement As Worksheet

Private TestEnv As Boolean
Private myLoc As String

Private wbFBB1 As Workbook
Private wsFBB1 As Worksheet, wsRef As Worksheet

Private wsSettings As Worksheet

Private mySess



Private Sub BtnRun_Click()
    Set wsRef = ThisWorkbook.Worksheets("FBB1")
    
    StartTime = Timer
    inputError = ""
    
    checkFields
    
    If inputError <> "" Then
        MsgBox ("Please ensure the following fields are filled before proceeding:" & inputError)
        Exit Sub
    End If
    
    TestEnv = False
    
    checkOpenSAP
    
    'If openC16 = True Then
    '    Set mySess = sessC16
    '    TestEnv = True
    'ElseIf openA16 = True Then
    '    Set mySess = sessA16
    '    TestEnv = True
    'ElseIf openU16 = True Then
    '    Set mySess = sessU16
    '    TestEnv = True
    'Else
    '    MsgBox ("Please ensure SAP C16, U16 or A16 is open (BP Test environment)")
    '    Exit Sub
    'End If
    If openP16 = False Then
    '    If openA16 = True Then
    '        If MsgBox("Using A16 test environment. Continue?", vbYesNo) = vbYes Then
    '            Set sessBP = sessA16
    '        Else
        
    '            Exit Sub
    '        End If
    '    Else
            MsgBox ("Please ensure you have SAP Blueprint open.")
            Exit Sub
        Else
            Set mySess = sessBP
        End If
    'End If
    
    
    RunPauseAll
    
    openFile
    
    
    doPosting
    
    wbCross.Close True
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MinutesElapsed = VBA.Format((Timer - StartTime) / 86400, "hh:mm:ss")
    
    RunActivateAll
    If inputError = "" Then
        MsgBox ("Automation process completed for Australia SAPL & SEHAL Preparation Posting process." & vbLf & "Time taken: " & MinutesElapsed)
    Else
        MsgBox ("Posting process Completed with time taken: " & MinutesElapsed & vbLf & "Errors list: " & inputError)
    End If
    Unload Me
    
    
End Sub





Private Sub BtnCancel_Click()
    Unload Me
End Sub




Private Sub btnSaveLocation_Click()
    Me.tbSaveLocation.Value = SearchFolderLocation
    wsSettings.Range("B5").Value = Me.tbSaveLocation.Value
End Sub


Private Sub btnCrossCheck_Click()
    Me.tbCrosscheck.Value = SearchFileLocation
    wsSettings.Range("B6").Value = Me.tbCrosscheck.Value
End Sub





Private Sub UserForm_Initialize()
    Set wsSettings = ThisWorkbook.Worksheets("Settings")
    
    If Not VBA.Environ("username") = wsSettings.Range("B1").Value Then
        wsSettings.Range("B:B").EntireColumn.Clear
        wsSettings.Range("B1").Value = VBA.Environ("username")
    Else
        Me.tbSaveLocation.Value = wsSettings.Range("B5").Value
        Me.tbCrosscheck.Value = wsSettings.Range("B6").Value
    End If
    
    
End Sub





Sub checkFields()
    If Me.tbSaveLocation.Value = "" Then
        inputError = inputError & vbLf & " - Save Location (1)"
    ElseIf Dir(Me.tbSaveLocation.Value, vbDirectory) = "" Then
        inputError = inputError & vbLf & " - Save Location (2)"
    End If
    
    If Me.tbCrosscheck.Value = "" Then
        inputError = inputError & vbLf & " - Cross Check File (1)"
    ElseIf Dir(Me.tbCrosscheck.Value) = "" Then
        inputError = inputError & vbLf & " - Cross Check File (2)"
    End If
    
    
    
End Sub


Sub makeFolder()
    Dim xCount As Long
    
    myLoc = Me.tbSaveLocation.Value & "FBB1 Journals/"
    xCount = 0
    
    While Dir(myLoc, vbDirectory) <> ""
        xCount = xCount + 1
        myLoc = Me.tbSaveLocation.Value & "FBB1 Journals (" & xCount & ")/"
    Wend
    
    MkDir myLoc
    
End Sub


Sub openFile()
    Dim ws As Worksheet
    Dim myName As String
    Dim xCount As Long
    
    Set wbCross = Workbooks.Open(Me.tbCrosscheck.Value, ReadOnly:=True)
    
    For Each ws In wbCross.Worksheets
        If ws.Name = "Settlement Jnl" Then
            Set wsSettlement = ws
            Exit For
        End If
    Next
    
    If wsSettlement Is Nothing Then
        wbCross.Close False
        Call RunActivateAll
        MsgBox ("Tab missing in CrossCheck file. Please ensure the right file is selected.")
        End
    End If
    
    xCount = 1
    myName = "GST Crosscheck_" & Format(DateAdd("m", -1, Date), "yyyy_mm_mmmm") & "_SETL_SAPLJV_SEHAL.xlsb"
    While Dir(Me.tbSaveLocation.Value & myName) <> ""
        myName = "GST Crosscheck_" & Format(DateAdd("m", -1, Date), "yyyy_mm_mmmm") & "_SETL_SAPLJV_SEHAL (" & xCount & ").xlsb"
        xCount = xCount + 1
    Wend
    
    wbCross.SaveAs tbSaveLocation & myName
    
    
End Sub



Sub doPosting()
    Dim myRef As String
    
    Dim myRow As Long
    Dim refRow As Long
    Dim nextItem As Boolean
    Dim ws As Worksheet
    
    Set ws = wsSettlement
    
    wsSettings.Range("E1").Value = wsSettings.Range("E1").Value + 1
    
    
    'Journal 1
    If ws.Range("L44").Value > 0 Then
        Call doJ1
    End If
    
    'Journal 2
    For refRow = 48 To 56
        If ws.Range("T" & refRow).Value <> "Charge to Customer account and raise journal to clear other debtor account" Then
            If (ws.Range("E" & refRow).Value <> 0 And ws.Range("E" & refRow).Value <> "") Or (ws.Range("P" & refRow).Value <> 0 And ws.Range("P" & refRow).Value <> "") Or (ws.Range("J" & refRow).Value <> 0 And ws.Range("J" & refRow).Value <> "") Then Call doFB41(refRow)
        End If
    Next
    
    'Journal 3
    For refRow = 60 To 76
        If ws.Range("T" & refRow).Value <> "Charge to Customer account and raise journal to clear other debtor account" Then
             If (ws.Range("E" & refRow).Value <> 0 And ws.Range("E" & refRow).Value <> "") Or (ws.Range("P" & refRow).Value <> 0 And ws.Range("P" & refRow).Value <> "") Or (ws.Range("J" & refRow).Value <> 0 And ws.Range("J" & refRow).Value <> "") Then Call doFB41(refRow)
        End If
    Next
    
    
    For refRow = 77 To 78
        If ws.Range("P" & refRow).Value < 0 Then
            If ws.Range("T" & refRow).Value <> "Charge to Customer account and raise journal to clear other debtor account" Then
                If (ws.Range("E" & refRow).Value <> 0 And ws.Range("E" & refRow).Value <> "") Or (ws.Range("P" & refRow).Value <> 0 And ws.Range("P" & refRow).Value <> "") Or (ws.Range("J" & refRow).Value <> 0 And ws.Range("J" & refRow).Value <> "") Then Call doFB41(refRow)
            End If
        End If
    Next
    
    
    'Journal 4
    For refRow = 85 To 91
        If ws.Range("T" & refRow).Value <> "Charge to Customer account and raise journal to clear other debtor account" Then
            If (ws.Range("E" & refRow).Value <> 0 And ws.Range("E" & refRow).Value <> "") Or (ws.Range("P" & refRow).Value <> 0 And ws.Range("P" & refRow).Value <> "") Or (ws.Range("J" & refRow).Value <> 0 And ws.Range("J" & refRow).Value <> "") Then Call doFB41(refRow)
        End If
    Next
    
    'Journal 5
    For refRow = 96 To 113
        If ws.Range("T" & refRow).Value <> "Charge to Customer account and raise journal to clear other debtor account" Then
            If (ws.Range("E" & refRow).Value <> 0 And ws.Range("E" & refRow).Value <> "") Or (ws.Range("P" & refRow).Value <> 0 And ws.Range("P" & refRow).Value <> "") Or (ws.Range("J" & refRow).Value <> 0 And ws.Range("J" & refRow).Value <> "") Then Call doFB41(refRow)
        End If
    Next
    
    wbCross.Save
    
End Sub


Sub doJ1()
    Dim refRow As Long
    Dim doneFirst As Boolean
    Dim tempVal As Double
    
    Dim ws As Worksheet
    
    Set ws = wsSettlement
    
    doneFirst = False
    
    With mySess
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
        .findById("wnd[0]/tbar[0]/okcd").Text = "FB41"
        .findById("wnd[0]").sendVKey 0
        
        .findById("wnd[0]/usr/ctxtBKPF-BLDAT").Text = Format(Date, "dd.mm.yyyy") 'document date
        
        .findById("wnd[0]/usr/ctxtBKPF-BUDAT").Text = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "dd.mm.yyyy") 'posting date
        
        .findById("wnd[0]/usr/txtBKPF-MONAT").Text = Month(Date)
        
        .findById("wnd[0]/usr/txtBKPF-BKTXT").Text = "GST_" & Format(DateAdd("m", -1, Date), "MMMYYYY") & "_SETTLEMENT" 'doc header text
        .findById("wnd[0]/usr/txtBKPF-XBLNR").Text = "SA" & wsSettings.Range("E1").Value 'reference
        
        
        For refRow = 12 To 42
            If ws.Range("E" & refRow).Value <> 0 Then
                If doneFirst = False Then
                    .findById("wnd[0]/usr/ctxtBKPF-BUKRS").Text = ws.Range("D" & refRow).Value 'company code
                    .findById("wnd[0]/usr/ctxtBKPF-BLART").Text = "SA" 'wS.Range("R" & refRow).Value 'document type
                    .findById("wnd[0]/usr/ctxtBKPF-WAERS").Text = "AUD" 'currency
                    
                    doneFirst = True
                Else
                    .findById("wnd[0]/usr/ctxtRF05A-NEWBK").Text = ws.Range("D" & refRow).Value 'company code
                End If
                
                
                .findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = ws.Range("B" & refRow).Value 'post key
                .findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = ws.Range("C" & refRow).Value 'account number
                .findById("wnd[0]").sendVKey 0
                
                
                If TestEnv = True Then
                    .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Replace(Abs(Round(ws.Range("E" & refRow).Value, 2)), ".", ",")
                Else
                    .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Abs(Round(ws.Range("E" & refRow).Value, 2))
                End If
                
                .findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = "SEHAL GST PAYMENT TO ATO " & UCase(Format(DateAdd("m", -1, Date), "MMMM YYYY"))
                
                If Not .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL", False) Is Nothing Then
                    Select Case ws.Range("D" & refRow).Value
                    Case "AU01"
                        .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "120378"
                    Case "AU10"
                        .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "136719"
                    Case "AU11"
                        .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "120388"
                    End Select
                Else
                    Select Case ws.Range("D" & refRow).Value
                    Case "AU01"
                        .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "120378"
                    Case "AU10"
                        .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "136719"
                    Case "AU11"
                        .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "120388"
                    End Select
                End If
                
            End If
            
            If wsSettlement.Range("J" & refRow).Value <> 0 Then
                If doneFirst = False Then
                    .findById("wnd[0]/usr/ctxtBKPF-BUKRS").Text = ws.Range("I" & refRow).Value
                    .findById("wnd[0]/usr/ctxtBKPF-BLART").Text = "SA" 'wS.Range("R" & refRow).Value 'document type
                    .findById("wnd[0]/usr/ctxtBKPF-WAERS").Text = "AUD" 'currency
                    doneFirst = True
                Else
                    .findById("wnd[0]/usr/ctxtRF05A-NEWBK").Text = ws.Range("I" & refRow).Value
                End If
                
                .findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = ws.Range("G" & refRow).Value
                .findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = ws.Range("H" & refRow).Value
                .findById("wnd[0]").sendVKey 0
                
                If TestEnv = True Then
                    .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Replace(Abs(Round(ws.Range("J" & refRow).Value, 2)), ".", ",")
                Else
                    .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Abs(Round(ws.Range("J" & refRow).Value, 2))
                End If
                .findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = "SEHAL GST PAYMENT TO ATO " & UCase(Format(DateAdd("m", -1, Date), "MMMM YYYY"))
                
                
                
                
                Select Case ws.Range("I" & refRow).Value
                Case "AU01"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "120378"
                Case "AU10"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "136719"
                Case "AU11"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "120388"
                End Select
                
            End If
        Next
        
        'post lists
        
        'ensure sum equals to L44 value
        
        'if not, trigger prompt
        
        
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/tbar[1]/btn[14]").press
        
        If Right(.findById("wnd[0]/usr/txtRF05A-AZSAL").Text, 1) = "-" Then
            tempVal = -1 * Left(.findById("wnd[0]/usr/txtRF05A-AZSAL").Text, Len(.findById("wnd[0]/usr/txtRF05A-AZSAL").Text) - 1)
        Else
            tempVal = .findById("wnd[0]/usr/txtRF05A-AZSAL").Text * 1
        End If
        
        If ws.Range("L44").Value - Abs(tempVal) < 1 Then
        'If tempVal <> wS.Range("L44").Value Then
            inputError = inputError & vbLf & " - Journal 1: Value not match at " & tempVal
        Else
            'item matched
            
            .findById("wnd[0]/tbar[0]/btn[11]").press
            
            
            If Not .findById("wnd[0]/usr/ctxtBSEG-ZFBDT", False) Is Nothing Then
                ws.Range("X" & refRow).Value = "Cross CoCd No " & .findById("wnd[1]/usr/txtBKPF-BVORG").Text 'cross co code
                ws.Range("X" & refRow).Value = ws.Range("X" & refRow).Value & ": " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/ctxtBKPF-BUKRS[0,0]").Text 'first company code
                ws.Range("X" & refRow).Value = ws.Range("X" & refRow).Value & ", " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/txtBKPF-BELNR[0,8]").Text 'first document no
                ws.Range("X" & refRow).Value = ws.Range("X" & refRow).Value & ", " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/txtBKPF-GJAHR[0,19]").Text 'first year
                
                ws.Range("X" & refRow).Value = ws.Range("X" & refRow).Value & " - " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/ctxtBKPF-BUKRS[1,0]").Text 'second company code
                ws.Range("X" & refRow).Value = ws.Range("X" & refRow).Value & ", " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/txtBKPF-BELNR[1,8]").Text 'second document no
                ws.Range("X" & refRow).Value = ws.Range("X" & refRow).Value & ", " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/txtBKPF-GJAHR[1,19]").Text 'second year
                
                '.findById("wnd[1]").HardCopy 'screenshot
                '.Range("B6").Value = .findById("wnd[0]/sbar").Text
                
                'wS.Range("X" & myRow).PasteSpecial 'paste
                .findById("wnd[1]").sendVKey 0 'close the popup
                
            Else
                ws.Range("X" & refRow).Value = .findById("wnd[0]/sbar").Text
                
            End If
            
        End If
        
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
        If Not .findById("wnd[1]/usr/btnSPOP-OPTION1", False) Is Nothing Then
            .findById("wnd[1]/usr/btnSPOP-OPTION1").press
        End If
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
    End With
    
End Sub



Sub doFB41(myRow As Long)
    Dim myCurr As String
    Dim myPaymentTerm As String
    Dim headerOK As Boolean
    
    Dim ws As Worksheet
    Set ws = wsSettlement
    
    myCurr = "AUD"
    
    If myRow >= 59 And myRow <= 78 Then
    
        If InStr(1, wsSettlement.Range("A" & myRow).Value, "Shell International Trading and Shipping Company Limited", vbTextCompare) <> 0 Then
            myCurr = "USD"
            myPaymentTerm = "Z30"
        ElseIf InStr(1, wsSettlement.Range("A" & myRow).Value, "Shell Tankers (Singapore) Private Limited", vbTextCompare) <> 0 Then
            myCurr = "USD"
            myPaymentTerm = "Z30"
        ElseIf InStr(1, wsSettlement.Range("A" & myRow).Value, "Shell Marine Products US", vbTextCompare) <> 0 Then
            myCurr = "USD"
            myPaymentTerm = "Z00"
        ElseIf InStr(1, wsSettlement.Range("A" & myRow).Value, "Shell Energy Australia (SEAU)", vbTextCompare) <> 0 Then
            myCurr = "USD"
            myPaymentTerm = "Z30"
        ElseIf InStr(1, wsSettlement.Range("A" & myRow).Value, "Shell Markets (Middle East) Limited", vbTextCompare) <> 0 Then
            myCurr = "USD"
            myPaymentTerm = "IR00"
        ElseIf InStr(1, wsSettlement.Range("A" & myRow).Value, "Shell Australia Services Company", vbTextCompare) <> 0 Then
            myPaymentTerm = "Z30"
        ElseIf InStr(1, wsSettlement.Range("A" & myRow).Value, "Trident LNG Shipping Services Pty Ltd", vbTextCompare) <> 0 Then
            myPaymentTerm = "IR00"
        ElseIf InStr(1, wsSettlement.Range("A" & myRow).Value, "Shell Tankers Australia Pty Ltd", vbTextCompare) <> 0 Then
            myPaymentTerm = "IR00"
            
        Else
            myCurr = "AUD"
            myPaymentTerm = ""
        End If
    End If
    
    headerOK = False
    
    If myRow = 112 Or myRow = 113 Then
        myCurr = "AUD"
    End If
    
    
    With mySess
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
        .findById("wnd[0]/tbar[0]/okcd").Text = "FB41"
        .findById("wnd[0]").sendVKey 0
        
        .findById("wnd[0]/usr/ctxtBKPF-BLDAT").Text = Format(Date, "dd.mm.yyyy") 'document date
        
        .findById("wnd[0]/usr/ctxtBKPF-BUDAT").Text = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "dd.mm.yyyy") 'posting date
        
        .findById("wnd[0]/usr/txtBKPF-MONAT").Text = Month(Date) 'period
        
        .findById("wnd[0]/usr/ctxtBKPF-BLART").Text = "SA" 'wS.Range("R" & myRow).Value 'doc type
        
        .findById("wnd[0]/usr/txtBKPF-BKTXT").Text = "GST_" & Format(DateAdd("m", -1, Date), "MMMYYYY") & "_SETTLEMENT" 'doc header text
        .findById("wnd[0]/usr/txtBKPF-XBLNR").Text = "SA" & wsSettings.Range("E1").Value 'reference
        
        .findById("wnd[0]/usr/ctxtBKPF-WAERS").Text = myCurr 'currency
        
        
        If wsSettlement.Range("E" & myRow).Value <> 0 And wsSettlement.Range("E" & myRow).Value <> "" Then  'B to F
            .findById("wnd[0]/usr/ctxtBKPF-BUKRS").Text = ws.Range("D" & myRow).Value 'company code
            .findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = ws.Range("B" & myRow).Value 'Posting Key
            .findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = ws.Range("C" & myRow).Value 'Account No
            .findById("wnd[0]").sendVKey 0
            
            headerOK = True
            'header end
            
            
            .findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = "SETTLEMENT " & Left(Replace(ws.Range("A" & myRow).Value, " ", ""), 36) 'Company name / text
            
            If myCurr = "AUD" Then
                If TestEnv = True Then
                    .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Replace(Abs(Round(ws.Range("E" & myRow).Value, 2)), ".", ",") 'amount for AUD
                Else
                    .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Abs(Round(ws.Range("E" & myRow).Value, 2)) 'amount for AUD
                End If
            Else
                If TestEnv = True Then
                    .findById("wnd[0]/usr/txtBSEG-DMBTR").Text = Replace(Abs(Round(ws.Range("E" & myRow).Value, 2)), ".", ",") 'amount for USD
                Else
                    .findById("wnd[0]/usr/txtBSEG-DMBTR").Text = Abs(Round(ws.Range("E" & myRow).Value, 2)) 'amount for USD
                End If
            End If
            
            If myPaymentTerm <> "" Then
                If Not .findById("wnd[0]/usr/ctxtBSEG-ZTERM", False) Is Nothing Then
                    .findById("wnd[0]/usr/ctxtBSEG-ZTERM").Text = myPaymentTerm 'payment term
                End If
                
                If Not .findById("wnd[0]/usr/ctxtBSEG-ZFBDT", False) Is Nothing Then
                    If .findById("wnd[0]/usr/lblBSEG-ZFBDT").Text <> "Due on" Then
                        If Weekday(DateSerial(Year(Date), Month(Date), 21)) = 1 Then
                            .findById("wnd[0]/usr/ctxtBSEG-ZFBDT").Text = Format(DateSerial(Year(Date), Month(Date), 22), "dd.mm.yyyy")
                        ElseIf Weekday(DateSerial(Year(Date), Month(Date), 21)) = 7 Then
                            .findById("wnd[0]/usr/ctxtBSEG-ZFBDT").Text = Format(DateSerial(Year(Date), Month(Date), 23), "dd.mm.yyyy")
                        Else
                            .findById("wnd[0]/usr/ctxtBSEG-ZFBDT").Text = Format(DateSerial(Year(Date), Month(Date), 21), "dd.mm.yyyy")
                        End If
                    End If
                End If
            End If
            
            If Not .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL", False) Is Nothing Then
                Select Case ws.Range("D" & myRow).Value
                Case "AU01"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "120378" 'cost center
                Case "AU10"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "136719" 'cost center
                Case "AU11"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "120388" 'cost center
                End Select
                
            ElseIf Not .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL", False) Is Nothing Then
                Select Case ws.Range("D" & myRow).Value
                Case "AU01"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "120378" 'cost center
                Case "AU10"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "136719" 'cost center
                Case "AU11"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "120388" 'cost center
                End Select
            End If
            
        End If
        
        
        
        
        
        If wsSettlement.Range("J" & myRow).Value <> 0 And wsSettlement.Range("J" & myRow).Value <> "" Then 'G to K only
            
            If headerOK = False Then
                'need to set header
                .findById("wnd[0]/usr/ctxtBKPF-BUKRS").Text = ws.Range("I" & myRow).Value 'company code
                headerOK = True
                
            Else
                .findById("wnd[0]/usr/ctxtRF05A-NEWBK").Text = ws.Range("I" & myRow).Value 'company code
                
            End If
            
            .findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = ws.Range("G" & myRow).Value 'Posting Key
            .findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = ws.Range("H" & myRow).Value 'Account No
            .findById("wnd[0]").sendVKey 0
            
            
            
            .findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = "SETTLEMENT " & Left(Replace(ws.Range("A" & myRow).Value, " ", ""), 36) 'Company name / text
            
            
            If myCurr = "AUD" Then
                If TestEnv = True Then
                    .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Replace(Abs(Round(ws.Range("J" & myRow).Value, 2)), ".", ",") 'amount for AUD
                Else
                    .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Abs(Round(ws.Range("J" & myRow).Value, 2)) 'amount for AUD
                End If
            Else
                If TestEnv = True Then
                    .findById("wnd[0]/usr/txtBSEG-DMBTR").Text = Replace(Abs(Round(ws.Range("J" & myRow).Value, 2)), ".", ",") 'amount for USD
                Else
                    .findById("wnd[0]/usr/txtBSEG-DMBTR").Text = Abs(Round(ws.Range("J" & myRow).Value, 2)) 'amount for USD
                End If
            End If
            
            
            If myPaymentTerm <> "" Then
                If Not .findById("wnd[0]/usr/ctxtBSEG-ZTERM", False) Is Nothing Then
                    .findById("wnd[0]/usr/ctxtBSEG-ZTERM").Text = myPaymentTerm 'payment term
                End If
                
                If Not .findById("wnd[0]/usr/ctxtBSEG-ZFBDT", False) Is Nothing Then
                    If .findById("wnd[0]/usr/lblBSEG-ZFBDT").Text <> "Due on" Then
                        If Weekday(DateSerial(Year(Date), Month(Date), 21)) = 1 Then
                            .findById("wnd[0]/usr/ctxtBSEG-ZFBDT").Text = Format(DateSerial(Year(Date), Month(Date), 22), "dd.mm.yyyy")
                        ElseIf Weekday(DateSerial(Year(Date), Month(Date), 21)) = 7 Then
                            .findById("wnd[0]/usr/ctxtBSEG-ZFBDT").Text = Format(DateSerial(Year(Date), Month(Date), 23), "dd.mm.yyyy")
                        Else
                            .findById("wnd[0]/usr/ctxtBSEG-ZFBDT").Text = Format(DateSerial(Year(Date), Month(Date), 21), "dd.mm.yyyy")
                        End If
                    End If
                End If
            End If
            
            If Not .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL", False) Is Nothing Then
                Select Case ws.Range("I" & myRow).Value
                Case "AU01"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "120378" 'cost center
                Case "AU10"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "136719" 'cost center
                Case "AU11"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "120388" 'cost center
                End Select
                
            ElseIf Not .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL", False) Is Nothing Then
                Select Case ws.Range("I" & myRow).Value
                Case "AU01"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "120378" 'cost center
                Case "AU10"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "136719" 'cost center
                Case "AU11"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "120388" 'cost center
                End Select
            End If
            
        End If
        
        
        
        
        
        If wsSettlement.Range("P" & myRow).Value <> 0 And wsSettlement.Range("P" & myRow).Value <> "" Then 'M to Q only
            If headerOK = False Then
                'need to set header
                .findById("wnd[0]/usr/ctxtBKPF-BUKRS").Text = ws.Range("O" & myRow).Value 'company code
                headerOK = True
                
            Else
                .findById("wnd[0]/usr/ctxtRF05A-NEWBK").Text = ws.Range("O" & myRow).Value 'company code
                
            End If
            
            .findById("wnd[0]/usr/ctxtRF05A-NEWBS").Text = ws.Range("M" & myRow).Value 'Posting Key
            .findById("wnd[0]/usr/ctxtRF05A-NEWKO").Text = ws.Range("N" & myRow).Value 'Account No
            .findById("wnd[0]").sendVKey 0
            
            
            
            .findById("wnd[0]/usr/ctxtBSEG-SGTXT").Text = "SETTLEMENT " & Left(Replace(ws.Range("A" & myRow).Value, " ", ""), 36) 'Company name / text
            
            
            If myCurr = "AUD" Then
                If TestEnv = True Then
                    .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Replace(Abs(Round(ws.Range("P" & myRow).Value, 2)), ".", ",") 'amount for AUD
                Else
                    .findById("wnd[0]/usr/txtBSEG-WRBTR").Text = Abs(Round(ws.Range("P" & myRow).Value, 2)) 'amount for AUD
                End If
            Else
                If TestEnv = True Then
                    .findById("wnd[0]/usr/txtBSEG-DMBTR").Text = Replace(Abs(Round(ws.Range("P" & myRow).Value, 2)), ".", ",") 'amount for USD
                Else
                    .findById("wnd[0]/usr/txtBSEG-DMBTR").Text = Abs(Round(ws.Range("P" & myRow).Value, 2)) 'amount for USD
                End If
            End If
            
            
            If myPaymentTerm <> "" Then
                If Not .findById("wnd[0]/usr/ctxtBSEG-ZTERM", False) Is Nothing Then
                    .findById("wnd[0]/usr/ctxtBSEG-ZTERM").Text = myPaymentTerm 'payment term
                End If
                
                If Not .findById("wnd[0]/usr/ctxtBSEG-ZFBDT", False) Is Nothing Then
                    If .findById("wnd[0]/usr/lblBSEG-ZFBDT").Text <> "Due on" Then
                        If Weekday(DateSerial(Year(Date), Month(Date), 21)) = 1 Then
                            .findById("wnd[0]/usr/ctxtBSEG-ZFBDT").Text = Format(DateSerial(Year(Date), Month(Date), 22), "dd.mm.yyyy")
                        ElseIf Weekday(DateSerial(Year(Date), Month(Date), 21)) = 7 Then
                            .findById("wnd[0]/usr/ctxtBSEG-ZFBDT").Text = Format(DateSerial(Year(Date), Month(Date), 23), "dd.mm.yyyy")
                        Else
                            .findById("wnd[0]/usr/ctxtBSEG-ZFBDT").Text = Format(DateSerial(Year(Date), Month(Date), 21), "dd.mm.yyyy")
                        End If
                    End If
                End If
            End If
            
            If Not .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL", False) Is Nothing Then
                Select Case ws.Range("O" & myRow).Value
                Case "AU01"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "120378" 'cost center
                Case "AU10"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "136719" 'cost center
                Case "AU11"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1006/ctxtCOBL-KOSTL").Text = "120388" 'cost center
                End Select
            ElseIf Not .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL", False) Is Nothing Then
                Select Case ws.Range("O" & myRow).Value
                Case "AU01"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "120378" 'cost center
                Case "AU10"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "136719" 'cost center
                Case "AU11"
                    .findById("wnd[0]/usr/subBLOCK:SAPLKACB:1004/ctxtCOBL-KOSTL").Text = "120388" 'cost center
                End Select
            End If
            
            
        End If
        
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/tbar[1]/btn[14]").press
        
        If .findById("wnd[0]/sbar").Text = "Terms of payment changed; Check" Then
            .findById("wnd[0]").sendVKey 0
        End If
        'if term of enter something something then i need to press enter again
        
        If Not .findById("wnd[1]/usr/lbl[0,2]", False) Is Nothing Then
            .findById("wnd[1]/usr/lbl[0,2]").SetFocus
            .findById("wnd[1]/usr/lbl[0,2]").caretPosition = 2
            .findById("wnd[1]").sendVKey 2
        End If
        
        If Right(.findById("wnd[0]/usr/txtRF05A-AZSAL").Text, 1) = "-" Then
            inputError = inputError & vbLf & " - Row " & myRow & " items does not sum to 0"
        ElseIf .findById("wnd[0]/usr/txtRF05A-AZSAL").Text <> 0 Then
            inputError = inputError & vbLf & " - Row " & myRow & " items does not equal 0"
        Else
            'item matched
            
            .findById("wnd[0]/tbar[0]/btn[11]").press
            '.findById("wnd[0]").HardCopy
            'wS.Range("X" & myRow).PasteSpecial
            
            If Not .findById("wnd[1]/usr/txtBKPF-BVORG", False) Is Nothing Then
                ws.Range("X" & myRow).Value = "Cross CoCd No " & .findById("wnd[1]/usr/txtBKPF-BVORG").Text 'cross co code
                ws.Range("X" & myRow).Value = ws.Range("X" & myRow).Value & ": " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/ctxtBKPF-BUKRS[0,0]").Text 'first company code
                ws.Range("X" & myRow).Value = ws.Range("X" & myRow).Value & ", " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/txtBKPF-BELNR[0,8]").Text 'first document no
                ws.Range("X" & myRow).Value = ws.Range("X" & myRow).Value & ", " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/txtBKPF-GJAHR[0,19]").Text 'first year
                
                ws.Range("X" & myRow).Value = ws.Range("X" & myRow).Value & " - " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/ctxtBKPF-BUKRS[1,0]").Text 'second company code
                ws.Range("X" & myRow).Value = ws.Range("X" & myRow).Value & ", " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/txtBKPF-BELNR[1,8]").Text 'second document no
                ws.Range("X" & myRow).Value = ws.Range("X" & myRow).Value & ", " & .findById("wnd[1]/usr/sub:SAPMF05A:0607/txtBKPF-GJAHR[1,19]").Text 'second year
                
                '.findById("wnd[1]").HardCopy 'screenshot
                '.Range("B6").Value = .findById("wnd[0]/sbar").Text
                
                'wS.Range("X" & myRow).PasteSpecial 'paste
                .findById("wnd[1]").sendVKey 0 'close the popup
                
            Else
                ws.Range("X" & myRow).Value = .findById("wnd[0]/sbar").Text
                
            End If
            
            
        End If
        
        .findById("wnd[0]/tbar[0]/btn[3]").press
        If Not .findById("wnd[1]", False) Is Nothing Then
            .findById("wnd[1]/usr/btnSPOP-OPTION1").press
        End If
        
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
    End With
    
    
End Sub












