VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf2 
   Caption         =   "Australia SAPL & SEHAL Automation: Email Template"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915.001
   OleObjectBlob   =   "4NAS_UF3_EmailFormat.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
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



Private OutApp As Object
Private OutMail As Object


Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub BtnRun_Click()
    Dim xCount As Long

    If Me.TextBox1.Value = "" Then
    ElseIf Not IsNumeric(Me.TextBox1.Value) Then
        MsgBox ("Please ensure value you entered is numeric")
        Exit Sub
    ElseIf Me.TextBox1.Value < 0 Or Me.TextBox1.Value > 10 Then
        MsgBox ("Please ensure value you entered is positive whole number from 0 to 10.")
        MsgBox (Round(Me.TextBox1.Value, 0) & vbLf & Round(Me.TextBox1.Value, 0) = Me.TextBox1.Value)
        Exit Sub
    End If
    
    If Me.TextBox2.Value = "" Then
    ElseIf Not IsNumeric(Me.TextBox2.Value) Then
        MsgBox ("Please ensure value you entered is numeric")
        Exit Sub
    ElseIf Me.TextBox2.Value < 0 Or Me.TextBox2.Value > 10 Then
        MsgBox ("Please ensure value you entered is positive whole number from 0 to 10.")
        Exit Sub
    End If
    
    If Me.TextBox3.Value = "" Then
    ElseIf Not IsNumeric(Me.TextBox3.Value) Then
        MsgBox ("Please ensure value you entered is numeric")
        Exit Sub
    ElseIf Me.TextBox3.Value < 0 Or Me.TextBox3.Value > 10 Then
        MsgBox ("Please ensure value you entered is positive whole number from 0 to 10.")
        Exit Sub
    End If
    
    If Me.TextBox4.Value = "" Then
    ElseIf Not IsNumeric(Me.TextBox4.Value) Then
        MsgBox ("Please ensure value you entered is numeric")
        Exit Sub
    ElseIf Me.TextBox4.Value < 0 Or Me.TextBox4.Value > 10 Then
        MsgBox ("Please ensure value you entered is positive whole number from 0 to 10.")
        Exit Sub
    End If
    
    If (Me.TextBox1.Value = "" Or Me.TextBox1.Value = 0) And (Me.TextBox2.Value = "" Or Me.TextBox2.Value = 0) And (Me.TextBox3.Value = "" Or Me.TextBox3.Value = 0) And (Me.TextBox4.Value = "" Or Me.TextBox4.Value = 0) Then
        MsgBox ("Please enter a number in any of the fields")
        Exit Sub
    End If
    
    
    'call pause
    Call RunPauseAll
    
    Set OutApp = CreateObject("Outlook.Application")
    
    If Not (Me.TextBox1.Value = "" Or Me.TextBox1.Value = 0) Then
        For xCount = 1 To Me.TextBox1.Value
            'call create email 1
            Call generate_Email1
        Next
    End If
    
    
    If Not (Me.TextBox2.Value = "" Or Me.TextBox2.Value = 0) Then
        For xCount = 1 To Me.TextBox2.Value
            'call create email 2
            Call generate_Email2
        Next
    End If
    
    
    If Not (Me.TextBox3.Value = "" Or Me.TextBox3.Value = 0) Then
        For xCount = 1 To Me.TextBox3.Value
            'call create email 3
            Call generate_Email3
        Next
    End If
    
    
    If Not (Me.TextBox4.Value = "" Or Me.TextBox4.Value = 0) Then
        For xCount = 1 To Me.TextBox4.Value
            'call create email 4
            Call generate_Email4
        Next
    End If
    
    Call RunActivateAll
    
    Unload Me
    
End Sub


Sub generate_Email1()
    Dim myBod As String
    
    myBod = "Hi,<p>Good day.<p>"
    myBod = myBod & "MIT is in the middle of the GST submission for the current month. Based on the transaction, we notice that the tax code/value posted in SAP is not align with supporting documents attached in SAP.<br>"
    myBod = myBod & "Please advise on this by/on WD7 so that we can complete our tax submissions on time.<br><br>"
    myBod = myBod & "Thank you"
    
    
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .Subject = "Tax Code Posted in SAP"
        .HTMLBody = myBod
        .Display
    End With
    
    
End Sub

Sub generate_Email2()
    
    Dim myBod As String
    
    
    myBod = "Hi,<p>Good day.<p>"
    myBod = myBod & "MIT is in the midst of GST submission for the current month. During the tax preparation, we notice that below transaction should have a tax value applied based on the basis of the chosen tax code.<br>"
    myBod = myBod & "However, it shows a nil value. Is it a Split Tax? Please advise on this by/on WD7 so that we can complete our tax submissions on time.<br><br>"
    myBod = myBod & "Thank you"
    
    
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .Subject = "Application of the Tax Code Posted in SAP"
        .HTMLBody = myBod
        .Display
    End With
    
    
    
    
End Sub

Sub generate_Email3()
    Dim myBod As String
    
    myBod = "Hi,<p>Good day.<p>"
    myBod = myBod & "MIT is in the middle of the GST submission for the current month. Based on the transaction, please advise on the original Invoice details to match the value of Credit Memo/Note raised.<br>"
    myBod = myBod & "And also kindly mention whether this is a full or partial Credit Memo/Note. Please advise on this by/on WD7 so that we can complete our tax submissions on time.<br><br>"
    myBod = myBod & "Thank you"
    
    
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .Subject = "Query on Credit Memo/ Note"
        .HTMLBody = myBod
        .Display
    End With
    
    
End Sub


Sub generate_Email4()
    Dim myBod As String
    
    myBod = "Hi,<p>Good day.<p>"
    myBod = myBod & "MIT in the middle of the GST submission for the current month. On the basis of the transaction, we would like to ensure that the tax applied in SAP is correct.<br>"
    myBod = myBod & "Therefore, appreciate your advice on shipping details to ensure that the tax code applied is valid.<br><br>"
    
    myBod = myBod & "Please advise on this by/on WD7 so that we can complete our tax submissions on time.<br><br>"
    myBod = myBod & "Thank you"
    
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .Subject = "Shipment Details on Documents Posted in SAP"
        .HTMLBody = myBod
        .Display
    End With
    
End Sub
