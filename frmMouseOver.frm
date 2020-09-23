VERSION 5.00
Begin VB.Form frmMouseOverTest 
   Caption         =   " Detect Form MouseOver and MouseExit events"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1545
      Left            =   3795
      Picture         =   "frmMouseOver.frx":0000
      ScaleHeight     =   1485
      ScaleWidth      =   1485
      TabIndex        =   5
      Top             =   4770
      Width           =   1545
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   270
      Left            =   3795
      TabIndex        =   4
      Top             =   4185
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3795
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3285
      Width           =   2400
   End
   Begin VB.TextBox txtEventLog 
      Height          =   7155
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   255
      Width           =   3180
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Polling"
      Enabled         =   0   'False
      Height          =   645
      Left            =   3795
      TabIndex        =   1
      Top             =   6720
      Width           =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Polling"
      Height          =   645
      Left            =   3795
      TabIndex        =   0
      Top             =   270
      Width           =   2400
   End
   Begin VB.Label Label2 
      Caption         =   "At presentwe only detect mouseover and mouseexit from certain types of controls, but you should find it usefull all the same."
      Height          =   870
      Left            =   3795
      TabIndex        =   7
      Top             =   1980
      Width           =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Press the start polling button then move the mouse arround this form."
      Height          =   705
      Left            =   3795
      TabIndex        =   6
      Top             =   1200
      Width           =   2400
   End
End
Attribute VB_Name = "frmMouseOverTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This project will demonstrate a good method of detecting
' MouseOver and MouseExit Events.

Private WithEvents FormMouseEvents As clsMouseOver
Attribute FormMouseEvents.VB_VarHelpID = -1

Private Sub Form_Load()
    
    ' This instruction is needed here to enable the
    ' class module that does the clever stuff for us.
    Set FormMouseEvents = New clsMouseOver

End Sub

Private Sub Command1_Click()

    txtEventLog.Text = ""

    ' The instruction below will start a routine running which continuously
    ' test the control under the mousepointer, and generate events as necessary.

    Call FormMouseEvents.StartPolling(Me, 100)
    
    ' Explanation of parameters used;
    '
    '       First Parameter (in this example Me) = Name of form to check for events
    '       Second Parameter (in this example 100) = Polling frequency in milliseconds

End Sub

Private Sub Command2_Click()

    ' This call will stop our program polling for Mouse Over/Exit events.
    Call FormMouseEvents.StopTimer

End Sub

Private Sub FormMouseEvents_MouseOver(sControlName As String)

    ' For demonstration purposes I have made the MouseOver event notify the user
    ' by putting a new message in the forms textbox.
    With txtEventLog
        .Text = txtEventLog.Text & "MouseOver: " & sControlName & vbCrLf
        .SelStart = Len(.Text)
    End With

    ' In your programs you might try using a Select Case instruction. eg;
    '
    '       Select Case sControlName
    '           Case is = Control1
    '               Call ActionToBeTakenOnMouseoverControl1
    '           Case is = Control2
    '               Call ActionToBeTakenOnMouseoverControl2
    '       End Select
    
End Sub

Private Sub FormMouseEvents_MouseExit(sControlName As String)
    
    ' For demonstration purposes I have made the MouseExit event notify the user
    ' by putting a new message in the forms textbox.
    With txtEventLog
        .Text = txtEventLog.Text & "MouseExit: " & sControlName & vbCrLf
        .SelStart = Len(.Text)
    End With
    
    ' In your programs you might try using a Select Case instruction. eg;
    '
    '       Select Case sControlName
    '           Case is = Control1
    '               Call ActionToBeTakenOnMouseExitControl1
    '           Case is = Control2
    '               Call ActionToBeTakenOnMouseExitControl2
    '       End Select

End Sub

Private Sub FormMouseEvents_PollingStarted()
    
    ' Update Command Buttons.
    Command1.Enabled = False
    Command2.Enabled = True

End Sub

Private Sub FormMouseEvents_PollingStopped()
    
    ' Update Command Buttons.
    Command1.Enabled = True
    Command2.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' We should not allow form exit until the class has been properly terminated.
    If Command2.Enabled = True Then
        Cancel = True
        Call FormMouseEvents.StopTimer
    End If
    
End Sub
