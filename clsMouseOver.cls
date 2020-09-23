VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Event PollingStarted()
Public Event PollingStopped()
Public Event MouseOver(sControlName As String)
Public Event MouseExit(sControlName As String)

Public hLastMouseOver As Long
Public sLastMouseOver As String

Private bExitRequired As Boolean

Public Sub StopTimer()

    ' Setting this flag to true causes the StartPolling routine to terminate.
    bExitRequired = True

End Sub

Public Sub StartPolling(frmPassedForm As Form, lInterval As Long)
On Error Resume Next

    ' Every <lInterval> milliseconds we will test the cursor position
    ' and report any mouseover/mouseexit events that we detect.

    Dim dblStart As Double, dblLastPoll As Double
    Dim MousePos As POINTAPI, ControlIdx As Integer, hWindow As Long, sControlName As String
    
    dblStart = Timer
    bExitRequired = False
    hLastMouseOver = 0
    
    ' Notify the calling program that we have started polling for mouseover events.
    RaiseEvent PollingStarted
    
    ' This loop will repeat until routine "StopPolling" sets bExitRequired to true.
    While Not bExitRequired
        
        ' Wait until an amount of time has passed that is not less than the passed milliseconds.
        dblLastPoll = Timer
        While Int(Timer * 1000) < Int(dblLastPoll * 1000) + lInterval
            DoEvents
        Wend
            
        ' Find the windows handle (hWindow) of the control underneath the mouse pointer.
        Call GetCursorPos(MousePos)
        hWindow = WindowFromPoint(MousePos.X, MousePos.Y)
            
        ' If MouseExit or MouseOver events are needed then generate them.
        If hWindow <> hLastMouseOver Then
                
            ' Report the MouseExit Event.
            If hLastMouseOver <> 0 Then RaiseEvent MouseExit(sControlName)
                
            ' Raise a MouseOver event if the control was found.
            sControlName = FindControlName(frmPassedForm, hWindow)
            If sControlName = "None" Then
                hLastMouseOver = 0
            Else
                RaiseEvent MouseOver(sControlName)
                hLastMouseOver = hWindow
            End If
                
        End If
                    
    Wend
    
    ' Notify the calling program that we have stopped polling for mouseover events.
    RaiseEvent PollingStopped
    
End Sub

Private Function FindControlName(frmPassedForm As Form, hWindow As Long) As String
On Error Resume Next

    ' This function will search the passed form for a control with
    ' a handle (hwnd) that matches the passed handle.
    
    Dim iControl As Integer, bFound As Boolean, sName As String
            
    ' Search for a control on the form with the passed windows handle.
    bFound = False: sName = "None"
    With frmPassedForm
        If .Controls.Count > 0 Then
            For iControl = 0 To .Controls.Count - 1
                bFound = (hWindow = .Controls(iControl).hWnd)
                If bFound Then sName = .Controls(iControl).Name: Exit For
            Next iControl
        End If
    End With

    ' Return the name of the found control (or "None" if we could not find it).
    FindControlName = sName

End Function
