<div align="center">

## MouseWheel


</div>

### Description

Quick and dirty code to get MouseWheel event information without any ocx, just a few constants and lines of code...
 
### More Info
 
' 1 - Create a new project

' 2 - Add 3 PictureBox (Picture1, Picture2, Picture3)

' 3 - Add a TextBox (Text1)

' 4 - Paste code

' Run

' over PictureBox and watch cursors

' Wheeling moves vertical cursor

' Shift Key multiplies 10 times wheel action

' Ctrl Key drives action to horizontal cursor

'

' Over 'Spin'TextBox

' Click to enable and then 'Wheel' and Watch

' Shift Key multiplies 10 times wheel action

' Ctrl key multiplies 100 times wheel action

No side effect if used as in sample project...

If U encounter any, or feel I may have some trouble, please let me know


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jean\-Pierre Perriere](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jean-pierre-perriere.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jean-pierre-perriere-mousewheel__1-5750/archive/master.zip)





### Source Code

```
' Assume:
' 1 - Create a new project
' 2 - Add 3 PictureBox (Picture1, Picture2, Picture3)
' 3 - Add a TextBox (keep name Text1)
'
' aver PictureBox
' Shift Key multiplies 10 times wheel action
' Ctrl Key drives action to horizontal scroll
'
' Over 'Spin'TextBox
' Shift Key multiplies 10 times wheel action
' Ctrl key multiplies 100 times wheel action
Option Explicit
'=================================
' Constante de GetSystemMetrics
'=================================
Const SM_MOUSEWHEELPRESENT As Long = 75 '  Vrai si molette
Private Declare Function GetSystemMetrics Lib "user32" ( _
  ByVal nIndex As Long _
) As Long
'=================================
' Constantes de messages
'=================================
Const WM_MOUSEWHEEL As Integer = &H20A '  action sur la molette
Const WM_MOUSEHOVER As Integer = &H2A1
Const WM_MOUSELEAVE As Integer = &H2A3
Const WM_KEYDOWN As Integer = &H100
Const WM_KEYUP As Integer = &H101
Const WM_CHAR As Integer = &H102
'=================================
' Constants Mask for MouseWheelKey
'=================================
Const MK_LBUTTON As Integer = &H1
Const MK_RBUTTON As Integer = &H2
Const MK_MBUTTON As Integer = &H10
Const MK_SHIFT As Integer = &H4
Const MK_CONTROL As Integer = &H8
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type MSG
  hwnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" ( _
  lpMsg As MSG, _
  ByVal hwnd As Long, _
  ByVal wMsgFilterMin As Long, _
  ByVal wMsgFilterMax As Long _
) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" ( _
  lpMsg As MSG _
) As Long
Private Declare Function TranslateMessage Lib "user32" ( _
  lpMsg As MSG _
) As Long
'==================================================
'  Fonction used for mouse tracking (Win 98)
'==================================================
Private Declare Function TRACKMOUSEEVENT Lib "user32" Alias "TrackMouseEvent" ( _
  lpEventTrack As TRACKMOUSEEVENT _
) As Boolean
Private Type TRACKMOUSEEVENT
  cbSize As Long
  dwFlags As Long
  hwndTrack As Long
  dwHoverTime As Long
End Type
  '======================================
  ' Constants for TrackMouseEvent type
  '======================================
  Const TME_HOVER As Long = &H1
  Const TME_LEAVE As Long = &H2
  Const TME_QUERY As Long = &H40000000
  Const TME_CANCEL As Long = &H80000000
  Const HOVER_DEFAULT As Long = &HFFFFFFFF
'==================================================
'  Fonction used for mouse tracking (old school)
'==================================================
Private Declare Function GetCursorPos Lib "user32" ( _
  lpPoint As POINTAPI _
) As Long
Private Declare Function WindowFromPoint Lib "user32" ( _
  ByVal X As Long, _
  ByVal Y As Long _
) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
  ByVal hwnd As Long, _
  ByVal lpClassName As String, _
  ByVal nMaxCount As Long _
) As Long
'=================================
' Variables for wheel tracking
'=================================
Dim m_blnWheelPresent As Boolean  ' true if mouse Wheel present
Dim m_blnWheelTracking As Boolean  ' true while pumping messages
Dim m_blnKeepSpinnig As Boolean    ' true = mouse still active away from source
Dim m_tMSG As MSG          ' messages structure
'==================================
' Constants for sample application
'==================================
Const m_sCurOffset As Single = 112   ' middle of cursor picture is 7 pixels away from side
Const m_WheelForward As Long = -1    ' Wheeling 'Down' like to walk down a window = increase value
Const m_WheelBackward As Long = 1    ' Wheeling 'Down'              = decrease value
'==================================
' Variables for sample application
'==================================
  'picture section
  Dim m_sScaleMultiplier_H As Single
  Dim m_sScaleMax_H As Single
  Dim m_sScaleMin_H As Single
  Dim m_sScaleValue_H As Single
  Dim m_sScaleMultiplier_V As Single
  Dim m_sScaleMax_V As Single
  Dim m_sScaleMin_V As Single
  Dim m_sScaleValue_V As Single
  'text section
  Dim m_lWalkWay As Long     ' Will be set to your choice m_WheelForward or m_WheelForward in initialise proc
  Dim m_lMutiplier_Small As Long
  Dim m_lMutiplier_Large As Long
  Dim m_lSampleValue As Long
Sub WatchForWheel(hClient As Long, Optional blnWheelAround As Boolean)
Dim i As Integer
Dim lResult As Long
Dim bResult As Boolean
Dim tTrackMouse As TRACKMOUSEEVENT
Dim tMouseCords As POINTAPI
Dim lX As Long, lY As Long '  mouse coordinates
Dim lCurrentHwnd As Long  '
Dim iDirection As Integer
Dim iKeys As Integer
If IsMissing(blnWheelAround) Then
  m_blnKeepSpinnig = False
Else
  m_blnKeepSpinnig = blnWheelAround
End If
m_blnWheelTracking = True
'With tTrackMouse
'  .cbSize =         ' sizeof tTrackMouse : how to calculate that ?
'  .dwFlags = TME_LEAVE
'  .dwHoverTime = HOVER_DEFAULT
'  .hwndTrack = hClient
'End With
'bResult = TRACKMOUSEEVENT(tTrackMouse)
  '********************************************************
  ' Message pump:
  ' gets all messages and checks for MouseWheel event
  '********************************************************
  Do While m_blnWheelTracking
    lResult = GetCursorPos(tMouseCords) ' Get current mouse location
      lX = tMouseCords.X
      lY = tMouseCords.Y
    lCurrentHwnd = WindowFromPoint(lX, lY) ' get the window under the mouse from mouse coordinates
    If lCurrentHwnd <> hClient Then
      If m_blnKeepSpinnig = False Then   ' Don't stop if true
        m_blnWheelTracking = False   ' We are off the client window
        Exit Do             ' so we stop tracking
      End If
    End If
    lResult = GetMessage(m_tMSG, Me.hwnd, 0, 0)
    lResult = TranslateMessage(m_tMSG)
    '=======================================
    ' on renvoie le message dans le circuit
    ' pour la gestion des événements
    '=======================================
    lResult = DispatchMessage(m_tMSG)
    DoEvents
    Select Case m_tMSG.message
      Case WM_MOUSEWHEEL
        '===============================================================
        ' Message is 'Wheel Rolling'
        '===============================================================
        Call WheelAction(hClient, m_tMSG.wParam)
      Case WM_MOUSELEAVE
        '======================================================
        ' Mouse Leave generated by TRACKMOUSEEVENT
        ' when mouse leaves client if TRACKMOUSEEVENT structure
        ' well filled (not here...)
        '======================================================
        m_blnWheelTracking = False
    End Select
    DoEvents
  Loop
End Sub
Sub WheelAction(hClient As Long, wParam)
Dim iKey As Integer
Dim iDir As Integer
'===============================================================
' We get wheel direction (left half of wParams)
' and Keys pressed while 'wheeling' (right half of wParams)
'===============================================================
iKey = CInt("&H" & (Right(Hex(wParam), 4)))
iDir = Sgn(wParam \ 32767)
'========================================================
' Generic code to get mouse buttons and keys information
'========================================================
'If iKey And MK_LBUTTON Then  - Left Button code -
'If iKey And MK_RBUTTON Then  - Right Button code -
'If iKey And MK_MBUTTON Then  - Middle Button code -
'If iKey And MK_SHIFT Then   - ShiftKey code -
'If iKey And MK_CONTROL Then  - ControlKey code -
Select Case hClient
  Case Picture1.hwnd
    '========================================================
    ' CtrlKey used to change scroll to be modified:
    ' on => Scroll_H off => Scroll_V
    '========================================================
    If iKey And MK_CONTROL Then
      '============================
      ' ShiftKey used as multiplier
      '============================
      If iKey And MK_SHIFT Then
        m_sScaleValue_H = m_sScaleValue_H + iDir * m_sScaleMultiplier_H
      Else
         m_sScaleValue_H = m_sScaleValue_H + iDir
      End If
      '============================
      ' Check limits
      '============================
      If m_sScaleValue_H <= m_sScaleMin_H Then m_sScaleValue_H = m_sScaleMin_H
      If m_sScaleValue_H >= m_sScaleMax_H Then m_sScaleValue_H = m_sScaleMax_H
      Picture3.Left = Picture1.Left + Picture1.Width - m_sCurOffset - m_sScaleValue_H * (Picture1.Width / m_sScaleMax_H)
    Else
      '============================
      ' CtrlKey used as multiplier
      '============================
      If iKey And MK_SHIFT Then
        m_sScaleValue_V = m_sScaleValue_V + iDir * m_sScaleMultiplier_V
      Else
         m_sScaleValue_V = m_sScaleValue_V + iDir
      End If
      '============================
      ' Check limits
      '============================
      If m_sScaleValue_V <= m_sScaleMin_V Then m_sScaleValue_V = m_sScaleMin_V
      If m_sScaleValue_V >= m_sScaleMax_V Then m_sScaleValue_V = m_sScaleMax_V
      Picture2.Top = Picture1.Top + Picture1.Height - m_sCurOffset - m_sScaleValue_V * (Picture1.Height / m_sScaleMax_V)
    End If
  Case Text1.hwnd
    '================================
    ' CtrlKey used as 100x multiplier
    ' ShiftKey used as 10x multiplier
    '================================
    If iKey And MK_CONTROL Then
      m_lSampleValue = m_lSampleValue + m_lWalkWay * iDir * m_lMutiplier_Large
    ElseIf iKey And MK_SHIFT Then
      m_lSampleValue = m_lSampleValue + m_lWalkWay * iDir * m_lMutiplier_Small
    Else
      m_lSampleValue = m_lSampleValue + m_lWalkWay * iDir
    End If
    Text1 = Trim(Str(m_lSampleValue))
'  Case Your_Next_Hwnd
    '
    '
'  Case Your_Last_Hwnd
End Select
End Sub
Sub initialize()
Dim i As Integer
'=================================
' Mouse section : check for wheel
'=================================
  m_blnWheelPresent = GetSystemMetrics(SM_MOUSEWHEELPRESENT)
'********************************************
' Begin Custom section
'
'********************************************
'================================================
' Drawing cursor shapes in picture2 and picture3
'================================================
Picture1.Move 240, 240, 3015, 1935
Picture1.ScaleMode = vbPixels
Picture1.AutoRedraw = True
For i = 255 To 0 Step -1
  Picture1.Line ((Picture1.ScaleWidth / 255) * i, (Picture1.ScaleHeight / 255) * i)- _
         (Picture1.ScaleWidth, Picture1.ScaleHeight), _
          RGB(i, i / 2, i / 2), B
Next i
With Picture2        '  Right cursor
  .AutoRedraw = True
  .Appearance = 0
  .BorderStyle = 0
  .BackColor = &H8000000F
  .ScaleMode = vbPixels
  .Height = 225
  .Left = Picture1.Left + Picture1.Width
  .Width = 225
End With
With Picture3        '  Bottom cursor
  .AutoRedraw = True
  .Appearance = 0
  .BorderStyle = 0
  .BackColor = &H8000000F
  .ScaleMode = vbPixels
  .Height = 225
  .Top = Picture1.Top + Picture1.Height
  .Width = 225
End With
For i = 0 To 7
  Picture2.Line (i, 7 - i)-(i, 7 + i)
  Picture3.Line (7 - i, i)-(7 + i, i)
Next i
'================================
' Picture1 PseudoScrolls section
'================================
  m_sScaleMultiplier_H = 10
  m_sScaleMax_H = 150
  m_sScaleMin_H = 0
  m_sScaleValue_H = m_sScaleMax_H / 2
  m_sScaleMultiplier_V = 10
  m_sScaleMax_V = 100
  m_sScaleMin_V = 0
  m_sScaleValue_V = m_sScaleMax_V / 2
  Picture2.Top = Picture1.Top + Picture1.Height - m_sCurOffset - m_sScaleValue_V * (Picture1.Height / m_sScaleMax_V)
  Picture3.Left = Picture1.Left + Picture1.Width - m_sCurOffset - m_sScaleValue_H * (Picture1.Width / m_sScaleMax_H)
'=========================
' Text1 section
'=========================
  m_lWalkWay = m_WheelForward
  m_lMutiplier_Small = 10
  m_lMutiplier_Large = 100
  m_lSampleValue = 100
  Text1.Move 3720, 240
  Text1 = Trim(Str(m_lSampleValue))
'=========================
' ToolTipText section
'=========================
Picture1.ToolTipText = "Ctrl = Scroll Horizontal Shift = 10x speed "
Text1.ToolTipText = "Click to enable  Ctrl = 100x  Shift = 10x  Return to validate Keyboad value entry"
End Sub
Private Sub Form_Click()
m_blnKeepSpinnig = False
DoEvents
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
m_blnKeepSpinnig = False
DoEvents
If m_blnWheelPresent Then
  If Not m_blnWheelTracking Then Call WatchForWheel(Picture1.hwnd)
End If
End Sub
Private Sub Text1_Click()
'**********************************************************
'  if blnWheelArround is set to 'True', we can
'  spin value even mouse away from text box
'  but it seems to be difficult to use any other
'  application (in fact we have to 'Ctrl-Alt-Del' VB to stop
'  if we try to activate other apps)
'
'  - if U know how to make it safe, please let me know -
'
'**********************************************************
If m_blnWheelPresent Then
  If Not m_blnWheelTracking Then Call WatchForWheel(Text1.hwnd, False)
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
'=================================================
'  Kills "No Default Key" Error beep when
'  Keying 'Return' to validate new keyboard value
'=================================================
If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    On Error Resume Next
      m_lSampleValue = CLng(Text1.Text)
  End If
End Sub
Private Sub Text1_LostFocus()
m_blnKeepSpinnig = False
DoEvents
End Sub
Private Sub Form_Load()
initialize
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
m_blnKeepSpinnig = False
m_blnWheelTracking = False
   DoEvents
End Sub
Private Sub Form_Unload(Cancel As Integer)
m_blnKeepSpinnig = False
m_blnWheelTracking = False
   DoEvents
End Sub
```

