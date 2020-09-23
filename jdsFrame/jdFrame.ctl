VERSION 5.00
Begin VB.UserControl jdFrame 
   BackStyle       =   0  'Transparent
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   ClipBehavior    =   0  'Keine
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   79
   ToolboxBitmap   =   "jdFrame.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "jdFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*************************************************
'Jens Duczmal 06.03.2001
'jds Windowless-Frame-Control
'written in 2001 by Jens Duczmal
'
'if any questions / hints pls send mail to
'JayDeeSolutions@web.de
'*************************************************

'*************************************************
'* API Declarations
'*************************************************
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long)
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long

'Constants for DrawEdge-API
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const BF_TOP = &H2
Private Const BF_LEFT = &H1
Private Const BF_BOTTOM = &H8
Private Const BF_RIGHT = &H4
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

'*************************************************
'* Public Enums
'*************************************************
Public Enum efrAppearance
   [2D]
   [3D Sunken]
   [3D Raised]
   [3D Etched]
   [3D Bumped]
   [3D Inset]
   [3D Outset]
End Enum

Public Enum efrBorderStyle
   frbdSolid
   frbdDash
   frbdDot
   frbdDashDot
   frbdDashDotDot
End Enum

Public Enum efrBackStyle
   frbsTransparent
   frbsSolid
End Enum

'*************************************************
'* Constanst with Default-Values
'*************************************************
Private Const cdefAppearance = 3
Private Const cdefBackColor = &H8000000F
Private Const cdefBackStyle = 0
Private Const cdefBorderColor = &H80000008
Private Const cdefBorderWidth = 1
Private Const cdefBorderStyle = 0

'*************************************************
'* Property Variables
'*************************************************
Private m_iAppearance      As Integer
Private m_lBackColor       As Long
Private m_iBackStyle       As Integer
Private m_lBorderColor     As Long
Private m_iBorderStyle     As Integer
Private m_iBorderWidth     As Integer

'*************************************************
'* Working Variables
'*************************************************
Private m_Hdc              As Long
Private m_Hwnd             As Long
Private m_Rect             As RECT
Private m_Regn             As Long

'*************************************************
'* Events
'*************************************************
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'*************************************************************
'Prozedur : Private Method pDraw
'Datum    : 06.03.2001
'Modul    : jdFrame
'Projekt  : jdsFrame
'Parameter:
'-------------------------------------------------------------
'Draws the Frame onto Usercontrols Device-Context
'*************************************************************
Private Sub pDraw()
Dim hPen          As Long     'Handle to new Pen (used for BorderColor)
Dim hPenOld       As Long     'Handle to old Pen used
Dim tp            As POINTAPI 'Just a ref for API LineTo
Dim lOldScale     As Long     'Hold the original Scalemode of Parent Object
Dim ix            As Integer

   'Irgnore Errors if any
   On Error Resume Next

   'First, we should set the Parentobjects Scalemode to pixels
   lOldScale = UserControl.Parent.ScaleMode
   UserControl.Parent.ScaleMode = vbPixels

   'Cause Usercontrol is windowless, it has no own Hwnd.
   'We will use Hwnd of Parent Object
   m_Hwnd = UserControl.Parent.hwnd
   
   'DC to be choosen from Usercontrol.
   'If we would use DC from parent Object we will have some drawing
   'problems sometimes
   m_Hdc = UserControl.hdc
   
   'Fill out RECT-Structure with data
   'Left and Top are always 0 so we could delete them anyway
   m_Rect.Left = UserControl.ScaleLeft
   m_Rect.Top = UserControl.ScaleTop
   m_Rect.Bottom = UserControl.ScaleTop + UserControl.ScaleHeight
   m_Rect.Right = UserControl.ScaleLeft + UserControl.ScaleWidth
   
   'Now start drawing our Frame
   Select Case m_iAppearance
      Case efrAppearance.[2D]
         'For 2D we have to create a Pen first, stating the BorderStyle
         '(Dotted etc.) and the Color
         hPen = CreatePen(m_iBorderStyle, 1, TranslateColor(m_lBorderColor))
         'Nor the new Pen will be "selected" into the DC. SelectObject will return
         'a Reference to the original used Pen. (we have to reset later !)
         hPenOld = SelectObject(m_Hdc, hPen)
         
         'Reduce Height and Width at 1 Pixel
         'otherwise both lines would be drawn outside the Frame
         '(don't know exactly why...but its the truth)
         m_Rect.Right = m_Rect.Right - 1
         m_Rect.Bottom = m_Rect.Bottom - 1

         'Hm... I added this code becuase I got some problems with
         'different PenWidths. Width of 2 draws 2 Points as Btm/Right
         'but just 1 Pt at Top/Left... 3 was o.k and 4 messed up again.
         
         'So now I decide to set PenWidth always to 1 and draw in a loop
         'It will not slow down much...I Assume it must have something to
         'do with the Windowless-Mode.
         For ix = 1 To m_iBorderWidth
            'Move to starting Point and then Draw 4 Lines
            MoveToEx m_Hdc, m_Rect.Left, m_Rect.Top, tp
            LineTo m_Hdc, m_Rect.Right, m_Rect.Top
            LineTo m_Hdc, m_Rect.Right, m_Rect.Bottom
            LineTo m_Hdc, m_Rect.Left, m_Rect.Bottom
            LineTo m_Hdc, m_Rect.Left, m_Rect.Top
            'Reduce Size of the Rectangle at 1 Point
            'to simulate thicker borders
            InflateRect m_Rect, -1, -1
         Next
         
         'we are going to restore the OldPen into DC again
         SelectObject m_Hdc, hPenOld
         'And now we have to delete our own created Pen
         DeleteObject hPen
         'If you don't to this or change the order, this code will
         'eat your system resources...Try it and start a Form with the frame
         'lets say 200 times..Guess your Resources will go down to <10 %
         'But if you do this as stated here, everything will work fine

      Case efrAppearance.[3D Bumped]
         'Draw Edge API will draw a 3D-Rectangle in the specified DC
         DrawEdge m_Hdc, m_Rect, EDGE_BUMP, BF_RECT
      Case efrAppearance.[3D Etched]
         DrawEdge m_Hdc, m_Rect, EDGE_ETCHED, BF_RECT
      Case efrAppearance.[3D Raised]
         DrawEdge m_Hdc, m_Rect, EDGE_RAISED, BF_RECT
      Case efrAppearance.[3D Sunken]
         DrawEdge m_Hdc, m_Rect, EDGE_SUNKEN, BF_RECT
      Case efrAppearance.[3D Inset]
         DrawEdge m_Hdc, m_Rect, BDR_SUNKENOUTER, BF_RECT
      Case efrAppearance.[3D Outset]
         DrawEdge m_Hdc, m_Rect, BDR_RAISEDINNER, BF_RECT
   End Select

   'Finally we have to restore the Old Scalemode of the parent object
   UserControl.Parent.ScaleMode = lOldScale
   
   Exit Sub


End Sub


'*************************************************
'Jens Duczmal 06.03.2001
'Force redraw of Control
'*************************************************
Private Sub RefreshControl()
   'got the problem that if you're changing the Appearance
   'within the IDE, Frame will not be redrawn until
   'moved or resized.
   
   'I tried some Api-Calls but without success
   'For the time being, it works if the Usercontrol
   'will be resized at 1 Pixel and restored again
   
   'No good code...but finally it works.
   UserControl.Extender.Width = UserControl.Extender.Width + 1
   UserControl.Extender.Width = UserControl.Extender.Width - 1
   pDraw
End Sub


'*************************************************
'Jens Duczmal 06.03.2001
'Usercontrol's Properties
'*************************************************
Public Property Get Appearance() As efrAppearance
Attribute Appearance.VB_Description = "Get/Set the Appearance of the Control (2D / 3D etc.)"
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Darstellung"
   Appearance = m_iAppearance
End Property

Public Property Let Appearance(ByVal iAppearance As efrAppearance)
   m_iAppearance = iAppearance
   PropertyChanged "Appearance"
   RefreshControl
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Get/Set the Backgroundcolor "
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
   BackColor = m_lBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   m_lBackColor = New_BackColor
   UserControl.BackColor = m_lBackColor
   PropertyChanged "BackColor"
   RefreshControl
   
End Property

Public Property Get BackStyle() As efrBackStyle
Attribute BackStyle.VB_Description = "Get/Set the Backgroundstyle of the Control (Solid / Transparent)"
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = ";Darstellung"
   BackStyle = m_iBackStyle
End Property
Public Property Let BackStyle(iStyle As efrBackStyle)
   m_iBackStyle = iStyle
   UserControl.BackStyle = m_iBackStyle
   PropertyChanged "BackStyle"
   RefreshControl
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Get/Set the Bordercolor of the Control (only in 2D-Mode)"
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
   BorderColor = m_lBorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   m_lBorderColor = New_BorderColor
   PropertyChanged "BorderColor"
   RefreshControl
End Property
Public Property Let BorderStyle(ByVal iStyle As efrBorderStyle)
Attribute BorderStyle.VB_Description = "Get/Set the Borderstyle of the Control in 2D-Mode (Dash, Dotted....)"
Attribute BorderStyle.VB_ProcData.VB_Invoke_PropertyPut = ";Darstellung"
   If iStyle <> 0 And m_iBorderWidth > 1 Then
      MsgBox "Border Width has been set to 1."
      m_iBorderWidth = 1
   End If
   m_iBorderStyle = iStyle
   PropertyChanged "BorderStyle"
   RefreshControl
End Property

Public Property Get BorderStyle() As efrBorderStyle
   BorderStyle = m_iBorderStyle
End Property
Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Get/Set the Borderwidth of the Control in 2D-Mode (Solid Border only)"
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = ";Darstellung"
   BorderWidth = m_iBorderWidth
End Property

Public Property Let BorderWidth(ByVal iWidth As Integer)
   If iWidth > 1 And m_iBorderStyle <> 0 Then
      MsgBox "Only Solid Border can have width more then 1"
      m_iBorderWidth = 1
   Else
      m_iBorderWidth = iWidth
   End If
   
   
   PropertyChanged "BorderWidth"
   RefreshControl
End Property

Public Sub Refresh()
   UserControl.Refresh
End Sub


Private Sub UserControl_Show()
   On Error Resume Next
   'This will put our Frame into the Background of your Form
   'Only Labels might be in Front of it..
   UserControl.Extender.ZOrder 1
End Sub

Private Sub UserControl_Paint()
   pDraw
End Sub

'*************************************************
'Jens Duczmal 06.03.2001
'HitTest will be raised before all other events.
'It will be raised if the Mouse is over the Usercontrol
'(only if Windowless and Transparent)
'*************************************************
Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
   'We return 3 for HitResult which means that Mouse is over our Frame
   'and therefore it can receive Mousedown / Up / Click etc.-Events
   HitResult = 3
End Sub

'*************************************************
'Jens Duczmal 06.03.2001
'Raise some events
'*************************************************
Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, x, y)
End Sub


Private Sub UserControl_InitProperties()
   m_iAppearance = cdefAppearance
   m_lBackColor = cdefBackColor
   m_iBackStyle = cdefBackStyle
   m_lBorderColor = cdefBorderColor
   m_iBorderStyle = cdefBorderStyle
   m_iBorderWidth = cdefBorderWidth
   UserControl.BackStyle = m_iBackStyle
   UserControl.BackColor = m_lBackColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 
   m_lBackColor = PropBag.ReadProperty("BackColor", cdefBackColor)
   m_iBackStyle = PropBag.ReadProperty("BackStyle", cdefBackStyle)
   m_iAppearance = PropBag.ReadProperty("Appearance", cdefAppearance)
   m_lBorderColor = PropBag.ReadProperty("BorderColor", cdefBorderColor)
   m_iBorderWidth = PropBag.ReadProperty("BorderWidth", cdefBorderWidth)
   m_iBorderStyle = PropBag.ReadProperty("BorderStyle", cdefBorderStyle)
      UserControl.BackStyle = m_iBackStyle
   UserControl.BackColor = m_lBackColor
   pDraw
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("BackColor", m_lBackColor, cdefBackColor)
   Call PropBag.WriteProperty("BackStyle", m_iBackStyle, cdefBackStyle)
   Call PropBag.WriteProperty("Appearance", m_iAppearance, cdefAppearance)
   Call PropBag.WriteProperty("BorderColor", m_lBorderColor, cdefBorderColor)
   Call PropBag.WriteProperty("BorderWidth", m_iBorderWidth, cdefBorderWidth)
   Call PropBag.WriteProperty("BorderStyle", m_iBorderStyle, cdefBorderStyle)
End Sub


'*************************************************
'Jens Duczmal 05.03.2001
'Translates a SystemColor into a StandardColor.
'Original Source taken from Randy Birch (I assume)
'*************************************************
Private Function TranslateColor(ByVal lVbColor As Long) As OLE_COLOR
    Dim lColor As Long
    On Error GoTo Err_Handler
    OleTranslateColor lVbColor, 0, lColor
    TranslateColor = lColor
Exit_Proc:
    Exit Function
Err_Handler:

    GoTo Exit_Proc
    Resume
End Function
