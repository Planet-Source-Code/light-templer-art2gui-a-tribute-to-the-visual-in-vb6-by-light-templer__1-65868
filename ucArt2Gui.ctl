VERSION 5.00
Begin VB.UserControl ucArt2Gui 
   Alignable       =   -1  'True
   Appearance      =   0  '2D
   AutoRedraw      =   -1  'True
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   ClipBehavior    =   0  'Keine
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   117
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   248
   ToolboxBitmap   =   "ucArt2Gui.ctx":0000
End
Attribute VB_Name = "ucArt2Gui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
'   ucArt2Gui.ctl
'

'   Started     :   10/19/2005
'   Created by  :   Light Templer
'   Last edit   :   07/21/2006

'   VERSION     :   1.03


'   CREDITS:
'                   To Carles P.V.
'                   All gradient subs used here are created by Carles.
'                   e.g.:  http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=60580&lngWId=1
'                   During development of this uc I had a problem with the radial gradients,
'                   I need a change in behavior I couldn't do on my own. His email with the
'                   code modifications was faster than light! :-)
'                   This usercontrol is kind of a tribut to Carles' graphic submissions on PSC.
'                   Thanks for all your code and your help!
'
'
'                   To Keith "LaVolpe" Fox
'                   I give his submission
'                   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=58562&lngWId=1
'                   a home, as he wanted to. When Carles is "Mr. Gradient" on PSC VB, Keith is "Mr. Region" ;-)))
'                   Many thanks for going into the details with 31 easy to handle shapes!
'                   Without his clipping region inspiration and his code this usercontrol wouldn't be
'                   what it is!
'
'
'                   To Randy Birch from vbNet
'                   http://vbnet.mvps.org/
'                   Thanks alot for one of the best and oldest VB sites on available on web!
'                   I used code from his site to select a color from standard color dialog.
'
'
'                   To Dana Seaman
'                   I took the anti-aliased circles sub from the PSC VB submission "Anti-Alias 2D Engine"
'                   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=38582&lngWId=1
'                   It tooks this uc from very ugly to very smooth looking.
'                   Thx for sharing this rarely type of code!
'
'
'                   To unknown dialog designer from an installer software company
'                   While installing Adobe PhotoShop CS(tm) for the company I work for
'                   I admired the fantastic dialog design of the installation wizzard which tooks me
'                   through the neccessary install steps.
'                   It leads directly to this usercontrol and the 'CoronaWithLineRight' design.
'                   Thx for this inspiration :-) !
'

'   HISTORY:
'
'   06.07.2006      V. 1.01 - Added an init flag to prevent an error raised in Sub PaintMemToDC() on
'                   dropping a new control onto a form when VB IDE - menu <Options> , Tab <General>,
'                   <Selection> Error trapping is set to 'Break on All Errors'.
'                   Many thanks to Riccardo Cohen for his comment to this problem!
'
'   21.07.2006      V. 1.02 - After a longer period of confusion Richard Mewett (MANY thanks!) discovered
'                   the problem with the different formats ( 12,345 vs 12.345 ) VB saves single var as
'                   properties in .frm files. Because of my German VB it wasn't possible to me to reproduce
'                   this problem on my development system. The small change in Usercontrol_ReadProperties()
'                   now handles both versions. Hope, thats it, folks and sorry for any trouble!
'
'   12.06.2006      Discovered a strange difference between VB6 ServicePack versions (5 and 6)
'                   After adding the statement  'Usercontrol.Refresh' to the end of  Sub DrawDesign() now
'                   everything should be alright.


Option Explicit


' ***********
' *  CONSTS *
' ***********
Private Const m_def_Design          As Long = 1&        ' enDesign.A2G_CoronaWithLineRight
Private Const m_def_Shape           As Long = 1&        ' enShape.A2G_Rectangle
Private Const m_def_ShapeLT         As Long = 1&        ' enShapeLeftTop.A2G_Shape_LT_LTBR
Private Const m_def_ShapeRB         As Long = 1&        ' enShapeRightBottom.A2G_Shape_BR_LBTR
Private Const m_def_LinGradWidth    As Long = 2&        ' 2 pixels wide
Private Const m_def_HiliteColor     As Long = &HDCDCDC  ' Nearly white
Private Const m_def_RadGradWidth    As Long = 100&
Private Const m_def_Angle           As Long = 0&
Private Const m_def_ArcSize         As Long = 20&       ' Size of arcs when shape is Rounded Rectangle

Private Const GRAB_HANDLE_SIZE      As Long = 10&
Private Const MAX_GRAD_AREAS        As Long = 20&       ' Used for grab handle array, too!

Private Const API_DIB_RGB_COLORS    As Long = 0&
Private Const API_INVALID_COLOR     As Long = -1&       ' Result from 'ColorToRGB()' if OleTranslateColor() fails

Private Const RGN_OR                As Long = 2&
Private Const RGN_XOR               As Long = 3&

Private Const CC_RGBINIT            As Long = &H1&
Private Const CC_FULLOPEN           As Long = &H2&
Private Const CC_ANYCOLOR           As Long = &H100&

Private Const PI                    As Single = 3.14159265358979
Private Const HalfPi                As Single = PI / 2
Private Const TO_DEG                As Single = 180 / PI
Private Const TO_RAD                As Single = PI / 180

Private Const cThin                 As Single = PI * 0.34
Private Const cThick                As Single = PI * 0.17



' ******************
' *  PUBLIC EVENTS *
' ******************
Public Event Error(lErrNo As Long, sErrMsg As String)



' ******************
' *  PUBLIC ENUMS  *
' ******************
Public Enum enDesign
    A2G_CoronaWithLineRight = 1
    A2G_CoronaWithLineLeft = 2
    A2G_LinearGradient = 3
    A2G_RadialGradient = 4
    A2G_MultiLinGradHor = 5
End Enum

Public Enum enShape
    A2G_Rectangle = 1
    A2G_Curved_Rectangle = 2
    A2G_Square = 3
    A2G_Ellipse = 4
    A2G_Circle = 5
    A2G_Hexagon = 6
    A2G_Octagon = 7
    A2G_Diagonal_Rectangle = 8
    A2G_Diamond = 9
End Enum

Public Enum enShapeLeftTop
    A2G_Shape_LT_LBTR = -1
    A2G_Shape_LT_VERT = 0
    A2G_Shape_LT_LTBR = 1
End Enum

Public Enum enShapeRightBottom
    A2G_Shape_BR_LTBBR = -1
    A2G_Shape_BR_VERT = 0
    A2G_Shape_BR_LBTR = 1
End Enum


' ******************
' *  PRIVATE ENUMS *
' ******************
Private Enum cThickness
   Thin
   Thick
End Enum

Private Enum enColor
    COLOR_BACKGROUND = 1
    COLOR_FOREGROUND = 2
    COLOR_HILITE = 3
    COLOR_AREA = 4
End Enum

Private Enum enGradientDirection
    gdHorizontal = 0
    gdVertical = 1
    gdDownwardDiagonal = 2
    gdUpwardDiagonal = 3
End Enum

Private Enum enGrabHandleType
    GHT_Color = 1
    GHT_Position = 2
    GHT_Size = 3
    GHT_Angle = 4
    GHT_Gamma = 5
    GHT_ShapeCorner = 6
End Enum


' ******************
' *  PRIVATE TYPES *
' ******************
Private Type tpPoint
    lX                  As Long
    lY                  As Long
End Type

Private Type tpRectXYWH
    lX                  As Long
    lY                  As Long
    lWidth              As Long
    lHeight             As Long
End Type

Private Type tpRectX1Y1X2Y2
    lX1                 As Long
    lY1                 As Long
    lX2                 As Long
    lY2                 As Long
End Type

Private Type tpBitmapInMemory
    lWidth              As Long
    lHeight             As Long
    lArrBitmap()        As Long
End Type

Private Type tpGrabHandle
    lX                  As Long
    lY                  As Long
    GrabHandleType      As enGrabHandleType
End Type

Private Type tpGradArea
    sngPosition         As Single                   ' Distance to start in percent / 100
    oColor              As OLE_COLOR                ' Destination color ( 'from' color is taken from last element of array)
    sngGradGamma        As Single                   ' "Velocity" of gradient
End Type

Private Type tpMvar
    
    ' Properties
    Design              As enDesign
    Shape               As enShape
    ShapeLT             As enShapeLeftTop
    ShapeRB             As enShapeRightBottom
    lLinGradWidth       As Long
    lRadGradWidth       As Long
    lAngle              As Long
    lArcSize            As Long                     ' Size of arcs when shape is Rounded Rectangle
    lGradientAreas      As Long                     ' For A2G_MultiLinGradHor and A2G_MultiLinGradVer: How many gradients?
    HiliteColor         As OLE_COLOR
    
    ' Privates
    flgIsInitDone       As Boolean
    RectLinGradP1       As tpRectXYWH               ' Part 1 of rect with grad on top and bottom
    RectLinGradP2       As tpRectX1Y1X2Y2           ' Part 2               "
    RectLinGradP3       As tpRectXYWH               ' Part 3               "
    RadGradient         As tpBitmapInMemory         ' Current radial gradient in memory
    RectRadGradDst      As tpRectXYWH               ' Radial gradient destination infos
    LinGradient         As tpBitmapInMemory
    Center              As tpPoint                  ' Center of radial gradient
    flgEditMode         As Boolean                  ' TRUE:  Uc is set to 'edit' mode by right mouseclick menu
    flgRefreshRegion    As Boolean                  ' TRUE:  Clipping region needs new size calculation and assignment
    CurrMovingGrabber   As Long                     ' Number of grabber (index into ArrGrabHandles() ) we are just moving (0 = none)
    LastMousePosition   As tpPoint
    lGrabHandles        As Long                     ' How many grab handles has the current design?
    lOrgAreaPosition    As Single
    
    ' Arrays
    arrGradAreas()                              As tpGradArea   ' Redimmed with 'lGradientAreas' in UserControl_ReadProperties()
    arrOriginalCol(1 To 4)                      As OLE_COLOR    ' On start of interactive change of a color save original color
    arrGrabHandles(1 To 3 * MAX_GRAD_AREAS + 1) As tpGrabHandle ' In EDIT mode here we holds the current set of grab handles
    
End Type


Private Type tpAPI_BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

Private Type tpAPI_CHOOSECOLORSTRUCT
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    rgbResult           As Long
    lpCustColors        As Long
    flags               As Long
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type


' *****************************
' *  PRIVATE API DECLARATIONS *
' *****************************
Private Declare Function API_StretchDIBits Lib "gdi32" Alias "StretchDIBits" _
        (ByVal hdc As Long, _
         ByVal x As Long, _
         ByVal y As Long, _
         ByVal dx As Long, _
         ByVal dy As Long, _
         ByVal SrcX As Long, _
         ByVal SrcY As Long, _
         ByVal wSrcWidth As Long, _
         ByVal wSrcHeight As Long, _
         ByRef lpBits As Any, _
         ByRef lpBitsInfo As Any, _
         ByVal wUsage As Long, _
         ByVal dwRop As Long) As Long

Private Declare Function API_CreateSolidBrush Lib "gdi32.dll" Alias "CreateSolidBrush" _
        (ByVal crColor As Long) As Long

Private Declare Function API_FillRect Lib "user32.dll" Alias "FillRect" _
        (ByVal hdc As Long, _
         lpRect As tpRectX1Y1X2Y2, _
         ByVal hBrush As Long) As Long

Private Declare Function API_DeleteObject Lib "gdi32.dll" Alias "DeleteObject" _
        (ByVal hObject As Long) As Long

Private Declare Function API_OleTranslateColor Lib "oleaut32.dll" Alias "OleTranslateColor" _
        (ByVal lOLEColor As Long, _
         ByVal lHPalette As Long, _
         lColorRef As Long) As Long

Private Declare Function API_GetPixel Lib "gdi32.dll" Alias "GetPixel" _
        (ByVal hdc As Long, _
         ByVal x As Long, _
         ByVal y As Long) As Long
         
Private Declare Function API_SetPixelV Lib "gdi32.dll" Alias "SetPixelV" _
        (ByVal hdc As Long, _
         ByVal x As Long, _
         ByVal y As Long, _
         ByVal crColor As Long) As Long

Private Declare Function API_ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" _
        (lpcc As tpAPI_CHOOSECOLORSTRUCT) As Long

Private Declare Function API_CreateRectRgn Lib "gdi32.dll" Alias "CreateRectRgn" _
        (ByVal X1 As Long, _
         ByVal Y1 As Long, _
         ByVal X2 As Long, _
         ByVal Y2 As Long) As Long
         
Private Declare Function API_SetWindowRgn Lib "user32.dll" Alias "SetWindowRgn" _
        (ByVal hWnd As Long, _
         ByVal hRgn As Long, _
         ByVal bRedraw As Boolean) As Long
         
Private Declare Function API_CombineRgn Lib "gdi32.dll" Alias "CombineRgn" _
        (ByVal hDestRgn As Long, _
         ByVal hSrcRgn1 As Long, _
         ByVal hSrcRgn2 As Long, _
         ByVal nCombineMode As Long) As Long
                 
Private Declare Function API_FrameRgn Lib "gdi32.dll" Alias "FrameRgn" _
        (ByVal hdc As Long, _
         ByVal hRgn As Long, _
         ByVal hBrush As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long) As Long
         
Private Declare Function API_OffsetRgn Lib "gdi32.dll" Alias "OffsetRgn" _
        (ByVal hRgn As Long, _
         ByVal x As Long, _
         ByVal y As Long) As Long
         
Private Declare Function API_FillRgn Lib "gdi32.dll" Alias "FillRgn" _
        (ByVal hdc As Long, _
         ByVal hRgn As Long, _
         ByVal hBrush As Long) As Long

Private Declare Function API_CreatePolygonRgn Lib "gdi32.dll" Alias "CreatePolygonRgn" _
        (ByRef lpPoint As tpPoint, _
         ByVal nCount As Long, _
         ByVal nPolyFillMode As Long) As Long

Private Declare Function API_CreateEllipticRgn Lib "gdi32.dll" Alias "CreateEllipticRgn" _
        (ByVal X1 As Long, _
         ByVal Y1 As Long, _
         ByVal X2 As Long, _
         ByVal Y2 As Long) As Long

Private Declare Function API_CreateRoundRectRgn Lib "gdi32.dll" Alias "CreateRoundRectRgn" _
        (ByVal lX1 As Long, _
         ByVal lY1 As Long, _
         ByVal lX2 As Long, _
         ByVal lY2 As Long, _
         ByVal lWidthEllipse As Long, _
         ByVal lHeightEllipse As Long) As Long



' *****************
' *  PRIVATE VARS *
' *****************
Private Mvar As tpMvar
'
'
'




' ******************************
' *  PRIVATE UserControl Stuff *
' ******************************
Private Sub UserControl_InitProperties()
    
    On Local Error Resume Next
    
    UserControl.BackColor = UserControl.Ambient.BackColor
    With Mvar
        .HiliteColor = m_def_HiliteColor
        .Center.lX = UserControl.ScaleWidth / 2
        .Center.lY = UserControl.ScaleHeight / 2
        .lArcSize = m_def_ArcSize
                
        ' Settings for A2G_CoronaWithLineRight as default. Change them if you prefer a different default design.
        Me.Design = m_def_Design
        Me.Shape = m_def_Shape
        Me.ShapeLT = m_def_ShapeLT
        Me.ShapeRB = m_def_ShapeRB
        .lLinGradWidth = m_def_LinGradWidth
        .lRadGradWidth = (UserControl.ScaleHeight / 3) * 2 - .lLinGradWidth
        
        .flgIsInitDone = True
    End With

End Sub

Private Sub UserControl_Terminate()
    
    RemoveClippingRegion

End Sub

Private Sub UserControl_LostFocus()
    
    With Mvar
        If .flgEditMode = True Then
            
            ' Leaving EDIT mode
            .flgEditMode = False
            UserControl.Cls
            DrawDesign
        End If
    End With
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim i           As Long
    Dim lHalfWidth  As Long
    
    With Mvar
        If .flgEditMode = True And Button <> 0 Then
            
            ' Only in EDIT mode when a mMouse button pressed:
            lHalfWidth = GRAB_HANDLE_SIZE / 2
            For i = 1 To .lGrabHandles
                If Abs(.arrGrabHandles(i).lX - x) <= lHalfWidth Then
                    If Abs(.arrGrabHandles(i).lY - y) <= lHalfWidth Then
                        ' Save current values to see the diff when mouse starts moving
                        .CurrMovingGrabber = i
                        .LastMousePosition.lX = x
                        .LastMousePosition.lY = y
                        
                        If .arrGrabHandles(i).GrabHandleType <> GHT_ShapeCorner Then
                            .arrOriginalCol(COLOR_BACKGROUND) = UserControl.BackColor
                            .arrOriginalCol(COLOR_FOREGROUND) = UserControl.ForeColor
                            .arrOriginalCol(COLOR_HILITE) = .HiliteColor
                            If .Design = A2G_MultiLinGradHor Then
                                .arrOriginalCol(COLOR_AREA) = .arrGradAreas((i + 2) \ 3).oColor
                                .lOrgAreaPosition = .arrGradAreas((i + 2) \ 3).sngPosition
                            End If
                        End If
                        
                        Exit For
                    End If
                End If
            Next i
        End If
    End With
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim lDeltaX     As Long
    Dim lDeltaY     As Long
    Dim lNewAngle   As Long
    Dim lDistance   As Long
    Dim i           As Long

    With Mvar
        If .flgEditMode = False Then
            If UserControl.Ambient.UserMode = False Then    ' Got a _MouseMove() event, but app isn't running?
                .flgEditMode = True

                ' ### We have entered this special EDIT mode. Prepare all we need for

                Select Case .Design

                    Case A2G_CoronaWithLineRight
                            SetGrabHandleInfos 4, GHT_Size, GHT_Size, GHT_Color, GHT_Color
                            
                            .arrGrabHandles(1).lX = .RectRadGradDst.lX
                            .arrGrabHandles(1).lY = .RectRadGradDst.lY + (.RectRadGradDst.lHeight / 2)

                            .arrGrabHandles(2).lX = UserControl.ScaleWidth - (.lLinGradWidth + (GRAB_HANDLE_SIZE / 2))
                            .arrGrabHandles(2).lY = UserControl.ScaleHeight / 10

                            .arrGrabHandles(3).lX = UserControl.ScaleWidth - .lLinGradWidth - GRAB_HANDLE_SIZE
                            .arrGrabHandles(3).lY = UserControl.ScaleHeight / 2 - (3 * GRAB_HANDLE_SIZE)

                            .arrGrabHandles(4).lX = UserControl.ScaleWidth - (1 * GRAB_HANDLE_SIZE)
                            .arrGrabHandles(4).lY = UserControl.ScaleHeight / 2 + (3 * GRAB_HANDLE_SIZE)


                    Case A2G_CoronaWithLineLeft
                            SetGrabHandleInfos 4, GHT_Size, GHT_Size, GHT_Color, GHT_Color
                            
                            .arrGrabHandles(1).lX = .lLinGradWidth + .RectRadGradDst.lWidth - (GRAB_HANDLE_SIZE / 2)
                            .arrGrabHandles(1).lY = .RectRadGradDst.lY + (.RectRadGradDst.lHeight / 2)
                            
                            .arrGrabHandles(2).lX = .lLinGradWidth + (GRAB_HANDLE_SIZE / 2)
                            .arrGrabHandles(2).lY = UserControl.ScaleHeight / 10
                            
                            .arrGrabHandles(3).lX = .lLinGradWidth + GRAB_HANDLE_SIZE
                            .arrGrabHandles(3).lY = UserControl.ScaleHeight / 2 - (3 * GRAB_HANDLE_SIZE)
                            
                            .arrGrabHandles(4).lX = GRAB_HANDLE_SIZE
                            .arrGrabHandles(4).lY = UserControl.ScaleHeight / 2 + (3 * GRAB_HANDLE_SIZE)


                    Case A2G_LinearGradient
                            SetGrabHandleInfos 3, GHT_Angle, GHT_Color, GHT_Color
                            
                            PutDotOnCircle .lAngle, 24, .arrGrabHandles(1)
                            PutDotOnRectangle .lAngle, 2 * GRAB_HANDLE_SIZE, .arrGrabHandles(2)
                            PutDotOnRectangle .lAngle + 180, 2 * GRAB_HANDLE_SIZE, .arrGrabHandles(3)


                    Case A2G_RadialGradient
                            SetGrabHandleInfos IIf(.Shape = A2G_Curved_Rectangle, 4, 3), GHT_Position, GHT_Color, GHT_Color

                            .arrGrabHandles(1).lX = .Center.lX
                            .arrGrabHandles(1).lY = .Center.lY
                            
                            .arrGrabHandles(2).lX = .Center.lX + (2 * GRAB_HANDLE_SIZE)
                            .arrGrabHandles(2).lY = .Center.lY + (2 * GRAB_HANDLE_SIZE)
                            
                            .arrGrabHandles(3).lX = UserControl.ScaleWidth - (2 * GRAB_HANDLE_SIZE)
                            .arrGrabHandles(3).lY = UserControl.ScaleHeight - (2 * GRAB_HANDLE_SIZE)
                            
                            
                    Case A2G_MultiLinGradHor
                            .lGrabHandles = .lGradientAreas * 3
                            For i = 0 To .lGradientAreas - 1
                                
                                .arrGrabHandles(i * 3 + 1).GrabHandleType = GHT_Position
                                .arrGrabHandles(i * 3 + 1).lX = UserControl.ScaleWidth * .arrGradAreas(i + 1).sngPosition - GRAB_HANDLE_SIZE
                                .arrGrabHandles(i * 3 + 1).lY = GRAB_HANDLE_SIZE
                                
                                .arrGrabHandles(i * 3 + 2).GrabHandleType = GHT_Color
                                .arrGrabHandles(i * 3 + 2).lX = UserControl.ScaleWidth * .arrGradAreas(i + 1).sngPosition - GRAB_HANDLE_SIZE
                                .arrGrabHandles(i * 3 + 2).lY = GRAB_HANDLE_SIZE * 3
                                
                                .arrGrabHandles(i * 3 + 3).GrabHandleType = GHT_Gamma
                                .arrGrabHandles(i * 3 + 3).lX = UserControl.ScaleWidth * .arrGradAreas(i + 1).sngPosition - GRAB_HANDLE_SIZE
                                .arrGrabHandles(i * 3 + 3).lY = GRAB_HANDLE_SIZE * 5
                                
                            Next i
                            
                End Select
                
                ' Add grab handle for changing shape parameter
                Select Case .Shape

                    Case A2G_Curved_Rectangle
                            .lGrabHandles = .lGrabHandles + 1
                            With .arrGrabHandles(.lGrabHandles)
                                .GrabHandleType = GHT_ShapeCorner
                                .lX = Mvar.lArcSize / 2
                                .lY = UserControl.ScaleHeight - .lX
                            End With
                            
                    Case A2G_Hexagon
                            ' Maybe later ...
        
                    Case A2G_Octagon
                            ' Room for extensions ... ;-)

                End Select
                
                DrawDesign

            End If
            
        ElseIf .CurrMovingGrabber <> 0 And Button <> 0 Then
            
            ' A grabber is dragged
            
            lDeltaX = .LastMousePosition.lX - x
            lDeltaY = .LastMousePosition.lY - y
            Select Case .Design

                Case A2G_CoronaWithLineRight
                        If .CurrMovingGrabber = 1 Then
                            ' When grabber moved adjust SIZE of radial gradient
                            If lDeltaX <> 0 Then
                                .LastMousePosition.lX = x
                                .arrGrabHandles(1).lX = .arrGrabHandles(1).lX - lDeltaX
                                Me.RadialGradientWidth = ((UserControl.ScaleWidth - .lLinGradWidth) - x) * 2
                            End If

                        ElseIf .CurrMovingGrabber = 2 Then
                            ' When grabber moved adjust WIDTH of linear gradient
                            If lDeltaX <> 0 Then
                                .LastMousePosition.lX = x

                                If .arrGrabHandles(2).lX - lDeltaX < UserControl.ScaleWidth - 1 And _
                                        .arrGrabHandles(2).lX - lDeltaX > GRAB_HANDLE_SIZE Then
                                    .arrGrabHandles(2).lX = .arrGrabHandles(2).lX - lDeltaX
                                    Me.LinearGradientWidth = Me.LinearGradientWidth + lDeltaX

                                    ' Adjust X position of other grab handles
                                    .arrGrabHandles(1).lX = .arrGrabHandles(1).lX - lDeltaX
                                    .arrGrabHandles(3).lX = .arrGrabHandles(3).lX - lDeltaX
                                End If
                            End If

                        ElseIf .CurrMovingGrabber = 3 Then
                            ' Adjust hilite color of radial gradient
                            ChangeColor Button, lDeltaX, COLOR_HILITE
                            
                        ElseIf .CurrMovingGrabber = 4 Then
                            ' Adjust forecolor of linear gradient part
                            ChangeColor Button, lDeltaX, COLOR_FOREGROUND
                            
                        End If


                Case A2G_CoronaWithLineLeft
                        If .CurrMovingGrabber = 1 Then
                            ' When grabber moved adjust size of radial gradient
                            If lDeltaX <> 0 Then
                                .LastMousePosition.lX = x
                                .arrGrabHandles(1).lX = .arrGrabHandles(1).lX - lDeltaX
                                Me.RadialGradientWidth = (x - .lLinGradWidth) * 2
                            End If

                        ElseIf .CurrMovingGrabber = 2 Then
                            ' When grabber moved adjust width of linear gradient
                            If lDeltaX <> 0 Then
                                .LastMousePosition.lX = x

                                If .arrGrabHandles(2).lX - lDeltaX > -1 And _
                                        .arrGrabHandles(2).lX - lDeltaX < UserControl.ScaleWidth - GRAB_HANDLE_SIZE Then
                                    .arrGrabHandles(2).lX = .arrGrabHandles(2).lX - lDeltaX
                                    Me.LinearGradientWidth = Me.LinearGradientWidth - lDeltaX

                                    ' Adjust X position of other grab handles
                                    .arrGrabHandles(1).lX = .arrGrabHandles(1).lX - lDeltaX
                                    .arrGrabHandles(3).lX = .arrGrabHandles(3).lX - lDeltaX
                                End If
                            End If

                        ElseIf .CurrMovingGrabber = 3 Then
                            ' Adjust hilite color of radial gradient
                            ChangeColor Button, lDeltaX, COLOR_HILITE
                            
                        ElseIf .CurrMovingGrabber = 4 Then
                            ' Adjust forecolor of linear gradient part
                            ChangeColor Button, lDeltaX, COLOR_FOREGROUND
                        
                        End If


                Case A2G_LinearGradient
                        If .CurrMovingGrabber = 1 Then
                            ' When grabber moved adjust angle of linear gradient
                            If lDeltaX Or lDeltaY Then
                                .LastMousePosition.lX = x
                                .LastMousePosition.lY = y

                                With UserControl
                                    lNewAngle = GetAngle(.ScaleWidth / 2, .ScaleHeight / 2, x, .ScaleHeight - y)
                                    If lNewAngle < 0 Then
                                        lNewAngle = lNewAngle + 360
                                    End If

                                    PutDotOnRectangle lNewAngle, 2 * GRAB_HANDLE_SIZE, Mvar.arrGrabHandles(2)
                                    PutDotOnRectangle lNewAngle + 180, 2 * GRAB_HANDLE_SIZE, Mvar.arrGrabHandles(3)

                                    Me.Angle = lNewAngle
                                End With
                            End If

                        ElseIf .CurrMovingGrabber = 2 Then
                            ' Adjust background ('from') color of linear gradient
                            ChangeColor Button, lDeltaX, COLOR_BACKGROUND
                            
                        ElseIf .CurrMovingGrabber = 3 Then
                            ' Adjust highlight ('to') color of linear gradient
                            ChangeColor Button, lDeltaX, COLOR_HILITE
                            
                        End If


                Case A2G_RadialGradient
                        If .CurrMovingGrabber = 1 Then
                            ' When grabber moved adjust center of radial gradient
                            If lDeltaX Then
                                .Center.lX = .Center.lX - lDeltaX
                                PropertyChanged "CenterX"
                            End If
                            If lDeltaY Then
                                .Center.lY = .Center.lY - lDeltaY
                                PropertyChanged "CenterY"
                            End If

                            If lDeltaX Or lDeltaY Then
                                .LastMousePosition.lX = x
                                .LastMousePosition.lY = y
                                .arrGrabHandles(1).lX = x
                                .arrGrabHandles(1).lY = y
                                .arrGrabHandles(2).lX = x + (2 * GRAB_HANDLE_SIZE)
                                .arrGrabHandles(2).lY = y + (2 * GRAB_HANDLE_SIZE)
                                Recalc_Dimensions
                                DrawDesign
                            End If

                        ElseIf .CurrMovingGrabber = 2 Then
                            ' Adjust hilight color of radial gradient
                            ChangeColor Button, lDeltaX, COLOR_HILITE

                        ElseIf .CurrMovingGrabber = 3 Then
                            ' Adjust background color of radial gradient
                            ChangeColor Button, lDeltaX, COLOR_BACKGROUND

                        ElseIf (.CurrMovingGrabber = 4 And .Shape <> A2G_Curved_Rectangle) Or _
                                (.CurrMovingGrabber = 5 And .Shape = A2G_Curved_Rectangle) Then
                            ' When grabber moved adjust size of radial gradient
                            If lDeltaX Or lDeltaY Then
                                .LastMousePosition.lX = x
                                .LastMousePosition.lY = y
                                .arrGrabHandles(.lGrabHandles).lX = .arrGrabHandles(.lGrabHandles).lX - lDeltaX
                                .arrGrabHandles(.lGrabHandles).lY = .arrGrabHandles(.lGrabHandles).lY - lDeltaY
                                Me.RadialGradientWidth = GetDistance(.Center.lX, .Center.lY, x, y) * 2
                            End If
                        End If
                
                
                Case A2G_MultiLinGradHor
                        If .arrGrabHandles(.CurrMovingGrabber).GrabHandleType = GHT_Position Then
                            ChangeAreaPosition Button, x
                            
                            For i = 0 To .lGradientAreas - 1
                                .arrGrabHandles(i * 3 + 1).lX = UserControl.ScaleWidth * .arrGradAreas(i + 1).sngPosition - GRAB_HANDLE_SIZE
                                .arrGrabHandles(i * 3 + 2).lX = UserControl.ScaleWidth * .arrGradAreas(i + 1).sngPosition - GRAB_HANDLE_SIZE
                                .arrGrabHandles(i * 3 + 3).lX = UserControl.ScaleWidth * .arrGradAreas(i + 1).sngPosition - GRAB_HANDLE_SIZE
                            Next i
                            
                            Recalc_Dimensions True
                            DrawDesign
                            
                        ElseIf .arrGrabHandles(.CurrMovingGrabber).GrabHandleType = GHT_Color Then
                            ChangeColor Button, lDeltaX, COLOR_AREA
                            Recalc_Dimensions True
                            DrawDesign
                            
                        ElseIf .arrGrabHandles(.CurrMovingGrabber).GrabHandleType = GHT_Gamma Then
                            ChangeGamma Button, lDeltaX
                            Recalc_Dimensions True
                            DrawDesign
                            
                        End If

            End Select
            
            If .arrGrabHandles(.CurrMovingGrabber).GrabHandleType = GHT_ShapeCorner Then
                ' We adjust the clipping region
                
                Select Case .Shape

                    Case A2G_Curved_Rectangle
                            If lDeltaX Then
                                Cls
                                .LastMousePosition.lX = x
                                .LastMousePosition.lY = y
                                With .arrGrabHandles(.lGrabHandles)
                                    .lX = x
                                    .lY = UserControl.ScaleHeight - x
                                End With
                                Mvar.flgRefreshRegion = True
                                Me.ArcSize = x * 2
                            End If
                                
                    Case A2G_Hexagon
                            ' Maybe later ...
                                                    
                    Case A2G_Octagon
                            ' Room for improvements ... ;-)

                End Select
                
            End If
            
        ElseIf Button = 0 And .Design = A2G_RadialGradient Then
            ' No button pressed, design is radial gradient and we are close to the circle?
            ' -> Show grab handle to adjust size of circle.
            lDistance = GetDistance(.Center.lX, .Center.lY, x, y)
            If lDistance > .lRadGradWidth / 2 - (2 * GRAB_HANDLE_SIZE) And _
                    lDistance < .lRadGradWidth / 2 + (2 * GRAB_HANDLE_SIZE) Then
                
                .lGrabHandles = IIf(.Shape = A2G_Curved_Rectangle, 5, 4)
                lNewAngle = 360 - GetAngle(.Center.lX, .Center.lY, x, y)

                .arrGrabHandles(.lGrabHandles).lX = .Center.lX + (.lRadGradWidth / 2) * Cos((360 - lNewAngle) * TO_RAD)
                .arrGrabHandles(.lGrabHandles).lY = .Center.lY + (.lRadGradWidth / 2) * Sin((360 - lNewAngle) * TO_RAD)
                UserControl.Cls
                DrawDesign
            Else
                If .lGrabHandles = 5 Or .lGrabHandles = 4 Then
                    .lGrabHandles = .lGrabHandles - 1
                    UserControl.Cls
                    DrawDesign
                End If
            End If
        End If

    End With
    UserControl.Refresh
    
End Sub

Private Sub SetGrabHandleInfos(lNumberOf As Long, ParamArray ParaArrGHTypes())
    ' Just a shotcut to save many lines of code
    
    Dim i As Long
    
    With Mvar
        .lGrabHandles = lNumberOf
        For i = 0 To UBound(ParaArrGHTypes())
            .arrGrabHandles(i + 1).GrabHandleType = ParaArrGHTypes(i)
        Next i
    End With
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Used in EDIT mode: We don't move a grabber anymore
    Mvar.CurrMovingGrabber = 0
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Dim i       As Long
    Dim sVal    As String
    
    On Error Resume Next
    
    
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    
    With Mvar
        .Design = PropBag.ReadProperty("Design", m_def_Design)
        .Shape = PropBag.ReadProperty("Shape", m_def_Shape)
        .ShapeLT = PropBag.ReadProperty("ShapeLT", m_def_ShapeLT)
        .ShapeRB = PropBag.ReadProperty("ShapeRB", m_def_ShapeRB)
        .HiliteColor = PropBag.ReadProperty("HiliteColor", m_def_HiliteColor)
        .lRadGradWidth = PropBag.ReadProperty("RadGradWidth", m_def_RadGradWidth)
        .lLinGradWidth = PropBag.ReadProperty("LinGradWidth", m_def_LinGradWidth)
        .lAngle = PropBag.ReadProperty("Angle", m_def_Angle)
        .Center.lX = PropBag.ReadProperty("CenterX", UserControl.ScaleWidth / 2)
        .Center.lY = PropBag.ReadProperty("CenterY", UserControl.ScaleHeight / 2)
        .lArcSize = PropBag.ReadProperty("ArcSize", m_def_ArcSize)
            
        ' Handle multi value property 'Gradient areas'
        .lGradientAreas = PropBag.ReadProperty("GradientAreas", 1)
        If .lGradientAreas > 0 Then
            ReDim .arrGradAreas(1 To .lGradientAreas)
            For i = 1 To .lGradientAreas
                .arrGradAreas(i).oColor = PropBag.ReadProperty("GradArea-Color" & i, vbWhite)
                
                sVal = PropBag.ReadProperty("GradArea-Position" & i, 1)
                .arrGradAreas(i).sngPosition = LocalProofGetSingle(sVal)
                
                sVal = PropBag.ReadProperty("GradArea-GradGamma" & i, 1)
                .arrGradAreas(i).sngGradGamma = LocalProofGetSingle(sVal)
                
            Next i
        End If
    End With
    
    Recalc_Dimensions
    Mvar.flgRefreshRegion = True
    Mvar.flgIsInitDone = True
    DrawDesign
    
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Dim i           As Long
    
    With PropBag
        .WriteProperty "BackColor", UserControl.BackColor, &H80000005
        .WriteProperty "ForeColor", UserControl.ForeColor, &H80000012
        .WriteProperty "Design", Mvar.Design, m_def_Design
        .WriteProperty "Shape", Mvar.Shape, m_def_Shape
        .WriteProperty "ShapeLT", Mvar.ShapeLT, m_def_ShapeLT
        .WriteProperty "ShapeRB", Mvar.ShapeRB, m_def_ShapeRB
        .WriteProperty "HiliteColor", Mvar.HiliteColor, m_def_HiliteColor
        .WriteProperty "RadGradWidth", Mvar.lRadGradWidth, m_def_RadGradWidth
        .WriteProperty "LinGradWidth", Mvar.lLinGradWidth, m_def_LinGradWidth
        .WriteProperty "Angle", Mvar.lAngle, m_def_Angle
        .WriteProperty "CenterX", Mvar.Center.lX, UserControl.ScaleWidth / 2
        .WriteProperty "CenterY", Mvar.Center.lY, UserControl.ScaleHeight / 2
        .WriteProperty "ArcSize", Mvar.lArcSize, m_def_ArcSize
        
        ' Handle multi value property 'Gradient areas'
        .WriteProperty "GradientAreas", Mvar.lGradientAreas, 1
        For i = 1 To Mvar.lGradientAreas
            .WriteProperty "GradArea-Color" & i, Mvar.arrGradAreas(i).oColor, vbWhite
            .WriteProperty "GradArea-Position" & i, Mvar.arrGradAreas(i).sngPosition, 1
            .WriteProperty "GradArea-GradGamma" & i, Mvar.arrGradAreas(i).sngGradGamma, 1
        Next i
        
    End With
    
End Sub


Private Sub UserControl_Paint()

    DrawDesign
        
End Sub

Private Sub UserControl_Resize()
    
    Select Case Mvar.Shape
    
        Case A2G_Diamond, A2G_Square, A2G_Circle
                UserControl.Height = UserControl.Width
                
    End Select
    
    Recalc_Dimensions
    Mvar.flgRefreshRegion = True
    DrawDesign

End Sub

Private Sub UserControl_Show()
    
    UserControl.Refresh
    
End Sub




' ***********************
' *  PRIVATE SUBS/FUNCS *
' ***********************

' ------------------------------------------------------------------------------------------------
'    Here all drawing starts
' ------------------------------------------------------------------------------------------------
Private Sub DrawDesign()
    ' This sub dispatches the callings of the wanted design and (in EDIT mode) we put the grab handles onto it.
    
    Dim i       As Long
    Dim k       As Long
    Dim lColor  As Long
    
    
    With Mvar
    
        If .flgIsInitDone = False Then
            ' No drawing until finished!
            
            Exit Sub
        End If
    
        Select Case .Design
    
            Case A2G_CoronaWithLineRight
                    DrawDesign_CoronaWithLine 0
                    
            Case A2G_CoronaWithLineLeft
                    DrawDesign_CoronaWithLine .RadGradient.lWidth / 2
                    
            Case A2G_LinearGradient
                    DrawDesign_LinearGradient
                    
            Case A2G_RadialGradient
                    DrawDesign_RadialGradient
            
            Case A2G_MultiLinGradHor
                    DrawDesign_MultiLinGradHor
            
        End Select
    
        ' In EDIT mode:  Draw the grab handles and interactive design stuff
        If Mvar.flgEditMode = True Then
        
            If .Design = A2G_LinearGradient Then
                ' Draw circle in the middle of screen
                Draw_AA_Circle UserControl.ScaleWidth / 2, UserControl.ScaleHeight / 2, 25, 25, RGB(210, 210, 210), Thick
            End If
        
            ' Why not reusing the gradient filled circles and doing it this (slower) way?
            ' Because of transparency and the shiny reflexion. And speed isn't a problem - we are just in designer mode, remember? ;-)
            For i = 1 To .lGrabHandles
                lColor = 215
                With Mvar.arrGrabHandles(i)
                    
                    ' Here we colorize the grabhandles, default for all unhandled is RED (Else part)
                    For k = 4 To 0 Step -1
                        Select Case .GrabHandleType
                            
                            Case GHT_Color
                                        Draw_AA_Circle .lX, .lY, k, k, RGB(0, lColor, 0), Thick
                            
                            Case GHT_ShapeCorner
                                        Draw_AA_Circle .lX, .lY, k, k, RGB(0, 0, lColor), Thick
                            
                            Case GHT_Gamma
                                        Draw_AA_Circle .lX, .lY, k, k, RGB(lColor, 0, lColor), Thick
                            
                            Case Else
                                        Draw_AA_Circle .lX, .lY, k, k, RGB(lColor, 0, 0), Thick

                        End Select
                        lColor = lColor + 10
                    Next k
                    
                    ' Draw 'Light reflexion point'
                    Draw_AA_Circle .lX - 2, .lY - 2, 1, 1, RGB(240, 240, 240), Thick
                End With
            Next i
        End If
        
        If .flgRefreshRegion = True Then
            RefreshAndAssignClippingRegion
        End If
        
    End With
    
End Sub


Private Sub DrawDesign_MultiLinGradHor()
    
    With Mvar
        PaintMemToDC 0, 0, .LinGradient.lWidth, .LinGradient.lHeight, 0, 0, .LinGradient
    End With
    
End Sub

Private Sub DrawDesign_CoronaWithLine(lRadGradCopyStart As Long)
        
    With Mvar
        ' DRAW GRADIENT CIRCLE
        PaintMemToDC lRadGradCopyStart, 0, .RadGradient.lWidth / 2, .RadGradient.lHeight, _
                .RectRadGradDst.lX, .RectRadGradDst.lY, .RadGradient
        
        
        ' DRAW GRADIENT LINE
        ' Top part: gradient
        PaintLinearGradient .RectLinGradP1, ColorToRGB(UserControl.BackColor), _
                ColorToRGB(UserControl.ForeColor), gdVertical
        ' Middle part: filled
        DrawFilledRect .RectLinGradP2, ColorToRGB(UserControl.ForeColor)
        ' Bottom part: gradient
        PaintLinearGradient .RectLinGradP3, ColorToRGB(UserControl.ForeColor), _
                ColorToRGB(UserControl.BackColor), gdVertical
    End With
    
End Sub

Private Sub DrawDesign_LinearGradient()

    With Mvar
        PaintMemToDC 0, 0, .LinGradient.lWidth, .LinGradient.lHeight, 0, 0, .LinGradient
    End With

End Sub


Private Sub DrawDesign_RadialGradient()
    
    With Mvar
        PaintMemToDC 0, 0, .RadGradient.lWidth, .RadGradient.lHeight, _
                     .Center.lX - .lRadGradWidth / 2, .Center.lY - .lRadGradWidth / 2, _
                     .RadGradient
    End With

End Sub


Private Sub Recalc_Dimensions(Optional flgNoCLS As Boolean)
    
    If flgNoCLS = False Then
        UserControl.Cls
    End If
    With Mvar
        
        Select Case .Design
            
            Case A2G_CoronaWithLineRight        ' Arc to the left, gradient line to the right
                    
                    ' Gradient line
                    .RectLinGradP1.lX = UserControl.ScaleWidth - .lLinGradWidth
                    .RectLinGradP1.lY = 0
                    .RectLinGradP1.lWidth = .lLinGradWidth
                    .RectLinGradP1.lHeight = UserControl.ScaleHeight / 3 + 1
                    
                    .RectLinGradP2.lX1 = .RectLinGradP1.lX
                    .RectLinGradP2.lY1 = .RectLinGradP1.lY + .RectLinGradP1.lHeight
                    .RectLinGradP2.lX2 = .RectLinGradP2.lX1 + .lLinGradWidth
                    .RectLinGradP2.lY2 = .RectLinGradP2.lY1 + .RectLinGradP1.lHeight
                    
                    .RectLinGradP3 = .RectLinGradP1
                    .RectLinGradP3.lY = UserControl.ScaleHeight - .RectLinGradP1.lHeight
                    
                    ' Radial gradient (Corona)
                    .RectRadGradDst.lX = (UserControl.ScaleWidth - .lLinGradWidth) - (.lRadGradWidth / 2)
                    If .RectRadGradDst.lX < 0 Then
                        .RectRadGradDst.lX = 0
                    End If
                    .RectRadGradDst.lY = (UserControl.ScaleHeight - .lRadGradWidth) / 2
                    .RectRadGradDst.lWidth = .lRadGradWidth / 2
                    .RectRadGradDst.lHeight = .lRadGradWidth
                    
                    PrepareRadialGradient .lRadGradWidth, _
                            ColorToRGB(.HiliteColor), ColorToRGB(UserControl.BackColor), _
                            .RadGradient
                    
                                
            Case A2G_CoronaWithLineLeft         ' Arc to the right, gradient line to the left
                    
                    ' Gradient line
                    .RectLinGradP1.lX = 0
                    .RectLinGradP1.lY = 0
                    .RectLinGradP1.lWidth = .lLinGradWidth
                    .RectLinGradP1.lHeight = UserControl.ScaleHeight / 3 + 1
                    
                    .RectLinGradP2.lX1 = .RectLinGradP1.lX
                    .RectLinGradP2.lY1 = .RectLinGradP1.lY + .RectLinGradP1.lHeight
                    .RectLinGradP2.lX2 = .RectLinGradP2.lX1 + .lLinGradWidth
                    .RectLinGradP2.lY2 = .RectLinGradP2.lY1 + .RectLinGradP1.lHeight
                    
                    .RectLinGradP3 = .RectLinGradP1
                    .RectLinGradP3.lY = UserControl.ScaleHeight - .RectLinGradP1.lHeight
                    
                    ' Radial gradient ("Corona")
                    .RectRadGradDst.lX = .lLinGradWidth
                    .RectRadGradDst.lY = (UserControl.ScaleHeight - .lRadGradWidth) / 2
                    .RectRadGradDst.lWidth = .lRadGradWidth / 2
                    .RectRadGradDst.lHeight = .lRadGradWidth
                    
                    PrepareRadialGradient .lRadGradWidth, _
                            ColorToRGB(.HiliteColor), ColorToRGB(UserControl.BackColor), _
                            .RadGradient
            
            
            Case A2G_LinearGradient             ' Simple linear gradient in any angle
                    PrepareLinearGradient UserControl.ScaleWidth, _
                            UserControl.ScaleHeight, _
                            .lAngle, _
                            ColorToRGB(.HiliteColor), ColorToRGB(UserControl.BackColor), _
                            .LinGradient
            
            
            Case A2G_RadialGradient             ' Simple radial gradient with moveable center
                    PrepareRadialGradient .lRadGradWidth, _
                            ColorToRGB(.HiliteColor), _
                            ColorToRGB(UserControl.BackColor), _
                            .RadGradient
            
            Case A2G_MultiLinGradHor            ' Multiple linear adjustable gradients
                    PrepareMultiColorLinearGradient UserControl.ScaleWidth, _
                            UserControl.ScaleHeight, _
                            gdHorizontal, _
                            .LinGradient

                                
        End Select
    
    End With
    
End Sub


Private Sub RaiseError(lErrNo As Long, sErrMsg As String)
    ' Centralized error handling makes changes easy

    RaiseEvent Error(lErrNo, sErrMsg)

End Sub


' ------------------------------------------------------------------------------------------------
'    Here we have Carles P.V. great Universal Speed Gradient Solution from May 2005 on PSC / VB .
'    http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=60477&lngWId=1
'    Thx for sharing this, Carles, because of I really wouldn't be able to do it this way on my own ... ;-)
'
'    My changes: I split the function into a render and a bring-to-dc part and added the
'    possibility to have more than one gradient. The gradients areas got a gamma factor for the
'    "velocity" to swap from origin to destination color.
'    Ok, this way it isn't the best speed thing anymore, but it increases the possibilities and
'    this usercontrol needs it this way.
' ------------------------------------------------------------------------------------------------
Private Sub PrepareMultiColorLinearGradient(ByVal lWidth As Long, _
                                            ByVal lHeight As Long, _
                                            ByVal GradientDirection As enGradientDirection, _
                                            ByRef LinGradient As tpBitmapInMemory)


    Dim lGrad()     As Long
    
    Dim R1          As Long
    Dim G1          As Long
    Dim B1          As Long
    Dim R2          As Long
    Dim G2          As Long
    Dim B2          As Long
    Dim dR          As Long
    Dim dG          As Long
    Dim dB          As Long
    
    Dim Scan        As Long
    Dim i           As Long
    Dim iEnd        As Long
    Dim iOffset     As Long
    Dim j           As Long
    Dim jEnd        As Long
    Dim iGrad       As Long
    Dim sngFactor   As Single
    
    Dim k           As Long
    Dim lAreaStart  As Long
    Dim lAreaEnd    As Long
    Dim lAreaWidth  As Long
    Dim lColorFrom  As Long
    Dim lColorTo    As Long
    Dim lStep       As Long
    
    
    With LinGradient
        '-- A minor check
        If (lWidth < 1 Or lHeight < 1) Then
            
            Exit Sub
        End If
        
        .lWidth = lWidth
        .lHeight = lHeight
        
        '-- Size gradient-colors array
        Select Case GradientDirection
        
            Case [gdHorizontal]
                    ReDim lGrad(0 To .lWidth - 1)
                
            Case [gdVertical]
                    ReDim lGrad(0 To .lHeight - 1)
                    
            Case Else
                    ReDim lGrad(0 To .lWidth + .lHeight - 2)
                    
        End Select
        iEnd = UBound(lGrad())
        
        '-- Calculate gradient-colors
        With Mvar
                        
            ' FROM color - We start with background color
            lColorFrom = ColorToRGB(UserControl.BackColor)
            '-- Decompose color
            lColorFrom = lColorFrom And &HFFFFFF
            R1 = lColorFrom Mod &H100&
            lColorFrom = lColorFrom \ &H100&
            G1 = lColorFrom Mod &H100&
            lColorFrom = lColorFrom \ &H100&
            B1 = lColorFrom Mod &H100&
            
            ' Preset end
            lAreaEnd = -1
            
            For k = 1 To .lGradientAreas
                
                ' On any following gradient area the new 'from' color is the last 'to' color
                If k > 1 Then
                    R1 = R2
                    G1 = G2
                    B1 = B2
                End If
                
                ' TO color
                lColorTo = ColorToRGB(.arrGradAreas(k).oColor)
                '-- Decompose color
                lColorTo = lColorTo And &HFFFFFF
                R2 = lColorTo Mod &H100&
                lColorTo = lColorTo \ &H100&
                G2 = lColorTo Mod &H100&
                lColorTo = lColorTo \ &H100&
                B2 = lColorTo Mod &H100&
                
                '-- Get color distances
                dR = R2 - R1
                dG = G2 - G1
                dB = B2 - B1
                    
                ' -- Get area start and end range in color ramp
                lAreaStart = lAreaEnd + 1
                lAreaEnd = .arrGradAreas(k).sngPosition * iEnd              ' Because of working with % values from 0 to 1
                                                                            ' the calculation within the color ramp and the
                                                                            ' screen position is very easy ;-)
                lAreaWidth = lAreaEnd - lAreaStart + 1
                
                If (iEnd = 0) Then
                    '-- Special case (1-pixel wide gradient)
                    lGrad(0) = ((B1 \ 2 + B2 \ 2)) + _
                               ((G1 \ 2 + G2 \ 2) * 256) + _
                               ((R1 \ 2 + R2 \ 2) * 65536)
                
                ElseIf .arrGradAreas(k).sngGradGamma <> 1 Then
                    ' Gamma <> 1, so we have a change
                    sngFactor = 1 / .arrGradAreas(k).sngGradGamma
                    lStep = -1
                    For i = lAreaStart To lAreaEnd
                        lStep = lStep + 1
                        lGrad(i) = (B1 + (dB * Gamma(lStep, lAreaWidth, sngFactor)) \ lAreaWidth) + _
                                   ((G1 + (dG * Gamma(lStep, lAreaWidth, sngFactor)) \ lAreaWidth) * 256) + _
                                   ((R1 + (dR * Gamma(lStep, lAreaWidth, sngFactor)) \ lAreaWidth) * 65536)
                    Next i
                    
                Else
                    ' Gamma = 1, standard calculation of color gradient
                    lStep = -1
                    For i = lAreaStart To lAreaEnd
                        lStep = lStep + 1
                        lGrad(i) = (B1 + (dB * lStep) \ lAreaWidth) + _
                                   ((G1 + (dG * lStep) \ lAreaWidth) * 256) + _
                                   ((R1 + (dR * lStep) \ lAreaWidth) * 65536)
                    Next i
                End If
            
            Next k
        End With
        
        
        '-- Size DIB array
        ReDim .lArrBitmap(.lWidth * .lHeight - 1) As Long
        iEnd = .lWidth - 1
        jEnd = .lHeight - 1
        Scan = .lWidth
        
        '-- Render gradient DIB
        Select Case GradientDirection
            
            Case [gdHorizontal]
                    For j = 0 To jEnd
                        For i = iOffset To iEnd + iOffset
                            .lArrBitmap(i) = lGrad(i - iOffset)
                        Next i
                        iOffset = iOffset + Scan
                    Next j
            
            Case [gdVertical]
                    For j = jEnd To 0 Step -1
                        For i = iOffset To iEnd + iOffset
                            .lArrBitmap(i) = lGrad(j)
                        Next i
                        iOffset = iOffset + Scan
                    Next j
            
            Case [gdDownwardDiagonal]
                    iOffset = jEnd * Scan
                    For j = 1 To jEnd + 1
                        For i = iOffset To iEnd + iOffset
                            .lArrBitmap(i) = lGrad(iGrad)
                            iGrad = iGrad + 1
                        Next i
                        iOffset = iOffset - Scan
                        iGrad = j
                    Next j
                
            Case [gdUpwardDiagonal]
                    iOffset = 0
                    For j = 1 To jEnd + 1
                        For i = iOffset To iEnd + iOffset
                            .lArrBitmap(i) = lGrad(iGrad)
                            iGrad = iGrad + 1
                        Next i
                        iOffset = iOffset + Scan
                        iGrad = j
                    Next j
                    
        End Select
        
    End With
    
End Sub

Private Function Gamma(lValue As Long, lRange As Long, sngFactor As Single) As Single

    Dim sngVal As Single
    
    sngVal = lValue / lRange
    sngVal = sngVal ^ sngFactor
    Gamma = sngVal * lRange

End Function

Private Sub PaintLinearGradient(RectXYWH As tpRectXYWH, _
                                ByVal Color1 As Long, _
                                ByVal Color2 As Long, _
                                ByVal GradientDirection As enGradientDirection)


    Dim uBIH    As tpAPI_BITMAPINFOHEADER
    Dim lBits() As Long
    Dim lGrad() As Long

    Dim R1      As Long
    Dim G1      As Long
    Dim B1      As Long
    Dim R2      As Long
    Dim G2      As Long
    Dim B2      As Long
    Dim dR      As Long
    Dim dG      As Long
    Dim dB      As Long

    Dim Scan    As Long
    Dim i       As Long
    Dim iEnd    As Long
    Dim iOffset As Long
    Dim j       As Long
    Dim jEnd    As Long
    Dim iGrad   As Long


    With RectXYWH
        '-- A minor check
        If (.lWidth < 1 Or .lHeight < 1) Then

            Exit Sub
        End If

        '-- Decompose colors
        Color1 = Color1 And &HFFFFFF
        R1 = Color1 Mod &H100&
        Color1 = Color1 \ &H100&
        G1 = Color1 Mod &H100&
        Color1 = Color1 \ &H100&
        B1 = Color1 Mod &H100&

        Color2 = Color2 And &HFFFFFF
        R2 = Color2 Mod &H100&
        Color2 = Color2 \ &H100&
        G2 = Color2 Mod &H100&
        Color2 = Color2 \ &H100&
        B2 = Color2 Mod &H100&

        '-- Get color distances
        dR = R2 - R1
        dG = G2 - G1
        dB = B2 - B1

        '-- Size gradient-colors array
        Select Case GradientDirection

            Case [gdHorizontal]
                    ReDim lGrad(0 To .lWidth - 1)

            Case [gdVertical]
                    ReDim lGrad(0 To .lHeight - 1)

            Case Else
                    ReDim lGrad(0 To .lWidth + .lHeight - 2)

        End Select

        '-- Calculate gradient-colors
        iEnd = UBound(lGrad())
        If (iEnd = 0) Then
            '-- Special case (1-pixel wide gradient)
            lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
        Else
            For i = 0 To iEnd
                lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
            Next i
        End If

        '-- Size DIB array
        ReDim lBits(.lWidth * .lHeight - 1) As Long
        iEnd = .lWidth - 1
        jEnd = .lHeight - 1
        Scan = .lWidth

        '-- Render gradient DIB
        Select Case GradientDirection

            Case [gdHorizontal]
                    For j = 0 To jEnd
                        For i = iOffset To iEnd + iOffset
                            lBits(i) = lGrad(i - iOffset)
                        Next i
                        iOffset = iOffset + Scan
                    Next j

            Case [gdVertical]
                    For j = jEnd To 0 Step -1
                        For i = iOffset To iEnd + iOffset
                            lBits(i) = lGrad(j)
                        Next i
                        iOffset = iOffset + Scan
                    Next j

            Case [gdDownwardDiagonal]
                    iOffset = jEnd * Scan
                    For j = 1 To jEnd + 1
                        For i = iOffset To iEnd + iOffset
                            lBits(i) = lGrad(iGrad)
                            iGrad = iGrad + 1
                        Next i
                        iOffset = iOffset - Scan
                        iGrad = j
                    Next j

            Case [gdUpwardDiagonal]
                    iOffset = 0
                    For j = 1 To jEnd + 1
                        For i = iOffset To iEnd + iOffset
                            lBits(i) = lGrad(iGrad)
                            iGrad = iGrad + 1
                        Next i
                        iOffset = iOffset + Scan
                        iGrad = j
                    Next j

        End Select

        '-- Define DIB header
        With uBIH
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = RectXYWH.lWidth
            .biHeight = RectXYWH.lHeight
        End With

        '-- Paint it!
        API_StretchDIBits UserControl.hdc, .lX, .lY, .lWidth, .lHeight, _
                0, 0, .lWidth, .lHeight, _
                lBits(0), uBIH, API_DIB_RGB_COLORS, vbSrcCopy
    End With

End Sub



' ----------------------------------------------------------------------------------------------------
'    Less fames but maybe even better than Carles solution above: His radial gradient sub! Hidden in
'    http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=60580&lngWId=1
'    I was very surprised to discover it and astonished about the easyness of the arc solution.
'    After having a small problem with dimensions of the gradient Carles gave the fastest support
'    you can even think of :-)  Once more: Thank you, Carles for your help!
'
'    For 'ucArt2Gui' I took (a suggestion of Carles in one of his emails) the byte array out of here
'    to modul scope level and split the gradient into a preparing and a drawing part to avoid
'    unneccessary recalculations and to get access to only parts of the gradient circle.
'    'Width' and 'Height' went into 'Size' because of I just need quadratic gradient circles only.
' ----------------------------------------------------------------------------------------------------
'
'    PREVENTED UPDATE
'    Nearly to the end of the development of uaArt2Gui Carles send me a great update which improves
'    speed by factor 2.5 . He gots a fine trick with a squareroot lockup table. Because of the code
'    flow was changed too (and I want to get ready with this project... ;-) ) I decided to don't
'    start from scratch with this part. Because of only ONE circular gradient is drawn,
'    speed isn't the most important thing here.
'    So Carles, can you forgive me? ;-)
'
' ----------------------------------------------------------------------------------------------------
Private Sub PrepareRadialGradient(ByVal Size As Long, _
                                  ByVal Color1 As Long, _
                                  ByVal Color2 As Long, _
                                  ByRef RadGradient As tpBitmapInMemory)

    Dim lGrad()   As Long
    Dim g         As Long
    
    Dim R1        As Long
    Dim G1        As Long
    Dim B1        As Long
    Dim R2        As Long
    Dim G2        As Long
    Dim B2        As Long
    Dim dR        As Long
    Dim dG        As Long
    Dim dB        As Long
    
    Dim Scan      As Long
    Dim iPad      As Long
    Dim jPad      As Long
    Dim i         As Long
    Dim ia        As Long
    Dim iaa       As Long
    Dim j         As Long
    Dim ja        As Long
    Dim jaa       As Long
    Dim iEnd      As Long
    Dim jEnd      As Long
    Dim Offset1   As Long
    Dim Offset2   As Long

    Dim sqr2      As Single
    Dim lQuadSize As Long

    '-- Minor check
    If Size < 1 Then
    
        Exit Sub
    End If
    
    '-- Calc. gradient length ('diagonal')
    lQuadSize = Size * Size
    g = Sqr(2 * lQuadSize) \ 2

    '-- Decompose colors
    R1 = (Color1 And &HFF&)
    G1 = (Color1 And &HFF00&) \ 256
    B1 = (Color1 And &HFF0000) \ 65536
    R2 = (Color2 And &HFF&)
    G2 = (Color2 And &HFF00&) \ 256
    B2 = (Color2 And &HFF0000) \ 65536

    sqr2 = Sqr(2)

    '-- Get color distances
    dR = Int((R2 - R1) * sqr2)
    dG = Int((G2 - G1) * sqr2)
    dB = Int((B2 - B1) * sqr2)

    '-- Size gradient-colors array
    ReDim lGrad(0 To g)

     '-- Calculate gradient-colors
    If (g = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
    Else
        For i = 0 To Int(g / sqr2) - 1
            lGrad(i) = B1 + (dB * i) \ g + 256 * (G1 + (dG * i) \ g) + 65536 * (R1 + (dR * i) \ g)
        Next i
        For i = Int(g / sqr2) To g
            lGrad(i) = B2 + 256 * G2 + 65536 * R2
        Next i
    End If

    With RadGradient
        .lWidth = Size
        .lHeight = Size
        
        '-- Size DIB array
        ReDim .lArrBitmap(lQuadSize - 1) As Long
        
        '== Render gradient DIB
            
        '-- First "quadrant"...
        Scan = Size
        iPad = Scan Mod 2
        jPad = Size Mod 2
    
        iEnd = Scan \ 2 + iPad - 1
        jEnd = Size \ 2 + jPad - 1
        Offset1 = jEnd * Scan + Scan \ 2
    
        ja = 1
        jaa = -1
        For j = 0 To jEnd
            ja = ja + jaa
            ia = ja + 1
            iaa = -1
            For i = Offset1 To Offset1 + iEnd
                ia = ia + iaa
                .lArrBitmap(i) = lGrad(Sqr(ia))
                iaa = iaa + 2
            Next i
            jaa = jaa + 2
            Offset1 = Offset1 - Scan
        Next j
    
        '-- Mirror first "quadrant"
        iEnd = iEnd - iPad
        Offset1 = 0
        Offset2 = Scan - 1
    
        For j = 0 To jEnd
            For i = 0 To iEnd
                .lArrBitmap(Offset1 + i) = .lArrBitmap(Offset2 - i)
            Next i
            Offset1 = Offset1 + Scan
            Offset2 = Offset2 + Scan
        Next j
    
        '-- Mirror first "half"
        iEnd = Scan - 1
        jEnd = jEnd - jPad
        Offset1 = (Size - 1) * Scan
        Offset2 = 0
    
        For j = 0 To jEnd
            For i = 0 To iEnd
                .lArrBitmap(Offset1 + i) = .lArrBitmap(Offset2 + i)
            Next i
            Offset1 = Offset1 - Scan
            Offset2 = Offset2 + Scan
        Next j
    End With
    
End Sub

Private Sub PaintMemToDC(ByVal lSrcX As Long, _
                                ByVal lSrcY As Long, _
                                ByVal lWidth As Long, _
                                ByVal lHeight As Long, _
                                ByVal lDstX As Long, _
                                ByVal lDstY As Long, _
                                ByRef MemBitmap As tpBitmapInMemory)
    
    Dim uBIH As tpAPI_BITMAPINFOHEADER
    
    
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = MemBitmap.lWidth
        .biHeight = MemBitmap.lHeight
        .biCompression = 0
    End With

    '-- Paint it!
    API_StretchDIBits UserControl.hdc, _
            lDstX, lDstY, lWidth, lHeight, _
            lSrcX, lSrcY, lWidth, lHeight, _
            MemBitmap.lArrBitmap(0), uBIH, _
            API_DIB_RGB_COLORS, vbSrcCopy

End Sub


Private Sub DrawFilledRect(RectXYWH As tpRectX1Y1X2Y2, lColor As Long)
                            
    Dim lBrush  As Long
    
    lBrush = API_CreateSolidBrush(lColor)
    API_FillRect UserControl.hdc, RectXYWH, lBrush
    API_DeleteObject lBrush
    
End Sub

Private Sub PrepareLinearGradient(ByVal Width As Long, _
                                  ByVal Height As Long, _
                                  ByVal Angle As Single, _
                                  ByVal Color1 As Long, _
                                  ByVal Color2 As Long, _
                                  ByRef LinGradient As tpBitmapInMemory)
    
    Const INT_ROT As Long = 1000&   ' Increase this value for more precision
    
    Dim lGrad()   As Long
    
    Dim lClr      As Long
    Dim R1        As Long
    Dim G1        As Long
    Dim B1        As Long
    Dim R2        As Long
    Dim G2        As Long
    Dim B2        As Long
    Dim dR        As Long
    Dim dG        As Long
    Dim dB        As Long
    
    Dim Scan      As Long
    Dim i         As Long
    Dim j         As Long
    Dim iIn       As Long
    Dim jIn       As Long
    Dim iEnd      As Long
    Dim jEnd      As Long
    Dim Offset    As Long
    
    Dim lQuad     As Long
    Dim AngleDiag As Single
    Dim AngleComp As Single
    
    Dim g         As Long
    Dim luSin     As Long
    Dim luCos     As Long
 
    
    '-- Minor check
    If (Width > 0 And Height > 0) Then
        
        '-- Right-hand [+] (ox=0º)
        Angle = -Angle + 90
        
        '-- Normalize to [0º;360º]
        Angle = Angle Mod 360
        If (Angle < 0) Then Angle = 360 + Angle
        
        '-- Get quadrant (0 - 3)
        lQuad = Angle \ 90
        
        '-- Normalize to [0º;90º]
        Angle = Angle Mod 90
        
        '-- Calc. gradient length ('distance')
        If (lQuad Mod 2 = 0) Then
            AngleDiag = Atn(Width / Height) * TO_DEG
        Else
            AngleDiag = Atn(Height / Width) * TO_DEG
        End If
        AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
        Angle = Angle * TO_RAD
        g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem
        
        '-- Decompose colors
        If (lQuad > 1) Then
            lClr = Color1
            Color1 = Color2
            Color2 = lClr
        End If
        
        R1 = (Color1 And &HFF&)
        G1 = (Color1 And &HFF00&) \ 256
        B1 = (Color1 And &HFF0000) \ 65536
        
        R2 = (Color2 And &HFF&)
        G2 = (Color2 And &HFF00&) \ 256
        B2 = (Color2 And &HFF0000) \ 65536
        
        '-- Get color distances
        dR = R2 - R1
        dG = G2 - G1
        dB = B2 - B1
        
        '-- Size gradient-colors array
        ReDim lGrad(0 To g - 1)
        
         '-- Calculate gradient-colors
        iEnd = g - 1
        If (iEnd = 0) Then
            '-- Special case (1-pixel wide gradient)
            lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
          Else
            For i = 0 To iEnd
                lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
            Next i
        End If
        
        '-- Size DIB array
        ReDim LinGradient.lArrBitmap(Width * Height - 1) As Long
        LinGradient.lWidth = Width
        LinGradient.lHeight = Height
        
        '-- Render gradient DIB
        iEnd = Width - 1
        jEnd = Height - 1
        
        Select Case lQuad
                    
            Case 0, 2
                    luSin = Sin(Angle) * INT_ROT
                    luCos = Cos(Angle) * INT_ROT
                    Offset = 0
                    Scan = Width
                
            Case 1, 3
                    luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
                    luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
                    Offset = jEnd * Width
                    Scan = -Width
        
        End Select
        
        jIn = 0
        iIn = 0
        For j = 0 To jEnd
            iIn = jIn
            For i = 0 To iEnd
                LinGradient.lArrBitmap(i + Offset) = lGrad(iIn \ INT_ROT)
                iIn = iIn + luSin
            Next i
            jIn = jIn + luCos
            Offset = Offset + Scan
        Next j
    
    End If

End Sub


Private Function ColorToRGB(ByVal oColor As OLE_COLOR) As Long
    ' Convert color values from OLE representation to RGB representation
    
    Dim lRGB    As Long
    Dim hPal    As Long
    
    ColorToRGB = IIf(API_OleTranslateColor(oColor, hPal, lRGB), API_INVALID_COLOR, lRGB)
    
End Function


Private Function ColorToRGBwithShading(ByVal oColor As OLE_COLOR, lOffset As Long) As Long
    ' Convert color values from OLE representation to RGB representation and
    ' add an offset to get a shading
    
    Const SUCCESS = 0
    
    Dim lRGB            As Long
    Dim hPal            As Long
    Dim R               As Long
    Dim g               As Long
    Dim B               As Long
    
    ColorToRGBwithShading = API_INVALID_COLOR
    If API_OleTranslateColor(oColor, hPal, lRGB) = SUCCESS Then
        
        R = (lRGB And &HFF&) + lOffset
        g = (lRGB And &HFF00&) \ 256 + lOffset
        B = (lRGB And &HFF0000) \ 65536 + lOffset
        
        If R < 0 Then
            R = 0
        ElseIf R > 255 Then
            R = 255
        End If
        
        If g < 0 Then
            g = 0
        ElseIf g > 255 Then
            g = 255
        End If
        
        If B < 0 Then
            B = 0
        ElseIf B > 255 Then
            B = 255
        End If
        
        ColorToRGBwithShading = RGB(R, g, B)
    End If
    
End Function


Private Function SelectFromStdColorDlg(Optional PresetColor As OLE_COLOR) As OLE_COLOR

    Const SUCCESS As Long = 1&

    Dim cc                          As tpAPI_CHOOSECOLORSTRUCT
    Dim lArrCustomColors(0 To 15)   As Long                         ' Presets for 16 custom colors

    With cc
        .flags = CC_ANYCOLOR Or CC_FULLOPEN Or CC_RGBINIT
        .rgbResult = PresetColor
        .lStructSize = Len(cc)
        .hwndOwner = UserControl.hWnd
        .lpCustColors = VarPtr(lArrCustomColors(0))
    
        If API_ChooseColor(cc) = SUCCESS Then
            SelectFromStdColorDlg = cc.rgbResult
        Else
            SelectFromStdColorDlg = API_INVALID_COLOR
        End If
    End With
    
End Function


Private Function PutDotOnCircle(ByVal lAngle As Long, lDistanceToCenter As Long, ByRef RETURN_Position As tpGrabHandle)
    ' Takes the center of a circle, an angle and a distance and calculates the position on the circle
        
    Dim lX As Long
    Dim lY As Long
    
    lX = UserControl.ScaleWidth / 2
    lY = UserControl.ScaleHeight / 2
    
    RETURN_Position.lX = lX + lDistanceToCenter * Cos((360 - lAngle) * TO_RAD)
    RETURN_Position.lY = lY + lDistanceToCenter * Sin((360 - lAngle) * TO_RAD)

End Function


Private Function PutDotOnRectangle(ByVal lAngle As Long, lGapToUcEdge As Long, ByRef RETURN_Position As tpGrabHandle)
    
    Dim lX1     As Long
    Dim lY1     As Long
    Dim lX2     As Long
    Dim lY2     As Long
    Dim lLenght As Long
    Dim dblM    As Double
    Dim dblB    As Double
    
    lAngle = IIf(lAngle > 360, lAngle - 360, lAngle)
    With UserControl
        ' Center point
        lX1 = .ScaleWidth / 2
        lY1 = .ScaleHeight / 2
        
        ' Distance to 2nd point long enough to ensure cutting the edges of the uc
        lLenght = IIf(.ScaleHeight > .ScaleWidth, .ScaleHeight, .ScaleWidth)
    
        ' Destination point
        lX2 = lX1 + lLenght * Cos((360 - lAngle) * TO_RAD)
        lY2 = lY1 + lLenght * Sin((360 - lAngle) * TO_RAD)

        Select Case lAngle
                                
                Case Is <= 45, Is > 315
                        RETURN_Position.lX = .ScaleWidth - lGapToUcEdge
                        dblM = (lY2 - lY1) / (lX2 - lX1)
                        dblB = lY1 - (dblM * lX1)
                        RETURN_Position.lY = dblM * RETURN_Position.lX + dblB
                        
                Case 46 To 135
                        RETURN_Position.lY = lGapToUcEdge
                        If lX1 = lX2 Then
                            RETURN_Position.lX = lX1
                        Else
                            dblM = (lY2 - lY1) / (lX2 - lX1)
                            dblB = lY1 - (dblM * lX1)
                            RETURN_Position.lX = (RETURN_Position.lY - dblB) / dblM
                        End If

                Case 136 To 225
                        RETURN_Position.lX = lGapToUcEdge
                        dblM = (lY2 - lY1) / (lX2 - lX1)
                        dblB = lY1 - (dblM * lX1)
                        RETURN_Position.lY = dblM * RETURN_Position.lX + dblB

                Case 226 To 315
                        RETURN_Position.lY = .ScaleHeight - lGapToUcEdge
                        If lX1 = lX2 Then
                            RETURN_Position.lX = lX1
                        Else
                            dblM = (lY2 - lY1) / (lX2 - lX1)
                            dblB = lY1 - (dblM * lX1)
                            RETURN_Position.lX = (RETURN_Position.lY - dblB) / dblM
                        End If
                        
            End Select
        

    End With

End Function


Private Sub ChangeColor(iButton As Integer, lDeltaX As Long, IndxColor As enColor)
    ' Here we handle the interactive change of a color
    
    Dim lNewColor As Long
    

    lNewColor = API_INVALID_COLOR
    With Mvar
        
        ' When grabber moved with LEFT button pressed adjust color directly
        If iButton = 1 And lDeltaX <> 0 Then
            lNewColor = ColorToRGBwithShading(.arrOriginalCol(IndxColor), lDeltaX)
    
        ' On RIGHT button show standard color selector to change the color
        ElseIf iButton = 2 Then
            lNewColor = SelectFromStdColorDlg(.arrOriginalCol(IndxColor))
            
        End If
                
        ' Now SET the new color
        If lNewColor <> API_INVALID_COLOR Then
            If IndxColor = COLOR_BACKGROUND Then
                Me.BackColor = lNewColor
                
            ElseIf IndxColor = COLOR_FOREGROUND Then
                Me.ForeColor = lNewColor
                
            ElseIf IndxColor = COLOR_HILITE Then
                Me.HiliteColor = lNewColor
            
            ElseIf IndxColor = COLOR_AREA Then
                Me.GradAreaColor((.CurrMovingGrabber + 2) \ 3) = lNewColor
                
            End If
        End If
    End With

End Sub


Private Sub ChangeGamma(iButton As Integer, lDeltaX As Long)
    ' Here we handle the interactive change of the gradient "velocity" with a gamma factor
    
    Dim sngNewGamma As Single
    
    With Mvar
        ' When grabber moved with LEFT button pressed adjust gamma directly
        If iButton = 1 And lDeltaX <> 0 Then
            sngNewGamma = 1 + (lDeltaX * 0.01)
            If sngNewGamma < 0.01 Then
                sngNewGamma = 0.01
            End If
            Me.GradAreaGamma((.CurrMovingGrabber + 2) \ 3) = sngNewGamma
        End If
    End With

End Sub


Private Sub ChangeAreaPosition(iButton As Integer, ByVal lX As Long)
    ' Here we handle the interactive change of the 'end' value gradient "velocity" with a gamma factor
    
    Dim sngNewPosition  As Single
    Dim sngScaledGHS2   As Single
    Dim lAreaNo         As Long
    
    With Mvar
        lAreaNo = (.CurrMovingGrabber + 2) \ 3
        sngScaledGHS2 = GRAB_HANDLE_SIZE / UserControl.ScaleWidth * 2
        ' When grabber moved with LEFT button pressed adjust position directly
        If iButton = 1 And lAreaNo < .lGradientAreas Then       ' Last area end cannot be moved!
            
            sngNewPosition = (lX + GRAB_HANDLE_SIZE) / UserControl.ScaleWidth
            
            ' Check left edge
            If sngNewPosition < sngScaledGHS2 Then
                sngNewPosition = sngScaledGHS2
            
            ' Check right edge
            ElseIf sngNewPosition > 1 Then
                sngNewPosition = 1
            
            End If
            
            If lAreaNo > 1 Then
                ' Check to left next area
                If sngNewPosition < Me.GradAreaPosition(lAreaNo - 1) + sngScaledGHS2 Then
                    sngNewPosition = Me.GradAreaPosition(lAreaNo - 1) + sngScaledGHS2
                End If
            End If
            
            ' Check to right next area
            If sngNewPosition > Me.GradAreaPosition(lAreaNo + 1) - sngScaledGHS2 Then
                sngNewPosition = Me.GradAreaPosition(lAreaNo + 1) - sngScaledGHS2
                
            End If
            
            Me.GradAreaPosition(lAreaNo) = sngNewPosition
        End If
    End With

End Sub


'   Dana Seaman's code to draw anti-aliased circles, a little bit modified by me. Released on PSC VB in January 2002 at
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=38582&lngWId=1
Private Sub Draw_AA_Circle(ByVal lCenterX As Long, _
                           ByVal lCenterY As Long, _
                           ByVal lRadiusX As Long, _
                           ByVal lRadiusY As Long, _
                           ByVal lColorRGB As Long, _
                  Optional ByVal Thickness As cThickness = Thick)
     
    Dim Bbg                 As Byte
    Dim Gbg                 As Byte
    Dim Rbg                 As Byte
    Dim savAlpha(1 To 4)    As Byte
    Dim Bblend              As Long
    Dim Bgr                 As Long
    Dim Cl                  As Long
    Dim Gblend              As Long
    Dim Strength            As Long
    Dim StrengthI           As Long
    Dim Quadrant            As Long
    Dim Radius              As Long
    Dim Rblend              As Long
    Dim RX1                 As Long
    Dim RX2                 As Long
    Dim RY1                 As Long
    Dim RY2                 As Long
    Dim savX(1 To 4)        As Long
    Dim savY(1 To 4)        As Long
    Dim X4                  As Long
    Dim Y4                  As Long
    Dim NewColor            As Long
    Dim Red                 As Long
    Dim Green               As Long
    Dim Blue                As Long
    Dim Ax                  As Single
    Dim Ay                  As Single
    Dim Bx                  As Single
    Dim By                  As Single
    Dim L1                  As Single
    Dim L2                  As Single
    Dim L3                  As Single
    Dim L4                  As Single
    Dim sngAngle            As Single
    Dim sngPointSpacing     As Single
    Dim X2                  As Single
    Dim Xp5                 As Single
    Dim Y2                  As Single
    Dim sngLS               As Single
    
    If lColorRGB Then
        Red = lColorRGB And &HFF&
        Green = (lColorRGB And &HFF00&) \ 256
        Blue = (lColorRGB And &HFF0000) \ 65536
    Else 'Color is Black
        Red = 0
        Green = 0
        Blue = 0
    End If
    
    Radius = lRadiusX
    If lRadiusY > lRadiusX Then
        Radius = lRadiusY
    End If

    sngLS = IIf(Thickness = Thick, cThick, cThin)
    
    If Radius < 0 Then
        sngPointSpacing = -sngLS / Radius
    ElseIf Radius = 0 Then
        sngPointSpacing = sngLS
    Else
        sngPointSpacing = sngLS / Radius
    End If

    For sngAngle = 0 To HalfPi Step sngPointSpacing
        X2 = lRadiusX * Cos(sngAngle)
        Y2 = lRadiusY * Sin(sngAngle)
        'Prevents error when vb rounds .5 down
        If X2 = Int(X2) Then X2 = X2 + 0.001
        If Y2 = Int(Y2) Then Y2 = Y2 + 0.001
        For Quadrant = 0 To 3
            
            Select Case Quadrant
                Case 0 '0-90°
                        Ax = X2 + lCenterX - 0.5
                        Ay = -Y2 + lCenterY - 0.5
                
                Case 1 '90-180°
                        Ax = X2 + lCenterX - 0.5
                        Ay = Y2 + lCenterY - 0.5
                
                Case 2 '180-270°
                        Ax = -X2 + lCenterX - 0.5
                        Ay = Y2 + lCenterY - 0.5
                
                Case 3 '270-360°
                        Ax = -X2 + lCenterX - 0.5
                        Ay = -Y2 + lCenterY - 0.5
            
            End Select
            
            Bx = Ax + 1
            By = Ay + 1
            RX1 = Ax
            RX2 = RX1 + 1
            Xp5 = RX1 + 0.5
            RY1 = Ay
            RY2 = By
            L1 = RY1 + 0.5 - Ay
            L2 = 256 * (Xp5 - Ax) - Xp5 + Ax
            L3 = 255 - L2
            L4 = By - RY2 + 0.5
            savX(1) = RX1
            savY(1) = RY1
            savX(2) = RX2
            savY(2) = RY1
            savY(3) = RY2
            savX(3) = RX1
            savY(4) = RY2
            savX(4) = RX2
            savAlpha(1) = L1 * L2
            savAlpha(2) = L1 * L3
            savAlpha(3) = L4 * L2
            savAlpha(4) = L4 * L3
            
            For Cl = 1 To 4
                Strength = savAlpha(Cl)
                X4 = savX(Cl)
                Y4 = savY(Cl)
                If Strength > 252 Then                              ' If > 99%
                    API_SetPixelV UserControl.hdc, X4, Y4, lColorRGB
                Else
                    Bgr = API_GetPixel(UserControl.hdc, X4, Y4)
                    If Bgr Then                                     ' If not black
                        Rbg = Bgr And &HFF&
                        Gbg = (Bgr And &HFF00&) \ &H100&
                        Bbg = (Bgr And &HFF0000) \ &H10000
                    Else
                        Rbg = 0
                        Gbg = 0
                        Bbg = 0
                    End If
                    StrengthI = 255 - Strength
                    Rblend = StrengthI * Rbg + Strength * Red
                    Gblend = StrengthI * Gbg + Strength * Green
                    Bblend = StrengthI * Bbg + Strength * Blue
                    NewColor = RGB(Rblend \ 256, Gblend \ 256, Bblend \ 256)
                    API_SetPixelV UserControl.hdc, X4, Y4, NewColor
                End If
            Next Cl
        Next Quadrant
    Next sngAngle

End Sub


Private Function GetAngle(ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long) As Long
    
    Dim dblResult As Double

    If lX2 = lX1 Then
        If lY2 > lY1 Then GetAngle = 90: Exit Function
        If lY2 < lY1 Then GetAngle = 270
        Exit Function
    End If
    
    If lY2 = lY1 Then
        If lX1 > lX2 Then GetAngle = 180
        Exit Function
    End If
    
    dblResult = Atn((lY2 - lY1) / (lX2 - lX1))
    
    If lX2 < lX1 Then dblResult = dblResult + PI
    
    GetAngle = dblResult / TO_RAD

End Function


Private Function GetDistance(ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long) As Long

    Dim sngDx As Single
    Dim sngDy As Single

    sngDx = lX2 - lX1
    sngDy = lY2 - lY1
    
    GetDistance = Sqr(sngDx * sngDx + sngDy * sngDy)

End Function



' ############################################################
' ###    FROM HERE WE HAVE ALL THE CLIPPING REGION STUFF   ###
' ############################################################

Private Sub RefreshAndAssignClippingRegion()
    '
    ' End of January 2005 Keith 'LaVolpe' Fox wrote on PSC VB:
    '
    ' "A what-if project that took off but really can't use the code for anything
    '  right now -- hope you can find a good home for it :) "
    '
    ' Okay, Keith: Here is my suggestion to give it an appropriately home ;-)
    '
    ' MANY thanks for your great ideas, the many days and hourswork you put in and for your nice little FYi ;-) !
    
    '
    ' A note to uncommented code:  Couldn't get Keith' solution fixed right now to draw
    ' the border in really ANY cases. Borders are not drawn when started compiled and are
    ' drawn in Edit mode when increasing the usercontrol by keyboard (Ctrl + Cursor).
    ' So this part is out to 'commen' right now ...
    
    
    Dim RgnClip As Long
'''    Dim tRgnTL  As Long         ' These regions are used to create & draw the borders
'''    Dim tRgnBR  As Long
'''    Dim tRgn    As Long
'''    Dim hBrushL As Long         ' Brushes used to draw the borders
'''    Dim hBrushR As Long
    Dim newCx   As Long         ' The ultimate size, modified if needed, of the new shape
    Dim newCy   As Long
    
        
    With Mvar
        
        If .Shape = A2G_Rectangle Then                              ' No clipping needed on a simple rectangle.
            RemoveClippingRegion
            
            Exit Sub
        End If
        
'''        ' Create two brushes for the borders
'''        hBrushL = API_CreateSolidBrush(vbWhite)
'''        hBrushR = API_CreateSolidBrush(RGB(64, 64, 64))

        ' Calculate the requested shape size
        newCx = UserControl.ScaleWidth
        newCy = UserControl.ScaleHeight

        Select Case .Shape
            
            Case A2G_Rectangle
                    RemoveClippingRegion
                    
            Case A2G_Curved_Rectangle
                    RgnClip = CreateRoundedRectangle(newCx, newCy)
                
            Case A2G_Hexagon, A2G_Diamond
                    ' Note: by passing equal cx & cy, a perfect diamond is drawn
                    RgnClip = CreateHexRegion(newCx, newCy)
                
            Case A2G_Octagon
                    RgnClip = CreateOctRegion(newCx, newCy)
    
            Case A2G_Diagonal_Rectangle
                    RgnClip = CreateDiagRectRegion(newCx, newCy, .ShapeLT, .ShapeRB)
    
            Case A2G_Circle, A2G_Ellipse
                    RgnClip = CreateEllipticRegion(newCx, newCy)
                
        End Select


'''        ' Drawing borders on shaped regions isn't exactly easy. This little algorithm could
'''        ' probably be used on very complicated shapes also ...
'''
'''        newCx = newCx - 1
'''        newCy = newCy - 1
'''
'''        ' Do the left & top border first. Create two rectangular regions of the shaped size:
'''        tRgnTL = API_CreateRectRgn(0, 0, newCx, newCy)
'''        tRgn = API_CreateRectRgn(0, 0, newCx, newCy)
'''
'''        ' Shift the new region left one to catch the left side
'''        API_OffsetRgn tRgnTL, -1, 0
'''        API_CombineRgn tRgnTL, tRgnTL, RgnClip, RGN_XOR
'''        API_OffsetRgn tRgnTL, 1, 0
'''
'''        ' Now using the temp region, shift it up one to catch the top side
'''        API_OffsetRgn tRgn, -1, -1
'''        API_CombineRgn tRgn, tRgn, RgnClip, RGN_XOR
'''        API_OffsetRgn tRgn, 1, 1
'''
'''        ' Add this to the new region & complete the left/top borders
'''        API_CombineRgn tRgnTL, tRgn, tRgnTL, RGN_OR
'''        API_DeleteObject tRgn
'''
'''        ' Do the same for the bottom & right borders
'''        tRgnBR = API_CreateRectRgn(0, 0, newCx + 0, newCy + 0)
'''        tRgn = API_CreateRectRgn(0, 0, newCx + 0, newCy + 0)
'''        API_OffsetRgn tRgnBR, 1, 0
'''        API_CombineRgn tRgnBR, tRgnBR, RgnClip, RGN_XOR
'''        API_OffsetRgn tRgnBR, -1, 0
'''        API_OffsetRgn tRgn, 1, 1
'''        API_CombineRgn tRgn, tRgn, RgnClip, RGN_XOR
'''        API_OffsetRgn tRgn, -1, -1
'''        API_CombineRgn tRgnBR, tRgn, tRgnBR, RGN_OR
'''        API_DeleteObject tRgn

        ' Apply the new window region & don't delete region; windows now owns it!
        API_SetWindowRgn UserControl.hWnd, RgnClip, True

'''        ' Draw the borders.
'''        API_FrameRgn UserControl.hdc, tRgnTL, hBrushL, 1, 1
'''        API_FrameRgn UserControl.hdc, tRgnBR, hBrushR, 1, 1
       
        ' Delete the regions and brushes
'''        API_DeleteObject tRgnBR
'''        API_DeleteObject tRgnTL
'''        API_DeleteObject hBrushL
'''        API_DeleteObject hBrushR
        API_DeleteObject RgnClip
        
        .flgRefreshRegion = False
        
    End With

End Sub

Private Function CreateHexRegion(cx As Long, cy As Long) As Long
    ' Function creates a horizontal/vertical hexagon region with perfectly smooth edges.
    ' The cx & cy parameters are the respective width & height of the region.
    ' A diamond will be created by passing cx & cy as equal values.

    Dim tpts(0 To 7) As tpPoint
    
    If cy > cx Then                         ' Calculate the vertical hex. A top layer, middle layer & bottom layer
        If cx < 4 Then cx = 4               ' Absolute minimum width & height of a hex region
        If cx Mod 2 Then cx = cx - 1        ' Ensure width is even width.
        
        tpts(0).lX = cx \ 2                 ' bot apex
        tpts(0).lY = cy
        tpts(1).lX = cx                     ' bot right
        tpts(1).lY = cy - tpts(0).lX
        tpts(2).lX = cx                     ' top right
        tpts(2).lY = tpts(0).lX - 1
        tpts(3).lX = tpts(0).lX             ' top apex
        tpts(3).lY = -1
        tpts(4).lX = tpts(0).lX - 1         ' added                         ' Add an extra point and modify. Trial & error
        tpts(4).lY = 0                                                      ' shows without this added point, getting a nice
        tpts(5).lX = 0                      ' top left                      ' smooth diagonal edge is impossible.
        tpts(5).lY = tpts(2).lY
        tpts(6).lX = 0                      ' bot left
        tpts(6).lY = tpts(1).lY
        tpts(7) = tpts(0)                   ' bot apex, close polygon
        
    Else                                    ' Calculate the horizontal hex. A left layer, middle layer & right layer
        If cy < 4 Then cy = 4               ' Absolute minimum width & height of a hex region
        If cy Mod 2 Then cy = cy - 1        ' Ensure height is odd since hex requires 3 layers.
    
        tpts(0).lX = 0                      ' left apex
        tpts(0).lY = cy \ 2
        tpts(1).lX = tpts(0).lY             ' bot left
        tpts(1).lY = cy
        tpts(2).lX = cx - tpts(0).lY        ' bot right
        tpts(2).lY = tpts(1).lY
        tpts(3).lX = cx                     ' right apex
        tpts(3).lY = tpts(0).lY
        tpts(4).lX = cx                                                     ' Add an extra point and modify. Trial & error
        tpts(4).lY = tpts(3).lY - 1                                         ' shows without this added point, getting a nice
        tpts(5).lX = tpts(2).lX + 1         ' top right                     ' smooth diagonal edge is impossible.
        tpts(5).lY = 0
        tpts(6).lX = tpts(1).lX - 1         ' top left
        tpts(6).lY = 0
        tpts(7).lX = tpts(0).lX             ' left apex, close polygon
        tpts(7).lY = tpts(0).lY - 1
        
    End If

    CreateHexRegion = API_CreatePolygonRgn(tpts(0), UBound(tpts) + 1, 2)

End Function


Private Function CreateOctRegion(cx As Long, cy As Long) As Long
    ' Function returns a handle to an octagonal region.
    ' The cx & cy parameters are the respective width & height of the region.

    Dim tpts(0 To 8) As tpPoint


    If cx < cy Then                         ' Vertical
        If cx < 4 Then cx = 4               ' Absolute minimum width & height of an octagon region
        If cx Mod 2 Then cx = cx - 1        ' Ensure height is even since oct requires 4 layers: A left, 2 middle and aright layer
        
    Else
        If cy < 4 Then cy = 4               ' Absolute minimum width & height of a octagon region
        If cy Mod 2 Then cy = cy - 1        ' Ensure height is even since oct requires 4 layers: A top, 2 middle and a bottom layer
    
    End If
        
    ' Calculate the octagon
    ' Note: Created regions do not include the right & bottom edges by design.
    
    ' Common points
    tpts(0).lY = cy
    tpts(1).lY = cy
    tpts(2).lX = cx                         ' mid bot right
    tpts(3).lX = cx                         ' mid top right
    tpts(8).lY = cy
    
    ' Different points
    If cx < cy Then                         ' Vertical
        tpts(0).lX = cx \ 4 + 1             ' bot left
        tpts(1).lX = cx - cx \ 4 - 1        ' bot right
        tpts(2).lY = cy - cx \ 4 - 1
        tpts(3).lY = cx \ 4
        tpts(4).lX = tpts(1).lX + 1         ' top right
        tpts(5).lX = cx \ 4                 ' top left
        tpts(6).lY = tpts(3).lY
        tpts(7).lY = tpts(2).lY
        tpts(8).lX = tpts(0).lX             ' bot left
        
    Else
        tpts(0).lX = cy \ 4 + 1             ' bot left
        tpts(1).lX = cx - cy \ 4 - 1        ' bot right
        tpts(2).lY = cy - cy \ 4 - 1
        tpts(3).lY = cy \ 4
        tpts(4).lX = tpts(1).lX + 1         ' top right
        tpts(5).lX = cy \ 4                 ' top left
        tpts(6).lY = tpts(3).lY
        tpts(7).lY = tpts(2).lY
        tpts(8).lX = tpts(0).lX             ' bot left
        
    End If
        
    CreateOctRegion = API_CreatePolygonRgn(tpts(0), UBound(tpts) + 1, 2)

End Function


Private Function CreateDiagRectRegion(cx As Long, _
                                      cy As Long, _
                                      SideAStyle As enShapeLeftTop, _
                                      SideBStyle As enShapeRightBottom) As Long

    ' The cx & cy parameters are the respective width & height of the region.
    ' SideAStyle is -1, 0 or 1 depending on horizontal/vertical shape, reflects the left or top side of the region
    '    -1 draws left/top edge like  /
    '     0 draws left/top edge like  |
    '     1 draws left/top edge like  \
    ' SideBStyle is -1, 0 or 1 depending on horizontal/vertical shape, reflects the right or bottom side of the region
    '    -1 draws right/bottom edge like  \
    '     0 draws right/bottom edge like  |
    '     1 draws right/bottom edge like  /
    
    
    Dim tpts(0 To 4) As tpPoint
    
    If cx > cy Then                                             ' Horizontal
    
        ' Absolute minimum width & height of a octagon region
        If Abs(SideAStyle + SideBStyle) = 2 Then                ' Has 2 opposing slanted sides
            If cx < cy * 2 Then cy = cx \ 2
            
        ElseIf SideAStyle > 0 Or SideBStyle > 0 Then            ' Has one slanted side
            If cx < cy Then cy = cx
        End If
        
        If SideAStyle < 0 Then
            tpts(0).lX = cy - 1
            tpts(1).lX = -1
            
        ElseIf SideAStyle > 0 Then
            tpts(1).lX = cy
            
        End If
        tpts(1).lY = cy
        tpts(2).lX = cx + Abs(SideBStyle < 0)
        If SideBStyle > 0 Then tpts(2).lX = tpts(2).lX - cy
        tpts(2).lY = cy
        tpts(3).lX = cx + Abs(SideBStyle < 0)
        If SideBStyle < 0 Then tpts(3).lX = tpts(3).lX - cy
    
    Else
    
        ' Absolute minimum width & height of a octagon region
        If Abs(SideAStyle + SideBStyle) = 2 Then                ' Has 2 opposing slanted sides
            If cy < cx * 2 Then cx = cy \ 2
            
        ElseIf SideAStyle > 0 Or SideBStyle > 0 Then            ' Has one slanted side
            If cy < cx Then cx = cy
            
        End If
        
        If SideAStyle < 0 Then
            tpts(0).lY = cx - 1
            tpts(3).lY = -1
            
        ElseIf SideAStyle > 0 Then
            tpts(3).lY = cx - 1
            tpts(0).lY = -1
            
        End If
        
        tpts(1).lY = cy
        If SideBStyle < 0 Then tpts(1).lY = tpts(1).lY - cx
        tpts(2).lX = cx
        tpts(2).lY = cy
        If SideBStyle > 0 Then tpts(2).lY = tpts(2).lY - cx
        tpts(3).lX = cx
    
    End If
    tpts(4) = tpts(0)
       
    CreateDiagRectRegion = API_CreatePolygonRgn(tpts(0), UBound(tpts) + 1, 2)

End Function

Private Function CreateEllipticRegion(lWidth As Long, lHeight As Long) As Long
    ' Ellipse or circle region
    
    CreateEllipticRegion = API_CreateEllipticRgn(0&, 0&, lWidth - 1, lHeight - 1)

End Function


Private Function CreateRoundedRectangle(lWidth As Long, lHeight As Long) As Long
    
    CreateRoundedRectangle = API_CreateRoundRectRgn(1&, 1&, lWidth - 2, lHeight - 2, Mvar.lArcSize, Mvar.lArcSize)

End Function


Private Sub RemoveClippingRegion()
    ' Take away any region clipping from the usercontrol

    API_SetWindowRgn UserControl.hWnd, 0&, True
    
End Sub


Private Function LocalProofGetSingle(sVal As String) As Single
    ' Same as VB's  Val(aString), but works with both versions of representation  12,567 (e.g. German)  and 12.567 (e.g. US)
    ' WHY?  Needed, because of German version of VB saves single property values in .frm files localized.

    Dim lPos As Long
    
    If Len(sVal) = 0 Then Exit Function
    
    lPos = InStr(sVal, ",")
    If lPos > 0 Then Mid$(sVal, lPos) = "."
    LocalProofGetSingle = Val(sVal)

End Function




' **********************
' *  PUBLIC SUBS/FUNCS *
' **********************
Public Sub Refresh()
    
    UserControl.Refresh

End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
    
    MsgBox " Art2GUI V.1.0 by Light Templer" + vbCrLf + vbCrLf + _
            "An extended Shape control." + vbCrLf + vbCrLf + _
            "Just right click in IDE and select" + vbCrLf + _
            "'Edit' to get the magic.", _
            vbInformation, " About this ucArt2GUI"

End Sub




' ****************
' *  PROPERTIES  *
' ****************
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
    
    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    
    Recalc_Dimensions True
    DrawDesign
    
End Property


Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
    
    ForeColor = UserControl.ForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
        
    Recalc_Dimensions True
    DrawDesign
    
End Property


Public Property Get HiliteColor() As OLE_COLOR
Attribute HiliteColor.VB_Description = "Hilite in a gradient circle"
    
    HiliteColor = Mvar.HiliteColor

End Property

Public Property Let HiliteColor(ByVal New_HiliteColor As OLE_COLOR)
    
    Mvar.HiliteColor = New_HiliteColor
    PropertyChanged "HiliteColor"
    
    Recalc_Dimensions True
    DrawDesign
    
End Property


Public Property Get Design() As enDesign
Attribute Design.VB_Description = "Art2GUI design gradient patterns"
    
    Design = Mvar.Design

End Property

Public Property Let Design(ByVal New_Design As enDesign)
    
    If New_Design < A2G_CoronaWithLineRight Or New_Design > A2G_MultiLinGradHor Then
        RaiseError 1001, "Invalid parameter (" & New_Design & ") to property 'Desgin'!"
        
        Exit Property
    End If
    Mvar.Design = New_Design
    PropertyChanged "Design"
        
    Mvar.flgEditMode = False
    Recalc_Dimensions
    DrawDesign

End Property


Public Property Get RadialGradientWidth() As Long

    RadialGradientWidth = Mvar.lRadGradWidth

End Property

Public Property Let RadialGradientWidth(ByVal New_RadialGradientWidth As Long)

    Mvar.lRadGradWidth = New_RadialGradientWidth
    PropertyChanged "RadGradWidth"
    
    Recalc_Dimensions
    DrawDesign

End Property

Public Property Get LinearGradientWidth() As Long

    LinearGradientWidth = Mvar.lLinGradWidth

End Property

Public Property Let LinearGradientWidth(ByVal New_LinearGradientWidth As Long)
    
    If New_LinearGradientWidth > -1 Then
        Mvar.lLinGradWidth = New_LinearGradientWidth
        PropertyChanged "LinGradWidth"
        
        Recalc_Dimensions
        DrawDesign
    End If

End Property


Public Property Get Angle() As Long
Attribute Angle.VB_Description = "Gradient angle"

    Angle = Mvar.lAngle

End Property

Public Property Let Angle(ByVal New_Angle As Long)
    
    If New_Angle > -1 And New_Angle < 361 Then
        Mvar.lAngle = New_Angle
        PropertyChanged "Angle"
        
        Recalc_Dimensions True
        If Mvar.Design = A2G_LinearGradient Then
            PutDotOnCircle New_Angle, 24, Mvar.arrGrabHandles(1)
        End If
        DrawDesign
    End If

End Property


Public Property Get CenterX() As Long
Attribute CenterX.VB_Description = "X position of gradient circle"

    CenterX = Mvar.Center.lX

End Property

Public Property Let CenterX(ByVal New_CenterX As Long)
    
    Mvar.Center.lX = New_CenterX
    PropertyChanged "CenterX"
    
    Recalc_Dimensions
    DrawDesign

End Property

Public Property Get CenterY() As Long
Attribute CenterY.VB_Description = "Y position of gradient circle"

    CenterY = Mvar.Center.lY

End Property

Public Property Let CenterY(ByVal New_CenterY As Long)
    
    Mvar.Center.lY = New_CenterY
    PropertyChanged "CenterY"
    
    Recalc_Dimensions
    DrawDesign

End Property


Public Property Get GradientAreas() As Long
Attribute GradientAreas.VB_Description = "How many gradient steps"

    GradientAreas = Mvar.lGradientAreas

End Property

Public Property Let GradientAreas(ByVal New_GradientAreas As Long)
    
    Dim i As Long
    
    With Mvar
        If New_GradientAreas < 1 Or _
                New_GradientAreas > MAX_GRAD_AREAS Or _
                New_GradientAreas = .lGradientAreas Then
                
            Exit Property
        End If
        
        ReDim Preserve .arrGradAreas(1 To New_GradientAreas)
        If New_GradientAreas > .lGradientAreas Then
            ' We need presets
            For i = .lGradientAreas + 1 To New_GradientAreas
                .arrGradAreas(i).oColor = vbWhite
                .arrGradAreas(i).sngGradGamma = 1
            Next i
            ' Recalc positions
            For i = 1 To New_GradientAreas
                .arrGradAreas(i).sngPosition = (100 / New_GradientAreas) / 100 * i
            Next i
        Else
            ' Set last area's position to the right edge of control
            .arrGradAreas(New_GradientAreas).sngPosition = 1
        End If
    
        .lGradientAreas = New_GradientAreas
    End With
    PropertyChanged "GradientAreas"
    
    Recalc_Dimensions
    DrawDesign

End Property


Public Property Get GradAreaGamma(lIndex As Long) As Single
Attribute GradAreaGamma.VB_Description = "Velocity of a gradient in a multi color gradient"

    If lIndex > 0 And lIndex <= Mvar.lGradientAreas Then
        GradAreaGamma = Mvar.arrGradAreas(lIndex).sngGradGamma
    Else
        Err.Raise 5
    End If

End Property

Public Property Let GradAreaGamma(lIndex As Long, ByVal New_Gamma As Single)
    
    If lIndex > 0 And lIndex <= Mvar.lGradientAreas Then
        If New_Gamma >= 0.01 Then
            Mvar.arrGradAreas(lIndex).sngGradGamma = New_Gamma
        Else
            Err.Raise 6
        End If
    Else
        Err.Raise 5
    End If
    PropertyChanged "GradArea-GradGamma" & lIndex

    Recalc_Dimensions True
    DrawDesign

End Property


Public Property Get GradAreaColor(lIndex As Long) As OLE_COLOR
Attribute GradAreaColor.VB_Description = "Indexed color value in a multi color gradient"

    If lIndex > 0 And lIndex <= Mvar.lGradientAreas Then
        GradAreaColor = Mvar.arrGradAreas(lIndex).oColor
    Else
        Err.Raise 5
    End If

End Property

Public Property Let GradAreaColor(lIndex As Long, ByVal New_Color As OLE_COLOR)
    
    If lIndex > 0 And lIndex <= Mvar.lGradientAreas Then
        Mvar.arrGradAreas(lIndex).oColor = New_Color
    Else
        Err.Raise 5
    End If
    PropertyChanged "GradArea-Color" & lIndex

    Recalc_Dimensions True
    DrawDesign

End Property


Public Property Get GradAreaPosition(lIndex As Long) As Single
Attribute GradAreaPosition.VB_Description = "Position of a gradient in a multi color gradient"

    If lIndex > 0 And lIndex <= Mvar.lGradientAreas Then
        GradAreaPosition = Mvar.arrGradAreas(lIndex).sngPosition
    Else
        Err.Raise 5
    End If

End Property

Public Property Let GradAreaPosition(lIndex As Long, ByVal New_Position As Single)
    
    If lIndex > 0 And lIndex <= Mvar.lGradientAreas Then
        If New_Position >= 0 And New_Position <= 1 Then
            Mvar.arrGradAreas(lIndex).sngPosition = New_Position
        Else
            Err.Raise 5
        End If
    Else
        Err.Raise 5
    End If
    PropertyChanged "GradArea-Position" & lIndex

    Recalc_Dimensions True
    DrawDesign

End Property


Public Property Get Shape() As enShape
Attribute Shape.VB_Description = "Shape of control"
    
    Shape = Mvar.Shape

End Property

Public Property Let Shape(ByVal New_Shape As enShape)
    
    If New_Shape < A2G_Rectangle Or New_Shape > A2G_Diamond Then
        RaiseError 1001, "Invalid parameter (" & New_Shape & ") for property 'Shape'!"
        
        Exit Property
    End If
    If Mvar.Shape = New_Shape Then
    
        Exit Property
    End If
    
    Mvar.Shape = New_Shape
    PropertyChanged "Shape"
    
     Select Case Mvar.Shape
    
        Case A2G_Diamond, A2G_Square, A2G_Circle
                UserControl.Height = UserControl.Width
                
    End Select
    
    Mvar.flgRefreshRegion = True
    DrawDesign

End Property


Public Property Get ShapeLT() As enShapeLeftTop
Attribute ShapeLT.VB_Description = "Left / Top part of some shapes"
    
    ShapeLT = Mvar.ShapeLT

End Property

Public Property Let ShapeLT(ByVal New_ShapeLT As enShapeLeftTop)
    
    If New_ShapeLT < A2G_Shape_LT_LBTR Or New_ShapeLT > A2G_Shape_LT_LTBR Then
        RaiseError 1001, "Invalid parameter (" & New_ShapeLT & ") for property 'ShapeLT'!"
        
        Exit Property
    End If
    If Mvar.ShapeLT = New_ShapeLT Then
    
        Exit Property
    End If
    
    Mvar.ShapeLT = New_ShapeLT
    PropertyChanged "ShapeLT"
    
    Mvar.flgRefreshRegion = True
    DrawDesign

End Property


Public Property Get ShapeRB() As enShapeRightBottom
Attribute ShapeRB.VB_Description = "Right / bottom part of some shapes"
    
    ShapeRB = Mvar.ShapeRB

End Property

Public Property Let ShapeRB(ByVal New_ShapeRB As enShapeRightBottom)
    
    If New_ShapeRB < A2G_Shape_BR_LTBBR Or New_ShapeRB > A2G_Shape_BR_LBTR Then
        RaiseError 1001, "Invalid parameter (" & New_ShapeRB & ") for property 'ShapeRB'!"
        
        Exit Property
    End If
    If Mvar.ShapeRB = New_ShapeRB Then
    
        Exit Property
    End If
    
    Mvar.ShapeRB = New_ShapeRB
    PropertyChanged "ShapeRB"
    
    Mvar.flgRefreshRegion = True
    DrawDesign

End Property


Public Property Get ArcSize() As Long
Attribute ArcSize.VB_Description = "Size of a gradient arc"

    ArcSize = Mvar.lArcSize

End Property

Public Property Let ArcSize(ByVal New_ArcSize As Long)
    
    If New_ArcSize > 0 And _
            New_ArcSize < UserControl.ScaleWidth And _
            New_ArcSize < UserControl.ScaleHeight Then
        
        Mvar.lArcSize = New_ArcSize
        PropertyChanged "ArcSize"
        
        Mvar.flgRefreshRegion = True
        DrawDesign
    End If

End Property



' OOPS!  Longer than expected by me when start coding ... ;-)

' #*#

