VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Shaped Borders"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkUseFlag 
      Caption         =   "India Flag"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   22
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox chkUseFlag 
      Caption         =   "Aussy Flag"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   21
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox chkUseFlag 
      Caption         =   "US Flag"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox cboRotation 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   2040
      List            =   "frmMain.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2910
      Width           =   810
   End
   Begin VB.TextBox txtSize 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2070
      MaxLength       =   3
      TabIndex        =   16
      Text            =   "5"
      Top             =   2355
      Width           =   690
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   885
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.OptionButton optSwatch 
      Height          =   240
      Index           =   6
      Left            =   330
      TabIndex        =   7
      Top             =   450
      Width           =   255
   End
   Begin VB.OptionButton optSwatch 
      Height          =   240
      Index           =   5
      Left            =   330
      TabIndex        =   6
      Top             =   1935
      Width           =   255
   End
   Begin VB.OptionButton optSwatch 
      Height          =   240
      Index           =   4
      Left            =   330
      TabIndex        =   5
      Top             =   1710
      Width           =   255
   End
   Begin VB.OptionButton optSwatch 
      Height          =   240
      Index           =   3
      Left            =   330
      TabIndex        =   4
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optSwatch 
      Height          =   240
      Index           =   2
      Left            =   330
      TabIndex        =   3
      Top             =   1215
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton optSwatch 
      Height          =   240
      Index           =   1
      Left            =   330
      TabIndex        =   2
      Top             =   975
      Width           =   255
   End
   Begin VB.OptionButton optSwatch 
      Height          =   240
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Results"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   2
      Left            =   0
      Picture         =   "frmMain.frx":0044
      Top             =   3360
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Image Image1 
      Height          =   2400
      Index           =   1
      Left            =   -120
      Picture         =   "frmMain.frx":19586
      Top             =   3360
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Image Image1 
      Height          =   2535
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":357C8
      Top             =   3360
      Visible         =   0   'False
      Width           =   5940
   End
   Begin VB.Label Label1 
      Caption         =   "(counter clockwise)"
      Height          =   240
      Index           =   6
      Left            =   495
      TabIndex        =   19
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Border Direction in Degrees ..."
      Height          =   240
      Index           =   5
      Left            =   360
      TabIndex        =   17
      Top             =   2700
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Border Size: (1 to 255)"
      Height          =   240
      Index           =   4
      Left            =   360
      TabIndex        =   15
      Top             =   2385
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Click Colors to Change..."
      Height          =   405
      Index           =   3
      Left            =   1470
      TabIndex        =   14
      Top             =   120
      Width           =   1050
   End
   Begin VB.Label LabelColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   1470
      TabIndex        =   13
      Top             =   1785
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Outline Color"
      Height          =   240
      Index           =   2
      Left            =   1440
      TabIndex        =   12
      Top             =   1470
      Width           =   1440
   End
   Begin VB.Label LabelColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1515
      TabIndex        =   11
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Shape Fill Color"
      Height          =   240
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   795
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "3D Fill Color"
      Height          =   240
      Index           =   0
      Left            =   315
      TabIndex        =   9
      Top             =   105
      Width           =   1050
   End
   Begin VB.Label LabelColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   435
      Width           =   270
   End
   Begin VB.Image imgSwatch 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1470
      Left            =   600
      Picture         =   "frmMain.frx":6684E
      Top             =   690
      Width           =   270
   End
   Begin VB.Image imgSample 
      Height          =   2370
      Index           =   0
      Left            =   -300
      Picture         =   "frmMain.frx":67A90
      Top             =   60
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.Image imgSample 
      Height          =   2055
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":83C22
      Top             =   120
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Image imgSample 
      Height          =   2115
      Index           =   2
      Left            =   0
      Picture         =   "frmMain.frx":93D44
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Another neat trick with regions.
' Was thinking of doing a jigsaw puzzle fun project & wanted a way to
' show puzzle pieces in 3D while providing user-defined style options.
' This is basically the result of that task.

' If you wish to use this in your routines, you'll need the
' DoSample routine & the DoThreeDedge routines.
' If you want a fast "region from bitmap" creator, use the one provided
' in Module1 or use your own as needed.

' The DoSample() routine would be modified for your specific target.

' Those APIs that apply to these routines are identified below....

' APIs used in DoThreeDedge()...
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32.dll" (ByRef lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long

' APIs used in DoSample()...
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' APIs used in CreateSampleBrush()...
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private smpCountry As Integer

Private Sub DoSample()
' This is designed for a sample project only. I would imagine should you want to
' mimic this routine for your own project, you would pass it a window handle
' and pretty much remove any lines pertaining to Command1 below. Tweak as needed.
' Command1.Enabled is used as a simple flag to indicate that we already
' displayed our sample window; that's all.

Dim newWidth As Long, newHeight As Long
' window regions used to format/paint our target window
Dim winRgn As Long  ' combined mainRgn and extRgn
Dim extRgn As Long ' clipping region used for painting; this is the 3D edge
Dim testRgn As Long ' base shaped region from sample images

Dim outlineBrush As Long, fillBrush As Long

' ensure a valid 3D border size is passed
If Val(txtSize) < 1 Then txtSize = "1"
If Val(txtSize) > 255 Then txtSize = 255
txtSize = Int(Val(txtSize.Text))

' create the base shaped region
testRgn = CreateShapedRegion2(imgSample(smpCountry).Picture.Handle)

Load frmTest    'load if test form needed

' create the new window region & return the winRgn & extRgn pointers
winRgn = DoThreeDedge(testRgn, extRgn, Val(txtSize), Val(cboRotation.Text), newWidth, newHeight)
DeleteObject testRgn    ' the original shaped region is no longer needed

' resize the new window to account for the 3D edge
frmTest.Move frmTest.Left, frmTest.Top, newWidth * Screen.TwipsPerPixelX, newHeight * Screen.TwipsPerPixelY

' blank out so ready for painting
If Not Command1.Enabled Then frmTest.Cls

' reset the back color as needed
frmTest.BackColor = LabelColor(1).BackColor
' select the 3D edge region as the clipping region so we can fill it
SelectClipRgn frmTest.hdc, extRgn

If optSwatch(6).Value = True Then
    ' solid 3D fill color
    fillBrush = CreateSolidBrush(LabelColor(0).BackColor)
Else
    ' bitmap 3D fill style
    fillBrush = CreateSampleBrush(Val(optSwatch(0).Tag))
End If
' fill the 3D edge portion
FillRgn frmTest.hdc, extRgn, fillBrush

' remove the clipping region so we can draw our borders
SelectClipRgn frmTest.hdc, ByVal 0&

' create brush & draw border around the entire region
outlineBrush = CreateSolidBrush(LabelColor(2).BackColor)
FrameRgn frmTest.hdc, winRgn, outlineBrush, 1, 1

' now we'll remove the 3D edge portion & draw another border (inner border)
CombineRgn extRgn, winRgn, extRgn, 4


If Len(chkUseFlag(0).Tag) Then
    ' this is an example of how you could blt an image over the main shape
    ' In the example below I use Render, but BitBlt/StretchBlt could just
    ' as easily be used. Notice how we set & them remove the clipping region
    Dim bltRect As RECT
    ' use this as the clipping region
    SelectClipRgn frmTest.hdc, extRgn
    ' get the left/top position of the main shape
    GetRgnBox extRgn, bltRect
    ' paint the sample fill
    With Image1(Val(chkUseFlag(0).Tag)).Picture
        Select Case Val(chkUseFlag(0).Tag)
        Case 0 ' us flag
        .Render frmTest.hdc + 0, bltRect.Left - 55, bltRect.Top - 8, _
            ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), _
            0, .Height, .Width, -.Height, ByVal 0&
        Case 1 ' austraila flag
        .Render frmTest.hdc + 0, bltRect.Left - 15, bltRect.Top - 0, _
            ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), _
            0, .Height, .Width, -.Height, ByVal 0&
        Case 2 ' india flag
        .Render frmTest.hdc + 0, bltRect.Left - 70, bltRect.Top - 0, _
            ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), _
            0, .Height, .Width, -.Height, ByVal 0&
        End Select
    End With
    ' remove the clipping region so we can draw other stuff if desired
    SelectClipRgn frmTest.hdc, ByVal 0&
End If

' here we will frame the inner border
FrameRgn frmTest.hdc, extRgn, outlineBrush, 1, 1

' clean up.  The only region not deleted is the winRgn 'cause we will assign
' it to our test window via SetWindowRgn
DeleteObject extRgn
DeleteObject outlineBrush
DeleteObject fillBrush

' update the changes as needed
SetWindowRgn frmTest.hwnd, winRgn, True
If Command1.Enabled Then
    frmTest.Show , Me
Else
    frmTest.Refresh
End If

End Sub

Private Function DoThreeDedge(ByVal inRegion As Long, ThreeDRgn As Long, _
        ByVal ThreeDsize As Byte, ByVal ThreeDangle As Integer, _
        RegionCX As Long, RegionCY As Long) As Long

' Function returns 2 regions, one as the return value and another in the
' ThreeDRgn parameter.  Each region must be destroyed by user unless assigned
' to a window region via SetWindowRgn API.

' Parameters.
' inRegion : this is the target shaped region used to create the new 3D region
' ThreeDRgn : this is the 3D portion of the new overall region
' ThreeDsize : this is the 3D size in bytes (1-255).
'       Zero will return a copy of the passed inRegion & ThreeDRgn will be zero
' ThreeDangle : A valid positive angle from 0 to 360 in 45 degree increments.
'   Any non-valid angle will be handled as 315 degrees
' RegionCx is the new calculated region width & can be used for sizing a window
' RegionCy is the new calculated region height & can be used for sizing a window

' regions to be created & modified
Dim pRgn As Long    ' this will be the return value
Dim cRgn As Long    ' temp region used for creating the 3D effect

' rectangles used to align the new regions with the old
Dim pRect As RECT
Dim wRect As RECT

' used for offsets when aligning regions
Dim cxAdjust As Long, cyAdjust As Long
Dim I As Long   ' loop variable
' direction of the 3D effect
Dim Xdir As Long, Ydir As Long

' Create temporary regions...

' make a copy of the passed region
pRgn = CreateRectRgnIndirect(pRect)
CombineRgn pRgn, inRegion, pRgn, 5
If ThreeDsize = 0 Then Exit Function

' make another copy to be ultimately used for the clipping region
ThreeDRgn = CreateRectRgnIndirect(pRect)

' make a 3rd copy to be used for creating the 3D effect
cRgn = CreateRectRgnIndirect(pRect)
CombineRgn cRgn, inRegion, cRgn, 5

' set the direction of X,Y depending on the passed rotation
Select Case ThreeDangle
Case 0, 360
    Xdir = 1: Ydir = 0
Case 45
    Xdir = 1: Ydir = -1
Case 90
    Xdir = 0: Ydir = -1
Case 135
    Xdir = -1: Ydir = -1
Case 180
    Xdir = -1: Ydir = 0
Case 225
    Xdir = -1: Ydir = 1
Case 270
    Xdir = 0: Ydir = 1
Case Else
    Xdir = 1: Ydir = 1
End Select
    
' we create the 3D effect by simply sliding the original region
' a pixel at a time in the appropriate direction enough times to
' satisify the size requested. Simple, huh?
For I = 1 To ThreeDsize
    OffsetRgn cRgn, Xdir, Ydir                  ' slide region
    CombineRgn pRgn, pRgn, cRgn, 2              ' add to overall region
Next
DeleteObject cRgn           ' no longer needed, delete now

' get the X,Y position of the combined region
GetRgnBox pRgn, pRect
' get the original position of the passed region
GetRgnBox inRegion, wRect
    
' create the actual 3D region
CombineRgn ThreeDRgn, pRgn, inRegion, 4
    
' when the 3D border is above or left to the original region, we need to
' do some offsetting to align everything up

' set the offsets. Not all passed regions will have a 0,0 top/left corner
If Xdir < 0 Then cxAdjust = pRect.Left Else cxAdjust = wRect.Left
If Ydir < 0 Then cyAdjust = pRect.Top Else cyAdjust = wRect.Top
    
' perform the offsets
OffsetRgn pRgn, -cxAdjust + wRect.Left, -cyAdjust + wRect.Top
OffsetRgn ThreeDRgn, -cxAdjust + wRect.Left, -cyAdjust + wRect.Top
    
' return the calculated region width & height
RegionCX = wRect.Right + ThreeDsize * Abs(Xdir) + Abs(wRect.Left)
RegionCY = wRect.Bottom + ThreeDsize * Abs(Ydir) + Abs(wRect.Top)
    
' return the new comined region
DoThreeDedge = pRgn

End Function

Private Function CreateSampleBrush(Style As Long) As Long
' Function returns a handle to a bitmap brush.
' Note: Win95, I believe, will only use the first 8x8 of the bitmap

' I used a 16x80 swatch with 6 stacked sample bitmaps

Dim xOffset As Long
Dim tDC As Long, hOldBmp As Long, hNewBmp As Long

' create temp DC & bitmap.
' If used outside of this form, we would use
' the return value of GetDC(GetDesktopWindow()) vs Me.hDC below
tDC = CreateCompatibleDC(Me.hdc)
hNewBmp = CreateCompatibleBitmap(Me.hdc, 16, 16)

' select fresh bitmap into our temp DC & Blt over the appropriate 16x16 bits
hOldBmp = SelectObject(tDC, hNewBmp)
With imgSwatch.Picture
    .Render tDC + 0, 0&, 0&, 16&, 16&, 0, _
        .Height - (.Width * Style), .Width, -.Width, ByVal 0&
End With

' remove the bitmap & replace original, then delete the DC
SelectObject tDC, hOldBmp
DeleteDC tDC

' create the brush & then delete the temp bitmap
CreateSampleBrush = CreatePatternBrush(hNewBmp)
DeleteObject hNewBmp

End Function

Private Sub cboRotation_Click()
' option to change degrees of 3D border
If Not Command1.Enabled Then DoSample
End Sub

Private Sub chkUseFlag_Click(Index As Integer)
If chkUseFlag(Index) = 0 Then
    If chkUseFlag(0).Value + chkUseFlag(1).Value + chkUseFlag(2).Value > 0 Then Exit Sub
    chkUseFlag(0).Tag = ""
Else
    Dim I As Integer
    For I = 0 To chkUseFlag.UBound
        If I <> Index Then chkUseFlag(I) = 0
    Next
    chkUseFlag(0).Tag = Index
    smpCountry = Index
End If
If Not Command1.Enabled Then DoSample
End Sub

Private Sub Command1_Click()
' show the example, then provide message box
' The Enabled status of this button will be checked to see if we
' have already displayed a sample form
DoSample
MsgBox "Simply change settings on this form & the sample will update automatically." & _
    vbNewLine & vbNewLine & "You can drag the sample & closing the main window will close the sample window too.", vbInformation + vbOKOnly, "FYI"
Command1.Enabled = False
End Sub

Private Sub Form_Load()
' initially set a 3D rotation value
cboRotation.ListIndex = 5
optSwatch(0) = True
Me.Move (Screen.Width - Width) \ 2 - Width, (Screen.Height - Height) \ 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
' remove the test form if needed
Unload frmTest
End Sub

Private Sub imgSwatch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' select the appropriate option button if user clicks on the swatch
If Button = vbLeftButton Then
    On Error Resume Next
    Dim swID As Integer
    swID = Y \ Screen.TwipsPerPixelY \ 16
    If optSwatch(swID).Value = False Then optSwatch(swID).Value = True
End If
End Sub

Private Sub LabelColor_Click(Index As Integer)
' get custom solid color & update the test form if needed
GetColor Index
If Index = 0 And optSwatch(6).Value = False Then
    optSwatch(6).Value = True
Else
    If Not Command1.Enabled Then DoSample
End If
End Sub

Private Sub optSwatch_Click(Index As Integer)
' update test form as user chooses different fill styles
If optSwatch(Index).Value = True Then
    optSwatch(0).Tag = Index
    If Not Command1.Enabled Then DoSample
End If
End Sub

Private Sub optSwatch_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' if solid color fill option selected & user clicks on the
' option button again, display the color dialog
If Index = 6 Then
    If optSwatch(Index).Value = True Then
        Call LabelColor_Click(0)
    End If
End If
End Sub

Private Sub GetColor(lblIndex As Integer)
' called when user needs to select color from color dialog
With dlgCommon
    .Flags = cdlCCRGBInit
    .Color = LabelColor(lblIndex).BackColor
End With
On Error Resume Next
dlgCommon.ShowColor
If Err.Number = 0 Then LabelColor(lblIndex).BackColor = dlgCommon.Color
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
' when user hits enter in text box, update sample form
If KeyAscii = vbKeyReturn Then
    If Not Command1.Enabled Then DoSample
    KeyAscii = 0
End If
End Sub
