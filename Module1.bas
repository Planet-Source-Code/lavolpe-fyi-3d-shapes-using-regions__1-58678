Attribute VB_Name = "Module1"
'Attribute VB_Name = "modRegionShape2"
Option Explicit

' BRIEF HISTORY
' \/\/\/\/\/\/\
' *** VERSION 1 - used GetPixel
' *** VERSION 2 - use of DIBs & SafeArrays vs GetPixel (Jun 04)
' UPDATED to include using DIB sections. The increase in speed is
' truly amazing over my earlier version.....
' Over a 300% increase in speed noted on smallish bitmaps (96x96) &
' Over a 425% increase noted on mid-sized bitmaps (265x265)

' This approach should be somewhat faster than typical approaches that
' create a region by using CreateRectRgn, CombineRgn & DeleteObject.
' That is 'cause this approach does not use any of those functions to
' create regions. Like typical approaches a rectangle format is used to
' "add to a region", however, this approach directly creates the region header
' & region structure and passes that to a single API to finish the job of
' creating the end result. For those that play around with such things,
' I think you will recognize the difference.

' EDITED: The ExtCreateRegion seems to have an undocumented restriction: it
' won't create regions comprising of more than 4K rectangles (win98). So to
' get around this for very complex bitmaps, I rewrote the function to create
' regions of 2K rectangles at a time if needed. This is still extremely
' fast. I compared the window shaping code from vbAccelerator with the
' SandStone.bmp (15,000 rects) in Windows directory. vbAccelerator's code
' averaged 4,900 ms. My routines averaged 77 ms & that's not a typo!

' EDITED: I was allowing default error trapping on UBound() to resize the
' rectangle array: when trying to update array element beyond UBound, error
' would occur & be redirected to resize the array. However, thanx to
' Robert Rayment, if the UBound checks are disabled in compile optimizations,
' then we get a crash. Therefore, checks made appropriately & a tiny loss of
' speed is the trade-off for safety.

' *** VERSION 3 - Use of GetDIBits vs SafeArray pointers
' The function CreateShapedRegion does not use all that fancy
' SafeArray stuff used in previous version.
' It was tweaked to use the bytes returned by GetDIBits API
' No noticeable loss of speed resulted

' *** VERSION 3.1 - Anti Regions & speed modifications (10 Jan 05)
' Function has an optional parameter to return the anti-Region.
' That is the region of only transparent pixels. This could be used, for
' example, with APIs like FillRgn to replace the "transparent" color with
' another color.

' Speed modifications...before tweaks & using the WinNT.bmp with white
'   transparent color (lots of white!), routines averaged 31ms to
'   process the bitmap. After tweaks, routines averaged 16ms. Not bad ;)
' Thanx on this suggestion goes to Carles P.V. who gave me hints towards
' this direction 6 months ago. Not knowing that much about DIBs then,
' I didn't see the light until recently.

' *** VERSION 3.2 - Rotation of Regions & Images (13 Jan 05)
' This latest update adds capability of rotating a region & therefore
' the image too at 90 degree intervals and mirroring vertically/
' horizonitally. The function has an option to return the
' rotated image along with the region: see RotateImageRegion routine
'   To rotate an existing region not dependent
'   upon an image, see RotateSimpleRegion

' *** VERSION 3.3 - Another Rotation Option & Stretching Option
' RotateSimpleRegion will rotate a region without need to pass an image
' StretchRegion stretches a region much like StretchBlt does to bitmaps
' ***********************************************************************

' GDI32 APIs
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByRef lpRgnData As Any) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef Bits As Any, ByRef BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
' Kernel32 APIs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
' User32 APIs
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' olepro32 APIs
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, _
    ipic As IPicture) As Long


' TYPEs used

' used to convert icons/bitmaps to stdPicture objects
Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type
' this is made public due to the ExtractRegionRects function since it
' returns an array of RECTs. That function can easily be tweaked to
' return a byte array which would not require this UDT to be public.
' The tweaking code is provided in that function but commented out
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' following UDT used for stretching a region
' See StretchRegion function
Private Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 3) As Byte ' used vs RGBQUAD structure
End Type
Public Enum dibRotationEnum
    Rotate90 = 0
    Rotate180 = 1
    Rotate270 = 2
    MirrorHorizontal = 3
    MirrorVertical = 4
End Enum
    
' Constants used
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const RGN_OR As Long = 2

' for testing only
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Function CreateShapedRegion2(ByVal hBitmap As Long, _
        Optional destinationDC As Long, _
        Optional ByVal transColor As Long = -1, _
        Optional returnAntiRegion As Boolean) As Long

'*******************************************************
' FUNCTION RETURNS A HANDLE TO A REGION IF SUCCESSFUL.
' If unsuccessful, function retuns zero.
'*******************************************************
' PARAMETERS
'=============
' hBitmap : handle to a bitmap to be used to create the region
' destinationDC : used by GetDibits API. If not supplied then desktop DC used
' transColor : the transparent color
' returnAntiRegion : If False (default) then the region excluding transparent
'       pixels will be returned.  If True, then the region including only
'       transparent pixels will be returned


' test for required variable first
If hBitmap = 0 Then Exit Function

Dim aStart As Long, aStop As Long ' testing purposes
aStart = GetTickCount()            ' testing purposes

' now ensure hBitmap handle passed is a usable bitmap
Dim bmpInfo As BITMAPINFO
If GetGDIObject(hBitmap, Len(bmpInfo), bmpInfo) = 0 Then Exit Function

' declare bunch of variables...
Dim srcDC As Long   ' DC to use for GetDibits
Dim rgnRects() As RECT ' array of rectangles comprising region
Dim rectCount As Long ' number of rectangles & used to increment above array
Dim rStart As Long ' pixel that begins a new regional rectangle

Dim X As Long, Y As Long ' loop counters
Dim lScanLines As Long ' used to size the DIB bit array

Dim bDib() As Byte  ' the DIB bit array
Dim bBGR(0 To 3) As Byte ' used to copy long to bytes
Dim tgtColor As Long ' a DIB pixel color
Dim rtnRegion As Long ' region handle returned by this function

On Error GoTo CleanUp
  
' use passed DC if supplied, otherwise use desktop DC
If destinationDC = 0 Then
    srcDC = GetDC(GetDesktopWindow())
Else
    srcDC = destinationDC
End If
    
' Scans must align on dword boundaries:
lScanLines = (bmpInfo.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
ReDim bDib(0 To lScanLines - 1, 0 To bmpInfo.bmiHeader.biHeight - 1)

' build a DIB header
' DIBs are bottom to top, so by using negative Height
' we will load it top to bottom
With bmpInfo.bmiHeader
   .biSize = Len(bmpInfo.bmiHeader)
   .biBitCount = 24
   .biHeight = -.biHeight
   .biPlanes = 1
   .biCompression = BI_RGB
   .biSizeImage = lScanLines * .biHeight
End With

' get the image into DIB bits,
' note that biHeight above was changed to negative so we reverse it form here on
Call GetDIBits(srcDC, hBitmap, 0, -bmpInfo.bmiHeader.biHeight, bDib(0, 0), bmpInfo, 0)
    
' now get the transparent color if needed
If transColor < 0 Then
    ' when negative value passed, use top left corner pixel color
    CopyMemory bBGR(0), bDib(0, 0), &H3
Else
    ' 24bit DIBs are stored as BGR vs RGB
    ' convert it now for one color vs converting each BGR pixel to RGB
    bBGR(2) = (transColor And &HFF&)
    bBGR(1) = (transColor And &HFF00&) \ &H100&
    bBGR(0) = (transColor And &HFF0000) \ &H10000
End If
' copy bytes to long
CopyMemory transColor, bBGR(0), &H4
    
With bmpInfo.bmiHeader
 
     ' start with an arbritray number of rectangles
    ReDim rgnRects(0 To .biWidth * 3)
    ' reset flag
    rStart = -1
    
    ' begin pixel by pixel comparisons
    For Y = 0 To Abs(.biHeight) - 1
        For X = 0 To .biWidth - 1
            ' my hack continued: we already saved a long as BGR, now
            ' get the current DIB pixel into a long (BGR also) & compare
            CopyMemory tgtColor, bDib(X * 3, Y), &H3
            
            ' test to see if next pixel is a target color
            If transColor = tgtColor Xor returnAntiRegion Then
                
                If rStart > -1 Then ' we're currently tracking a rectangle,
                                    ' so let's close it
                    ' see if array needs to be resized
                   If rectCount + 1 = UBound(rgnRects) Then _
                       ReDim Preserve rgnRects(0 To UBound(rgnRects) + .biWidth * 3)
                    
                    ' add the rectangle to our array
                    SetRect rgnRects(rectCount + 2), rStart, Y, X, Y + 1
                    rStart = -1 ' reset flag
                    rectCount = rectCount + 1     ' keep track of nr in use
                End If
            
            Else
                ' not a target color
                If rStart = -1 Then rStart = X ' set start point
            
            End If
        Next X
        If rStart > -1 Then
            ' got to end of bitmap without hitting another transparent pixel
            ' but we're tracking so we'll close rectangle now
           
                ' see if array needs to be resized
           If rectCount + 1 = UBound(rgnRects) Then _
               ReDim Preserve rgnRects(0 To UBound(rgnRects) + .biWidth * 3)
                ' add the rectangle to our array
            SetRect rgnRects(rectCount + 2), rStart, Y, X, Y + 1
            rStart = -1 ' reset flag
            rectCount = rectCount + 1     ' keep track of nr in use
        End If
    Next Y
End With
Erase bDib
        
On Error Resume Next
' check for failure & engage backup plan if needed
If rectCount Then
    ' there were rectangles identified, try to create the region
    rtnRegion = CreatePartialRegion(rgnRects(), 2, rectCount + 1, 0, bmpInfo.bmiHeader.biWidth)
    
    ' ok, now to test whether or not we are good to go...
    ' if less than 2000 rectangles, API should have worked & if it didn't
    ' it wasn't due O/S restrictions -- failure
    
    If rtnRegion = 0 And rectCount > 2000 Then
        rtnRegion = CreateWin98Region(rgnRects, rectCount + 1, 0, bmpInfo.bmiHeader.biWidth)
    End If

End If

CleanUp:

If destinationDC <> srcDC Then ReleaseDC GetDesktopWindow(), srcDC
Erase rgnRects()

If Err Then
    If rtnRegion Then DeleteObject rtnRegion
    Err.Clear
    MsgBox "Shaped Region failed. Windows could not create the region."
Else
    CreateShapedRegion2 = rtnRegion
    aStop = GetTickCount()          ' testing purposes
'    MsgBox aStop - aStart & " ms"  ' unRem to show message box
End If


End Function

Private Function CreatePartialRegion(rgnRects() As RECT, lIndex As Long, uIndex As Long, leftOffset As Long, Cx As Long) As Long
' Called when large region fails (can be the case with Win98) and also called
' when rotation a region 90 or 270 degrees (see RotateSimpleRegion)

On Error Resume Next
' Note: Ideally contiguous rectangles of equal height & width should be combined
' into one larger rectangle. However, thru trial & error I found that Windows
' does this for us and taking the extra time to do it ourselves
' is to cumbersome & slows down the results.

' the first 32 bytes of a region is the header describing the region.
' Well 32 bytes equates to 2 rectangles (16 bytes each), so I'll
' cheat a little & use rectangles to store the header
With rgnRects(lIndex - 2) ' bytes 0-15
    .Left = 32                      ' length of region header in bytes
    .Top = 1                        ' required cannot be anything else
    .Right = uIndex - lIndex + 1    ' number of rectangles for the region
    .Bottom = .Right * 16&          ' byte size used by the rectangles; can be zero
End With
With rgnRects(lIndex - 1) ' bytes 16-31 bounding rectangle identification
    .Left = leftOffset                  ' left
    .Top = rgnRects(lIndex).Top         ' top
    .Right = leftOffset + Cx            ' right
    .Bottom = rgnRects(uIndex).Bottom   ' bottom
End With
' call function to create region from our byte (RECT) array
CreatePartialRegion = ExtCreateRegion(ByVal 0&, (rgnRects(lIndex - 2).Right + 2) * 16, rgnRects(lIndex - 2))
If Err Then Err.Clear
End Function

Private Function CreateWin98Region(rgnRects() As RECT, rectCount As Long, leftOffset As Long, Cx As Long) As Long
' Pulled out of main routine 'cause now two routines use the same logic
' and we will simply share this part of the code

' Win98 has problems with regional rectangles over 4000
' So, we'll try again in case this is the prob with other systems too.
' We'll step it at 2000 at a time which is stil very quick

Dim X As Long, Y As Long ' loop counters
Dim win98Rgn As Long     ' partial region
Dim rtnRegion As Long    ' combined region & return value of this function

' we start with 2 'cause first 2 RECTs is the header
For X = 2 To rectCount Step 2000

    If X + 2000 > rectCount Then
        Y = rectCount
    Else
        Y = X + 2000
    End If
    
    ' attempt to create partial region
    win98Rgn = CreatePartialRegion(rgnRects(), X, Y, leftOffset, Cx)
    
    If win98Rgn = 0 Then    ' failure
        ' cleaup combined region if needed
        If rtnRegion Then DeleteObject rtnRegion
        Exit For ' abort
    Else
        If rtnRegion Then ' already started
            ' use combineRgn, but only every 2000th time
            CombineRgn rtnRegion, rtnRegion, win98Rgn, RGN_OR
            DeleteObject win98Rgn
        Else    ' first time thru
            rtnRegion = win98Rgn
        End If
    End If
Next
' done; return result
CreateWin98Region = rtnRegion
End Function

Private Function HandleToPicture(ByVal hHandle As Long, isBitmap As Boolean) As Picture
' Convert an icon/bitmap handle to a Picture object

On Error GoTo ExitRoutine

    Dim pic As PICTDESC
    Dim guid(0 To 3) As Long
    
    ' initialize the PictDesc structure
    pic.cbSize = Len(pic)
    If isBitmap Then pic.pictType = vbPicTypeBitmap Else pic.pictType = vbPicTypeIcon
    pic.hIcon = hHandle
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect pic, guid(0), True, HandleToPicture

ExitRoutine:
End Function

Public Function RotateImageRegion(ByVal hBitmap As Long, _
        Optional ByVal destinationDC As Long, _
        Optional ByVal transColor As Long = -1, _
        Optional ByVal returnAntiRegion As Boolean, _
        Optional ByVal Rotation As dibRotationEnum = MirrorHorizontal, _
        Optional returnImage As StdPicture) As Long
        
' Lead Note: By removing the last IF:THEN:ELSE statements at the
' very bottom, this routine can easily be used to simply rotate images.

'************************************************************
' FUNCTION RETURNS A HANDLE TO A ROTATED REGION IF SUCCESSFUL.
' IT WILL ALSO RETURN THE ROTATED hBitmap AS A NEW STDPICTURE.
' If unsuccessful, function retuns zero and returns no image.
'************************************************************
' PARAMETERS
'=============
' hBitmap : handle to a bitmap to be used to create the region
' destinationDC : used by GetDibits API. If not supplied then desktop DC used
' transColor : the transparent color
' returnAntiRegion : If False (default) then the region excluding transparent
'       pixels will be returned.  If True, then the region including only
'       transparent pixels will be returned
' Rotation : Option to rotate region or mirror region
' returnImage : If stdPicture supplied, the modified image will be returned


' test for required variable first
If hBitmap = 0 Then Exit Function

' now ensure hBitmap handle passed is a usable bitmap
Dim bmpInfo As BITMAPINFO ' used to get/set DIB sections
If GetGDIObject(hBitmap, Len(bmpInfo), bmpInfo) = 0 Then Exit Function

' declare bunch of variables...
Dim srcDC As Long   ' DC to use for Set/GetDibits
Dim X As Long, Y As Long ' loop counters
Dim lScanLines As Long ' used to size the DIB bit array
Dim bDib() As Byte  ' the DIB bit array
Dim bDibNew() As Byte ' the rotated DIB bit array
Dim hDib As Long ' handle to modified image
Dim rotBmpInfo As BITMAPINFO ' used for modified image
Dim rightEdge As Long ' cached calculation
Dim bottomEdge As Long 'cached calculation

On Error GoTo ExitRoutine
' use passed DC if supplied, otherwise use desktop DC
If destinationDC = 0 Then
    srcDC = GetDC(GetDesktopWindow())
Else
    srcDC = destinationDC
End If
    
' Scans must align on dword boundaries:
lScanLines = (bmpInfo.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
ReDim bDib(0 To lScanLines - 1, 0 To bmpInfo.bmiHeader.biHeight - 1)

On Error GoTo CleanUp
' build a DIB header
' DIBs are bottom to top, so by using negative Height
' we will load it top to bottom
With bmpInfo.bmiHeader
   .biSize = Len(bmpInfo.bmiHeader)
   .biBitCount = 24
   .biHeight = -.biHeight
   .biPlanes = 1
   .biCompression = BI_RGB
   .biSizeImage = lScanLines * Abs(.biHeight)
End With

' get the image into DIB bits,
' note that biHeight above was changed to negative so we reverse it form here on
Call GetDIBits(srcDC, hBitmap, 0, Abs(bmpInfo.bmiHeader.biHeight), bDib(0, 0), bmpInfo, 0)
rotBmpInfo = bmpInfo
    
' each rotation has its own loop for debugging considerations
' and also speed: adding multiple IF's in loop would make it a bit slower
Select Case Rotation

Case 0, 2 ' rotate 90 or 270 respectively
    
    ' see if we need to adjust image & region to a new width & height
    If Abs(bmpInfo.bmiHeader.biHeight) <> bmpInfo.bmiHeader.biWidth Then
        ' cache the source height to be used in calculation & UDT element below
        lScanLines = (bmpInfo.bmiHeader.biHeight * -3 + 3) And &HFFFFFFFC
        With rotBmpInfo.bmiHeader
           .biSize = Len(bmpInfo.bmiHeader)
           .biHeight = -bmpInfo.bmiHeader.biWidth
           .biWidth = -bmpInfo.bmiHeader.biHeight
           .biSizeImage = lScanLines * Abs(.biHeight)
        End With
    End If
    hDib = CreateDIBSection(srcDC, rotBmpInfo, 0&, 0&, 0&, DIB_RGB_COLORS)
    
    If hDib Then
        ReDim bDibNew(0 To lScanLines - 1, 0 To Abs(rotBmpInfo.bmiHeader.biHeight) - 1)
        With rotBmpInfo.bmiHeader
            If Rotation = 0 Then    ' rotate 90 degrees
                bottomEdge = UBound(bDib, 2)
                For Y = 0 To Abs(.biHeight) - 1
                    For X = 0 To .biWidth - 1
                        ' fill destination left to right, top to bottom
                        ' using source from bottom to top, left to right
                        CopyMemory bDibNew(X * 3, Y), bDib(Y * 3, bottomEdge - X), &H3
                    Next
                Next
            Else                    ' rotate 270 degrees
                ' cache the original source width
                rightEdge = (bmpInfo.bmiHeader.biWidth - 1) * 3
                For Y = 0 To Abs(.biHeight) - 1
                    For X = 0 To .biWidth - 1
                        ' fill destination left to right, top to bottom
                        ' using source from top to bottom, right to left
                        CopyMemory bDibNew(X * 3, Y), bDib(rightEdge - Y * 3, X), &H3
                    Next
                Next
            End If
        End With
    End If
    
Case 1, 3, 4 ' rotate 180, mirror horizontally or vertically respectively
    hDib = CreateDIBSection(srcDC, rotBmpInfo, 0&, 0&, 0&, DIB_RGB_COLORS)
    If hDib Then
        ReDim bDibNew(0 To UBound(bDib, 1), 0 To UBound(bDib, 2))
        With rotBmpInfo.bmiHeader
        
            ' cache the width & height for use later
            rightEdge = (.biWidth - 1) * 3
            bottomEdge = Abs(.biHeight) - 1
            
            If Rotation = 4 Then    ' mirror vertically
                For Y = 0 To bottomEdge
                    For X = 0 To .biWidth - 1
                        ' fill destination left to right, top to bottom
                        ' using source from left to right, bottom to top
                        CopyMemory bDibNew(X * 3, Y), bDib(X * 3, bottomEdge - Y), &H3
                    Next
                Next
                
            ElseIf Rotation = 3 Then    ' mirror horizontally
                For Y = 0 To bottomEdge
                    For X = 0 To .biWidth - 1
                        ' fill destination left to right, top to bottom
                        ' using source from right to left, top to bottom
                        CopyMemory bDibNew(rightEdge - X * 3, Y), bDib(X * 3, Y), &H3
                    Next
                Next
            Else                        ' rotate 180 degrees
                For Y = 0 To bottomEdge
                    For X = 0 To .biWidth - 1
                        ' fill destination left to right, top to bottom
                        ' using source from right to left, bottom to top
                        CopyMemory bDibNew(X * 3, Y), bDib(rightEdge - X * 3, bottomEdge - Y), &H3
                    Next
                Next
            End If
        End With
    End If
    
Case Else
    ' bad parameter passed, we will use the passed image as is.
    ' Therefore hDib = 0 & next set of IFs handle that
End Select

CleanUp:
If Err Then Err.Clear
On Error Resume Next
If hDib Then
    ' rotated image bits, set the bits to the hDIB
    SetDIBits srcDC, hDib, 0, Abs(rotBmpInfo.bmiHeader.biHeight), bDibNew(0, 0), rotBmpInfo, 0&
Else
    ' error or bad calling convention, use source image & bits
    hDib = CreateDIBSection(srcDC, bmpInfo, 0&, 0&, 0&, DIB_RGB_COLORS)
    If hDib Then SetDIBits srcDC, hDib, 0, Abs(bmpInfo.bmiHeader.biHeight), bDib(0, 0), bmpInfo, 0&
End If
Erase bDibNew
If transColor < 0 Then
    ' when negative value passed, use top left corner pixel color
    ' from the source image, not the rotated image
    bmpInfo.bmiColors(0) = bDib(2, 0)
    bmpInfo.bmiColors(1) = bDib(1, 0)
    bmpInfo.bmiColors(2) = bDib(0, 0)
    bmpInfo.bmiColors(3) = 0
    CopyMemory transColor, bmpInfo.bmiColors(0), &H4
End If
Erase bDib
If srcDC <> destinationDC Then ReleaseDC GetDesktopWindow(), srcDC
If hDib Then
    ' ok we have an image to play with, let's convert it to a stdPicture
    ' and then send that picture to create the shaped region
    Set returnImage = HandleToPicture(hDib, True)
    If returnImage Is Nothing Then
        DeleteObject hDib
    Else
        If returnImage.Handle = 0 Then
            Set returnImage = Nothing
            DeleteObject hDib
        Else
            RotateImageRegion = CreateShapedRegion2(returnImage.Handle, destinationDC, transColor, returnAntiRegion)
        End If
    End If
End If
ExitRoutine:
End Function


Public Function ExtractRegionRects(hRgn As Long, Optional includeHeader As Boolean) As RECT()
' Lead Note: To have this function return a byte array vs a RECT array.
' If using UDTs in your app is not doable, change function to return Bytes
' Rem & unRem the 3 lines noted in this routine

' #1: to return a byte array replace function above with
'Public Function ExtractRegionRects(hRgn As Long, Optional includeHeader As Boolean) As Byte()


' Regions are comprised of simply rectangles.  The rectangles are always
' sorted top to bottom, left to right and no rectangle ever overlaps
' another rectangle

'************************************************************
' Function will return the rectangle structure of any region either as
' a RECT array or Byte array (depending on how you tweak this routine)

' If the function fails, a -1 ubound array is returned
'************************************************************
' PARAMETERS
'=============
' hRgn : handle to a valid, existing region


Dim vRects() As RECT
' #2: to return a byte array swap above & next DIM statements
'Dim vRects() As Byte

ReDim vRects(-1 To -1)


If hRgn = 0 Then Exit Function

Dim rgnDatOffset As Long
Dim vRgnData() As Byte  ' byte array of the entire region, including header
Dim rSize As Long       ' used for API return value
'Dim boundRect As RECT ' unRem if you want to extract region's bounding rectangle


' 1st get the buffer size needed to return rectangle info from this region
rSize = GetRegionData(hRgn, ByVal 0&, ByVal 0&)
If rSize > 0 Then   ' success
    
    ' create the buffer & call function again to fill the buffer
    ReDim vRgnData(0 To rSize - 1) As Byte
    If rSize = GetRegionData(hRgn, rSize, vRgnData(0)) Then     ' success
    
        ' Here are some tips of the structure returned
        ' Bytes 8-11 are the number of rectangles in the region
        ' Bytes 12-15 is structure size information -- not important for what we need
        ' Bytes 16-31 are the bounding rectangle's dimensions
        ' Bytes 32 to end of structure are the individual rectangle's dimensions
        ' The rectangle structure (RECT) is 16 bytes or LenB(RECT)
        
        ' Let's retrieve the number of rectangles in the structure (bytes:8-11)
        CopyMemory rSize, vRgnData(8), ByVal 4&
        
        If includeHeader Then
            rSize = rSize + 2
        Else
            rgnDatOffset = 1
        End If
        
        ' #3: to return a byte array swap following ReDim statements
            ' Resize our Rectangle/Byte array
            ReDim vRects(0 To rSize - 1)
'            ReDim vRects(0 To (UBound(vRgnData) - 31))
            
        ' we want all bytes starting with byte 32 or byte 0 depending on
        ' the value of regionShapeUse
        CopyMemory vRects(0), vRgnData(32 * rgnDatOffset), (UBound(vRgnData) - (31 * rgnDatOffset))
        
        ' Here's how we can extract the bounding rectangle.
        ' Using the API GetRgnBox will do the same
'        CopyMemory boundRect, vRgnData(16), ByVal 16&   ' (bytes:16-31)
        
    End If
    
    Erase vRgnData
    
End If

ExtractRegionRects = vRects()

End Function

Public Function RotateSimpleRegion(hSource As Long, fromWindow As Boolean, ByVal Rotation As dibRotationEnum) As Long
' this function does not require an image to calculate a rotated region,
' where the function RotateImageRegion does require an image
' Any valid region can be passed.

'************************************************************
' FUNCTION RETURNS A HANDLE TO A ROTATED REGION IF SUCCESSFUL.
' If unsuccessful, function retuns zero.
'************************************************************
' PARAMETERS
'=============
' hSource : handle to a valid window region or of a valid window
' fromWindow : if True, then hSource must be window handle,
'              otherwise hSource must be a window region
' Rotation : Option to rotate region or mirror region.
'           Invalid parameter causes passed region to be copied as is

' ensure required parameter passed
If hSource = 0 Then Exit Function

Dim rgnRects() As RECT  ' region data in RECT structure
Dim hRgn As Long        ' handle to region where RECTs will be extracted from
Dim tRect As RECT       ' temporary RECT structure

' get the region data in RECT format
If fromWindow Then
    ' may ask: Why do you force user to say if source is Window or not when
    ' IsWindow() will tell you?  Answer: I have experienced IsWindow() to
    ' return true when the parameter is a region handle and not a Window handle.
    If IsWindow(hSource) Then
        ' if window handle is passed, get the region from that handle
        hRgn = CreateRectRgn(0, 0, 0, 0)
        If GetWindowRgn(hSource, hRgn) = 0 Then
            ' failed--this can happen if no region has been applied to
            ' the window. Therefore, the region is the window dimensions...
            GetWindowRect hSource, tRect
            hRgn = CreateRectRgn(0, 0, tRect.Right - tRect.Left + 1, tRect.Bottom - tRect.Top + 1)
            If hRgn = 0 Then Exit Function
        End If
    Else
        Exit Function
    End If
Else    ' use passed handle as a region handle
    hRgn = hSource
End If
rgnRects = ExtractRegionRects(hRgn, True)

' if above function fails, abort here
If UBound(rgnRects) < 0 Then GoTo ExitRoutine

Dim rtnRgn As Long      ' handle to region returned by this function
Dim Looper As Long      ' loop counter
Dim partialRgn As Long  ' used for 90 & 270 degree rotations only
Dim partialCt As Long   ' used for 90 & 270 degree rotations only
Dim rgnBytes() As RECT  ' used for rotations other than 90 or 270 degrees
Dim rowRect As RECT     ' marker to indicate new row of rectangles being accessed

' get the bounding rectangle, used for various purposes
GetRgnBox hRgn, rowRect

Select Case Rotation
Case 0, 2 ' rotate 90 or 270 degrees respectively
    ' this is a bit more involved....
    ' As mentioned earlier, RECTs must be stacked top to bottom, left to right
    ' That is how they are extracted from the region, and in order to rotate,
    ' we can't simply swap as done above because doing so would produce
    ' a left to right, top to bottom stack -- a no-no.  So, since I will be
    ' using the CombineRgn API, I don't really want to use it on each RECT,
    ' which would be slower. I can use it on an entire row
    ' which would be significantly quicker.
    ' For example: using a specific test image, there were 18,594 rectangles in
    '   the region, but only 256 rows. That means we Created & Destroyed
    '   18,338 less regions vs creating/destroying a region for each RECT
    
    ' add an extra RECT to mark the end of the loop
    ReDim Preserve rgnRects(0 To UBound(rgnRects) + 1)
    ' make the last RECT fall outside our bounding RECT
    SetRect rgnRects(UBound(rgnRects)), 0, rowRect.Bottom + 1, rowRect.Right, rowRect.Bottom + 2
    
    ' the 1st 2 RECTs are the header, so we start from 2 vs 0
    For Looper = 2 To UBound(rgnRects)
        With rgnRects(Looper)
            ' test for new row
            If .Top > rowRect.Top Then
                
                If partialCt Then
                    ' when we have a previous row rotated,
                    ' then do the partial region on that previous row
                    partialRgn = CreatePartialRegion(rgnRects(), 2, partialCt + 1, rowRect.Left, rgnRects(Looper - 1).Right)
                    
                    ' test for errors
                    If partialRgn = 0 Then
                        If rtnRgn Then DeleteObject rtnRgn
                        Exit Function
                    Else ' no error? then combine the region(s)
                        If rtnRgn Then
                            CombineRgn rtnRgn, rtnRgn, partialRgn, RGN_OR
                            DeleteObject partialRgn
                        Else
                            rtnRgn = partialRgn
                        End If
                    End If
                End If
                
                ' set up flag to indicate when new row is being accessed
                SetRect rowRect, rowRect.Left, .Top, rowRect.Right, rowRect.Bottom
                partialCt = 0   ' reset the rectangle counter
            End If
            
            ' depending on the rotation, calculate the rotated rectangle
            If Rotation Then  ' rotate 270
                SetRect rgnRects(partialCt + 2), .Top, rowRect.Right - .Right, .Bottom, rowRect.Right - .Left
            Else              ' rotate 90
                SetRect rgnRects(partialCt + 2), rowRect.Bottom - .Bottom, .Left, rowRect.Bottom - .Top, .Right
            End If
            ' increment the rectangle counter
            partialCt = partialCt + 1
        End With
    Next

Case 1, 3, 4 ' rotate 180, mirror horizontally/vertically respectively
    
    ' each of these are in their own loop to make loops faster & for debugging.
    ' Notes. Region rectangles are stacked. We will loop thru the appropriate
    '  row of rectangles on the source and resize them into the destination array.
    
    ' rowRect allows us to determine when we are moving to a new row of RECTs
    
    ' size new region bytes & copy the region header
    ReDim rgnBytes(0 To UBound(rgnRects))
    CopyMemory rgnBytes(0), rgnRects(0), 32&
    
    If Rotation = 3 Then ' mirror horizontal
        ' simply swap each row's RECTs, from right to left
        rowRect.Bottom = rowRect.Top - 1
        For Looper = 2 To UBound(rgnRects)
            With rgnRects(Looper)
                If .Top > rowRect.Bottom Then SetRect rowRect, rowRect.Left, .Top, rowRect.Right, .Bottom
                SetRect tRect, rowRect.Right - .Right, .Top, rowRect.Right - .Left, .Bottom
            End With
            rgnBytes(Looper) = tRect
        Next
        
    ElseIf Rotation = 4 Then ' mirror vertical
        ' simply swap each row's RECTs, from bottom to top
        rowRect.Top = rowRect.Bottom + 1
        For Looper = 2 To UBound(rgnRects)
            With rgnRects(UBound(rgnRects) - Looper + 2)
                If .Top < rowRect.Top Then SetRect rowRect, rowRect.Left, .Top, rowRect.Right, rowRect.Bottom
                SetRect tRect, .Left, rowRect.Bottom - .Bottom, .Right, rowRect.Bottom - .Top
            End With
            rgnBytes(Looper) = tRect
        Next
        
    Else    ' rotate 180
        ' swap each row's RECTs from bottom to top & from right to left
        rowRect.Top = rowRect.Bottom + 1
        For Looper = 2 To UBound(rgnRects)
            With rgnRects(UBound(rgnRects) - Looper + 2)
                If .Top < rowRect.Top Then SetRect rowRect, rowRect.Left, .Top, rowRect.Right, rowRect.Bottom
                SetRect tRect, rowRect.Right - .Right, rowRect.Bottom - .Bottom, rowRect.Right - .Left, rowRect.Bottom - .Top
            End With
            rgnBytes(Looper) = tRect
        Next
    End If
    ' create the new region
    rtnRgn = ExtCreateRegion(ByVal 0&, (UBound(rgnBytes) + 1) * 16, rgnBytes(0))
    
    ' ok, now to test whether or not we are good to go...
    ' if less than 2000 rectangles, API should have worked & if it didn't
    ' it wasn't due O/S restrictions -- failure
    If rtnRgn = 0 And UBound(rgnBytes) > 2000 Then
        rtnRgn = CreateWin98Region(rgnBytes, UBound(rgnBytes), rowRect.Left, rgnBytes(1).Right - rgnBytes(1).Left + 1)
    End If

Case Else ' invalid param, copy the region based of the bytes
    
    rtnRgn = ExtCreateRegion(ByVal 0&, (UBound(rgnRects) + 1) * 16, rgnRects(0))

    ' ok, now to test whether or not we are good to go...
    ' if less than 2000 rectangles, API should have worked & if it didn't
    ' it wasn't due O/S restrictions -- failure
    If rtnRgn = 0 And UBound(rgnRects) > 2000 Then
        rtnRgn = CreateWin98Region(rgnRects, UBound(rgnRects), rowRect.Left, rgnRects(1).Right - rgnRects(1).Left + 1)
    End If

End Select

Erase rgnRects()
Erase rgnBytes()

ExitRoutine:
' if we had to get a copy of a window region, then delete that region
If hSource <> hRgn Then DeleteObject hRgn

RotateSimpleRegion = rtnRgn
End Function

Public Function StretchRegion(hSrcRgn As Long, xScale As Single, yScale As Single) As Long

' Routine will stretch a region similar to how StretchBlt stretches bitmaps

' hSrcRgn is the region to be stretched
' xScale is percentage of increase or decrease in width
' yScale is percentage of increase or decrease in height
' (i.e., 1.5 for 50% increase and 0.5 for 50% decrease)

' One final note here. I did not take the time to modify my routines
' to handle a failed region in Win98 where the number of rectangles
' may exceed 4,000. If you plan on using this routine, strongly suggest
' modifying the CreateWin98Region & CreatePartialRegion functions to
' accept an XFORM parameter so you can create the region in pieces.
' Examples on how to call the CreateWin98Region are sprinkled throughout
' these routines.

If hSrcRgn = 0 Then Exit Function

Dim xRgn As Long, hBrush As Long, r As RECT

'// ' these are the only UDT members you can set and have function
'   compatible with Win98/Me

    Dim xFrm As XFORM
    With xFrm
        .eDx = 0
        .eDy = 0
        .eM11 = xScale
        .eM12 = 0
        .eM21 = 0
        .eM22 = yScale
    End With
    Dim hRgn As Long, dwCount As Long, pRgnData() As Byte
        
    ' get size of region to stretch
    dwCount = GetRegionData(hSrcRgn, 0, ByVal 0&)
    ' create a byte struction to hold that data and get that data
    ReDim pRgnData(0 To dwCount - 1) As Byte
    If dwCount = GetRegionData(hSrcRgn, dwCount, pRgnData(0)) Then
        ' create the stretched region
        hRgn = ExtCreateRegion(xFrm, dwCount, pRgnData(0))
        Erase pRgnData
    End If
StretchRegion = hRgn
End Function



