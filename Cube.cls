VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**********************************************
'Author - Mark Gordon (msg555)
'Member of VbCity.com
'Date September 19th, 2004
'Copyright Mark Gordon. 2004
'
'This class is free for use in any application
'**********************************************

Private Type Point3D
    x As Long
    y As Long
    z As Long
End Type
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
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function CreateBitmap Lib "GDI32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePolygonRgn Lib "GDI32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "GDI32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateDCAsNull Lib "GDI32" Alias "CreateDCA" (ByVal lpDriverName As String, _
           lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateDC Lib "GDI32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObjectAPI Lib "GDI32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

'*******************************
'Internal Sizes
'*******************************
Private cWIDTH As Long
Private cHEIGHT As Long
Private cDEPTH As Long

'*******************************
'Constants used for Rotation
'*******************************
'Private Const Pi = 3.14159265
Private Const SPIN = 0.3 / 180

'*******************************
'These represent the ORIGINAL Verticies of the object
'*******************************
Private Vert(7) As Point3D

'*******************************
'Array of Collections that holds the Coordinates of each Face.  See CreateCube()
'*******************************
Private Faces(5) As Collection

'Public Variables

'*******************************
'Roll
'   Determines The Rotation around the Z axis
'Pitch
'   Determines The Rotation around the X axis
'Yaw
'   Determines The Rotation around the Y axis
'*******************************
Public Roll As Double, Pitch As Double, Yaw As Double

'*******************************
'All values are the Center Positions of the Cube
'Z determines how close the Cube appears
'The higher Z, the Further away the Cube is
'*******************************
Public x As Long, y As Long, z As Long

'*******************************
'DoubleBuffer
'   All Faces are Draw in Memory, and then are drawn to the specified Device Context (hDc) ounce
'TransparentBackground
'   Determines if the Background color should be drawn to the specified Device Context (hDc).
'   Note.  The Background color has to be white for drawing purposes
'*******************************
Public DoubleBuffer As Boolean, TransparentBackground As Boolean

'*******************************
'GDI Variables

    'Memory Dc's
    Private MemoryDc As Long
    Private TempDC As Long
    Private MonoMaskDC As Long
    
    'Bitmaps
    Private hMonoMask As Long
    
    'Width and Height of the Memory DC's.  Equal to the Maximum distance on point in the cube could be from another
    Private MemorySize As Long
    
    'Brushes
    Private WhiteBrsh As Long
    '*******************************
    'SideColor
    '   Holds the Color of the Side of the Face(FaceConst)
    'SidBrsh
    '   Holds the Brush Object for the Face(FaceConst)
    '*******************************
    Private SideColor(5) As Long
    Private SideBrsh(5) As Long
'*******************************

'Used purely for Intellesence purposes
Public Enum FaceConst
    One = 0
    Two = 1
    Three = 2
    Four = 3
    Five = 4
    Six = 5
End Enum


'*******************************
'Draws the Cube at the position given by the cube's x, y, and z properties
'   with rotation given by Roll, Pitch, and Yaw properties
'*******************************
Public Sub Draw(hDC As Long)
Dim TwoD(7) As POINTAPI

Dim LowZ As Long

'*******************************
'This was a rather ingenious method of figuring out which faces to draw
'The ConvertVertsto2d function returns what the index of the closest Z Value
'Knowing this we can assume that any Face not containing that point is should not be drawn
'Only three faces can be seen on a Cube
'*******************************
LowZ = CovertVertsto2d(TwoD) 'Note that it returns the index of the Closest Z, not the value

Dim Rc As RECT
If DoubleBuffer Then
    'Redraw Background
    Rc.Right = MemorySize
    Rc.Bottom = MemorySize
    FillRect MemoryDc, Rc, WhiteBrsh
ElseIf Not TransparentBackground Then
    'Draw Background Color if not double buffering
    With Rc
        .Left = Me.x - MemorySize / 2
        .Right = .Left + MemorySize
        .Top = Me.y - MemorySize / 2
        .Bottom = .Top + MemorySize
    End With
    FillRect hDC, Rc, WhiteBrsh
End If
    
    
Dim i As Long
For i = 0 To 5
    If MemberInCollection(LowZ, Faces(i)) Then
        DrawFace hDC, Faces(i), TwoD, i
    End If
Next

If DoubleBuffer Then
    Dim DstLeft As Long, DstTop As Long
    DstLeft = Me.x - MemorySize / 2
    DstTop = Me.y - MemorySize / 2
    If TransparentBackground Then
        'Get B&W Mask.  Black Represents the Cube and White Represents the Background
        BitBlt MonoMaskDC, 0, 0, MemorySize, MemorySize, MemoryDc, 0, 0, vbSrcCopy
        
        'Draw the Background on TempDC
        BitBlt TempDC, 0, 0, MemorySize, MemorySize, hDC, DstLeft, DstTop, vbSrcCopy
        
        'Draw the Mask onto the Target DC
        BitBlt TempDC, 0, 0, MemorySize, MemorySize, MonoMaskDC, 0, 0, vbSrcAnd
    
        'Get the Negative of MonoMask
        BitBlt MonoMaskDC, 0, 0, MemorySize, MemorySize, MonoMaskDC, 0, 0, vbNotSrcCopy
        
        'Change MemoryDC's Background Color
        BitBlt MemoryDc, 0, 0, MemorySize, MemorySize, MonoMaskDC, 0, 0, vbSrcAnd
        
        'Draw MemoryDC Normally
        BitBlt TempDC, 0, 0, MemorySize, MemorySize, MemoryDc, 0, 0, vbSrcPaint
        
        'Finnally Draw The Contents
        BitBlt hDC, DstLeft, DstTop, MemorySize, MemorySize, TempDC, 0, 0, vbSrcCopy
    Else
        BitBlt hDC, DstLeft, DstTop, MemorySize, MemorySize, MemoryDc, 0, 0, vbSrcCopy
    End If
End If
End Sub

'Returns the Index of the Lowest (Closest) Z Value
Private Function CovertVertsto2d(Point2D() As POINTAPI) As Long
Dim LowZ As Double

Dim i As Long
Dim CosX As Double, CosY As Double, CosZ As Double, SinX As Double, SinY As Double, SinZ As Double
CosX = Cos(SPIN * Pitch)
SinX = Sin(SPIN * Pitch)
CosY = Cos(SPIN * Yaw)
SinY = Sin(SPIN * Yaw)
CosZ = Cos(SPIN * Roll)
SinZ = Sin(SPIN * Roll)

Dim t As Point3D
Dim n As Point3D
For i = 0 To 7
    n = Vert(i)
    With n
        t = n
        .x = CosZ * t.x - SinZ * t.y
        .y = CosZ * t.y + SinZ * t.x
        
        t = n
        .x = CosY * t.x - SinY * t.z
        .z = CosY * t.z + SinY * t.x
        
        t = n
        .y = CosX * t.y - SinX * t.z
        .z = CosX * t.z + SinX * t.y
        
        'Add Depth and Coordinates
        Point2D(i).x = n.x * 0.999 ^ (n.z + Me.z) + Me.x
        Point2D(i).y = n.y * 0.999 ^ (n.z + Me.z) + Me.y
        
        If i = 0 Then
            CovertVertsto2d = 0
            LowZ = n.z
        ElseIf n.z < LowZ Then
            CovertVertsto2d = i
            LowZ = n.z
        End If
    End With
Next
End Function

Private Sub Class_Initialize()
Me.TransparentBackground = True
Me.DoubleBuffer = True

Randomize
SetColor 0, Rnd * &HFFFFFF
SetColor 1, Rnd * &HFFFFFF
SetColor 2, Rnd * &HFFFFFF
SetColor 3, Rnd * &HFFFFFF
SetColor 4, Rnd * &HFFFFFF
SetColor 5, Rnd * &HFFFFFF

cWIDTH = 100
cHEIGHT = 100
cDEPTH = 100

CreateCube
End Sub

Private Sub CreateCube()
Dim i As Long
For i = 0 To 7
    With Vert(i)
        .x = IIf(i Mod 4 = 0 Or i Mod 4 = 3, -cWIDTH / 2, cWIDTH / 2)
        .y = IIf(i Mod 4 = 0 Or i Mod 4 = 1, -cHEIGHT / 2, cHEIGHT / 2)
        .z = IIf(i <= 3, -cDEPTH / 2, cDEPTH / 2)
    End With
Next

'Create Face Objects (Collections)
Dim n As Collection

'Face 1
Set n = New Collection
n.Add 0
n.Add 1
n.Add 2
n.Add 3
Set Faces(0) = n

'Face 2
Set n = New Collection
n.Add 0
n.Add 1
n.Add 5
n.Add 4
Set Faces(1) = n

'Face 3
Set n = New Collection
n.Add 0
n.Add 3
n.Add 7
n.Add 4
Set Faces(2) = n

'Face 4
Set n = New Collection
n.Add 1
n.Add 2
n.Add 6
n.Add 5
Set Faces(3) = n

'Face 5
Set n = New Collection
n.Add 2
n.Add 3
n.Add 7
n.Add 6
Set Faces(4) = n

'Face 6
Set n = New Collection
n.Add 4
n.Add 5
n.Add 6
n.Add 7
Set Faces(5) = n

'Yayy!! Memory Dc time
DeleteDC MemoryDc

MemorySize = Sqr(cWIDTH ^ 2 + cHEIGHT ^ 2 + cDEPTH ^ 2)
MemoryDc = newDc(MemorySize, MemorySize)

TempDC = newDc(MemorySize, MemorySize)

MonoMaskDC = CreateCompatibleDC(MemoryDc)
hMonoMask = CreateBitmap(MemorySize, MemorySize, 1, 1, ByVal 0&)
SelectObject MonoMaskDC, hMonoMask

WhiteBrsh = CreateSolidBrush(RGB(127, 127, 0))
'All Done :)
End Sub

Private Function Dist(p1 As Point3D, p2 As Point3D) As Double
Dist = Sqr((p1.x - p2.x) ^ 2 + (p1.y - p2.y) ^ 2 + (p1.z - p2.z) ^ 2)
End Function

Private Sub DrawFace(hDC As Long, Face As Collection, P2D() As POINTAPI, SideNumber As Long)
Dim Pts(3) As POINTAPI
Pts(0) = P2D(Face(1))
Pts(1) = P2D(Face(2))
Pts(2) = P2D(Face(3))
Pts(3) = P2D(Face(4))

Dim Rgn As Long

If DoubleBuffer Then
    'MovePoints About
    Dim i As Long
    For i = 0 To 3
        With Pts(i)
            .x = .x - Me.x + MemorySize / 2
            .y = .y - Me.y + MemorySize / 2
        End With
    Next
    
    Rgn = CreatePolygonRgn(Pts(0), UBound(Pts) + 1, 1)
    
    FillRgn MemoryDc, Rgn, SideBrsh(SideNumber)
Else
    Rgn = CreatePolygonRgn(Pts(0), UBound(Pts) + 1, 1)
    FillRgn hDC, Rgn, SideBrsh(SideNumber)
End If

DeleteObject Rgn
End Sub

Private Function MemberInCollection(Member, C As Collection) As Boolean
Dim i As Long
For i = 1 To C.Count
    If Member = C(i) Then
        MemberInCollection = True
        Exit Function
    End If
Next
End Function

Public Sub SetColor(Face As FaceConst, Value As OLE_COLOR)
If Face < One Or Face > Six Then Debug.Assert False
SideColor(Face) = Value
DeleteObject SideBrsh(Face)
SideBrsh(Face) = CreateSolidBrush(Value)
End Sub

Public Function GetColor(Face As FaceConst) As OLE_COLOR
If Face < 0 Or Face > 5 Then Debug.Assert False
GetColor = SideColor(Face)
End Function

Private Function newDc(ByVal Width As Long, ByVal Height As Long) As Long
Dim m_hDC As Long
Dim m_hBmp As Long
Dim m_hBmpOld
If Width And Height Then
    Dim lhDCC As Long
    lhDCC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
    If Not (lhDCC = 0) Then
       newDc = CreateCompatibleDC(lhDCC)
       If Not (newDc = 0) Then
          m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
          If Not (m_hBmp = 0) Then
             m_hBmpOld = SelectObject(newDc, m_hBmp)
             If Not (m_hBmpOld = 0) Then
                DeleteDC lhDCC
                Exit Function
             End If
          End If
       End If
       DeleteDC lhDCC
    End If
    
    newDc = 0
    'Insufficient Memory
End If
End Function

Private Sub Class_Terminate()
DeleteDC MemoryDc
DeleteDC MonoMaskDC
DeleteDC TempDC

DeleteObject WhiteBrsh

Dim i As Long
For i = 0 To 5
    DeleteObject SideBrsh(i)
Next
End Sub

Public Property Get Width() As Long
Width = cWIDTH
End Property
Public Property Let Width(n As Long)
cWIDTH = n
CreateCube
End Property

Public Property Get Height() As Long
Height = cHEIGHT
End Property
Public Property Let Height(n As Long)
cHEIGHT = n
CreateCube
End Property

Public Property Get Depth() As Long
Depth = cDEPTH
End Property
Public Property Let Depth(n As Long)
cDEPTH = n
CreateCube
End Property
