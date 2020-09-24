Attribute VB_Name = "modEffects"
'Effects module by Wiktor Toporek
'mail: witek1@konto.pl
'   or wtoporek@gmail.com

Public Buffer As PictureBox

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Enum eEnlargeMode
    CenterEnlarge = 0
    VerticalEnlarge = 1
    HorizontalEnlarge = 2
End Enum
Public Enum eLaserCorner
    RightBottomCorner = 0
    RightUpperCorner = 1
    LeftBottomCorner = 2
    LeftUpperCorner = 3
End Enum



Private Sub Wait(ms As Integer)
    DoEvents
    Sleep CLng(ms)
    DoEvents
End Sub
Public Sub BrickLayer(CenterX As Single, CenterY As Single, picSrc As PictureBox, picDest As PictureBox, Optional BrickSize As Integer = 32)
    picDest.Cls
    Dim eX As Single, eY As Single
    eX = CenterX - picSrc.ScaleWidth / 2
    eY = CenterY - picSrc.ScaleHeight / 2
    Dim pX As Integer, pY As Integer
    
    For pY = 0 To Int(picSrc.ScaleHeight / BrickSize)
        For pX = 0 To Int(picSrc.ScaleWidth / BrickSize)
            BitBlt picDest.hdc, eX + pX * BrickSize, eY + pY * BrickSize, BrickSize, BrickSize, picSrc.hdc, pX * BrickSize, pY * BrickSize, vbSrcCopy
            picDest.Refresh
            Wait 30
        Next
    Next
End Sub

Public Sub BlackBox(CenterX As Single, CenterY As Single, picSrc As PictureBox, picDest As PictureBox, Optional Speed As Integer = 5)
    picDest.Cls
    Dim eX As Single, eY As Single
    eX = CenterX - picSrc.ScaleWidth / 2
    eY = CenterY - picSrc.ScaleHeight / 2
    
    Dim BBoxWidth As Single
    Dim BBoxHeight As Single
    BBoxWidth = picSrc.ScaleWidth
    BBoxHeight = picSrc.ScaleHeight
    picDest.FillStyle = 0
    picDest.FillColor = 0
    
    Do While BBoxWidth > 0 And BBoxHeight > 0
        picDest.Cls
        BitBlt picDest.hdc, eX, eY, picSrc.ScaleWidth, picSrc.ScaleHeight, picSrc.hdc, 0, 0, vbSrcCopy
        Rectangle picDest.hdc, CenterX - BBoxWidth / 2, CenterY - BBoxHeight / 2, CenterX + BBoxWidth / 2, CenterY + BBoxHeight / 2
        Wait 20
        BBoxWidth = BBoxWidth - Speed
        BBoxHeight = BBoxHeight - Speed
    Loop
    BitBlt picDest.hdc, eX, eY, picSrc.ScaleWidth, picSrc.ScaleHeight, picSrc.hdc, 0, 0, vbSrcCopy
    picDest.Refresh
End Sub
Public Sub BlackCircle(CenterX As Single, CenterY As Single, picSrc As PictureBox, picDest As PictureBox, Optional Speed As Integer = 5)
    picDest.Cls
    Dim eX As Single, eY As Single
    eX = CenterX - picSrc.ScaleWidth / 2
    eY = CenterY - picSrc.ScaleHeight / 2
    
    Dim CircleSize As Single
    CircleSize = IIf(picSrc.ScaleWidth > picSrc.ScaleHeight, picSrc.ScaleWidth * 1.3, picSrc.ScaleHeight * 1.3)

    Buffer.Cls
    Buffer.BackColor = picDest.BackColor
    Buffer.FillStyle = 0
    Buffer.FillColor = 0
    Buffer.AutoSize = True
    
    Do While CircleSize > 0

        Buffer.Picture = picSrc.Image
        Ellipse Buffer.hdc, (Buffer.ScaleWidth / 2) - CircleSize / 2, (Buffer.ScaleHeight / 2) - CircleSize / 2, (Buffer.ScaleWidth / 2) + CircleSize / 2, (Buffer.ScaleHeight / 2) + CircleSize / 2
        picDest.Cls
        BitBlt picDest.hdc, eX, eY, Buffer.ScaleWidth, Buffer.ScaleHeight, Buffer.hdc, 0, 0, vbSrcCopy
        
        Wait 20
        CircleSize = CircleSize - Speed
    Loop
    BitBlt picDest.hdc, eX, eY, picSrc.ScaleWidth, picSrc.ScaleHeight, picSrc.hdc, 0, 0, vbSrcCopy
    picDest.Refresh
End Sub
Public Sub Laser(CenterX As Single, CenterY As Single, picSrc As PictureBox, picDest As PictureBox, Optional Precision As Integer = 5, Optional LaserCorner As eLaserCorner = 0)
    picDest.Cls
    Dim eX As Single, eY As Single
    eX = CenterX - picSrc.ScaleWidth / 2
    eY = CenterY - picSrc.ScaleHeight / 2
    
    Dim pX As Single, pY As Single
    
    For pX = 0 To picSrc.ScaleWidth - 1 Step 3
        picDest.Cls
        For pY = 0 To picSrc.ScaleHeight - 1 Step Precision
            picDest.ForeColor = GetPixel(picSrc.hdc, CLng(pX), CLng(pY))
            Select Case LaserCorner
                Case 0
                    picDest.Line (eX + pX, eY + pY)-(picDest.ScaleWidth, picDest.ScaleHeight)
                Case 1
                    picDest.Line (eX + pX, eY + pY)-(picDest.ScaleWidth, 0)
                Case 2
                    picDest.Line (eX + pX, eY + pY)-(0, picDest.ScaleHeight)
                Case 3
                    picDest.Line (eX + pX, eY + pY)-(0, 0)
            End Select
        Next
        BitBlt picDest.hdc, eX, eY, pX, picSrc.ScaleHeight, picSrc.hdc, 0, 0, vbSrcCopy
        picDest.Refresh
        Wait 10
    Next
    picDest.Cls
    BitBlt picDest.hdc, eX, eY, picSrc.ScaleWidth, picSrc.ScaleHeight, picSrc.hdc, 0, 0, vbSrcCopy
    picDest.Refresh

    
End Sub
Public Sub Checker(CenterX As Single, CenterY As Single, picSrc As PictureBox, picDest As PictureBox, Optional SquareSize As Integer = 32)
    picDest.Cls
    Dim eX As Single, eY As Single
    eX = Int(CenterX - picSrc.ScaleWidth / 2)
    eY = Int(CenterY - picSrc.ScaleHeight / 2)
    
    Dim Steps As Integer
    Dim YSteps As Integer
    Dim pX As Integer
    Dim pY As Integer
    picDest.Cls
    For Steps = 1 To 2
        If Steps = 2 Then YSteps = 1
        For pX = 0 To Int(picSrc.ScaleWidth / SquareSize)
            For pY = YSteps To Int(picSrc.ScaleHeight / SquareSize) Step 2
                BitBlt picDest.hdc, eX + pX * SquareSize, eY + pY * SquareSize, SquareSize, SquareSize, picSrc.hdc, pX * SquareSize, pY * SquareSize, vbSrcCopy
            Next
            picDest.Refresh
            Wait SquareSize * 5
            YSteps = 1 - YSteps
        Next
    Next
End Sub

Public Sub Enlarge(CenterX As Single, CenterY As Single, picSrc As PictureBox, picDest As PictureBox, Optional Mode As eEnlargeMode)
    Dim I As Integer
    Dim PicWidth As Single, PicHeight As Single
    For I = 1 To 100
        picDest.Cls
        
        PicWidth = IIf(Mode = 1, picSrc.ScaleWidth, picSrc.ScaleWidth * (I / 100))
        PicHeight = IIf(Mode = 2, picSrc.ScaleHeight, picSrc.ScaleHeight * (I / 100))
        
        
        StretchBlt picDest.hdc, CenterX - PicWidth / 2, CenterY - PicHeight / 2, PicWidth, PicHeight, picSrc.hdc, 0, 0, picSrc.ScaleWidth, picSrc.ScaleHeight, vbSrcCopy
        Wait 10
    Next
    
End Sub
Public Sub Slash(CenterX As Single, CenterY As Single, picSrc As PictureBox, picDest As PictureBox, Optional Speed As Integer = 5)
    Dim eX As Single, eY As Single
    eX = CenterX - picSrc.ScaleWidth / 2
    eY = CenterY - picSrc.ScaleHeight / 2
    
    Dim X As Single
    For X = 1 To picSrc.ScaleWidth Step Speed
        picDest.Cls
        BitBlt picDest.hdc, eX, eY, X, picSrc.ScaleHeight / 2, picSrc.hdc, 0, 0, vbSrcCopy
        BitBlt picDest.hdc, CenterX + picSrc.ScaleWidth / 2 - X, CenterY, X, picSrc.ScaleHeight / 2, picSrc.hdc, picSrc.ScaleWidth - X, picSrc.ScaleHeight / 2, vbSrcCopy
        Wait 20
    Next
    picDest.Cls
    BitBlt picDest.hdc, eX, eY, picSrc.ScaleWidth, picSrc.ScaleHeight, picSrc.hdc, 0, 0, vbSrcCopy

End Sub
Public Sub Emerge(CenterX As Single, CenterY As Single, picSrc As PictureBox, picDest As PictureBox, Optional Speed As Integer = 5)
    Dim eX As Single
    eX = CenterX - picSrc.ScaleWidth / 2
    
    Dim Y As Single
    For Y = 1 To picSrc.ScaleHeight Step Speed
        picDest.Cls
        BitBlt picDest.hdc, eX, CenterY - Int(Y / 2), picSrc.ScaleWidth, Y, picSrc.hdc, 0, 0, vbSrcCopy
        Wait 20
    Next
    picDest.Cls
    BitBlt picDest.hdc, eX, CenterY - picSrc.ScaleHeight / 2, picSrc.ScaleWidth, picSrc.ScaleHeight, picSrc.hdc, 0, 0, vbSrcCopy

End Sub
