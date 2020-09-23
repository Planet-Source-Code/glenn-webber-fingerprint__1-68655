Attribute VB_Name = "Module10"
Option Explicit

Dim i As Integer, j As Integer
Dim red As Integer, green As Integer, blue As Integer
Dim fi As Integer, fj As Integer
Dim RedSum As Integer, GreenSum As Integer, BlueSum As Integer
Dim weight As Single
Dim offset As Integer
Dim x As Integer
Dim y As Integer
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Dim ImagePixels(0 To 2, 0 To 600, 0 To 600) As Integer

Dim m As Integer
Dim mi As Integer
Dim mj As Integer
Dim temp As Integer
Dim pixel As Long
Dim reda(0 To 8) As Integer
Dim greena(0 To 8) As Integer
Dim bluea(0 To 8) As Integer
Dim sum(0 To 1, 0 To 8) As Integer
Public Sub readit(ByRef Picture1 As PictureBox)
x = Picture1.ScaleWidth
y = Picture1.ScaleHeight
    For i = 0 To y - 1
        For j = 0 To x - 1
            pixel = GetPixel(Picture1.hdc, j, i)
            red = pixel Mod 256
            green = ((pixel And &HFF00) / 256&) Mod 256&
            blue = (pixel And &HFF0000) / 65536
            ImagePixels(0, i, j) = red
            ImagePixels(1, i, j) = green
            ImagePixels(2, i, j) = blue
        Next
        DoEvents
    Next
End Sub
Public Sub ProcessCustom3(ByRef Picture1 As PictureBox, ByRef customfilter() As Integer, filternorm As Integer, filterbias As Integer, takeabs As Boolean)
    
    readit Picture1
    
    offset = 1
    For i = offset To y - offset - 1
        For j = offset To x - offset - 1
            RedSum = 0: GreenSum = 0: BlueSum = 0
            For fi = -offset To offset
                For fj = -offset To offset
                    weight = customfilter(fi + 1, fj + 1)
                    RedSum = RedSum + ImagePixels(0, i + fi, j + fj) * weight
                    GreenSum = GreenSum + ImagePixels(1, i + fi, j + fj) * weight
                    BlueSum = BlueSum + ImagePixels(2, i + fi, j + fj) * weight
                Next
            Next
            If takeabs = True Then
            red = Abs(RedSum / filternorm + filterbias)
            green = Abs(GreenSum / filternorm + filterbias)
            blue = Abs(BlueSum / filternorm + filterbias)
            Else
            red = (RedSum / filternorm + filterbias)
            green = (GreenSum / filternorm + filterbias)
            blue = (BlueSum / filternorm + filterbias)
            If red > 255 Then
            red = 255
            Else
            If red < 0 Then red = 0
            End If
            If green > 255 Then
            green = 255
            Else
            If green < 0 Then green = 0
            End If
            If blue > 255 Then
            blue = 255
            Else
            If blue < 0 Then blue = 0
            End If
            End If
            SetPixelV Picture1.hdc, j, i, RGB(red, green, blue)
        Next
DoEvents
    Next

End Sub
Public Sub ProcessCustom5(ByRef Picture1 As PictureBox, ByRef customfilter() As Integer, filternorm As Integer, filterbias As Integer, takeabs As Boolean)
    
    readit Picture1
    
    offset = 2
    For i = offset To y - offset - 1
        For j = offset To x - offset - 1
            RedSum = 0: GreenSum = 0: BlueSum = 0
            For fi = -offset To offset
                For fj = -offset To offset
                    weight = customfilter(fi + 2, fj + 2)
                    RedSum = RedSum + ImagePixels(0, i + fi, j + fj) * weight
                    GreenSum = GreenSum + ImagePixels(1, i + fi, j + fj) * weight
                    BlueSum = BlueSum + ImagePixels(2, i + fi, j + fj) * weight
                Next
            Next
            If takeabs = True Then
            red = Abs(RedSum / filternorm + filterbias)
            green = Abs(GreenSum / filternorm + filterbias)
            blue = Abs(BlueSum / filternorm + filterbias)
            Else
            red = (RedSum / filternorm + filterbias)
            green = (GreenSum / filternorm + filterbias)
            blue = (BlueSum / filternorm + filterbias)
            If red > 255 Then
            red = 255
            Else
            If red < 0 Then red = 0
            End If
            If green > 255 Then
            green = 255
            Else
            If green < 0 Then green = 0
            End If
            If blue > 255 Then
            blue = 255
            Else
            If blue < 0 Then blue = 0
            End If
            End If
            SetPixelV Picture1.hdc, j, i, RGB(red, green, blue)
        Next
DoEvents
    Next

End Sub
Public Sub processmedian(ByRef Picture1 As PictureBox, ByVal filterbias As Integer, ByVal takeabs As Boolean)
    readit Picture1
    
    offset = 1
    For i = offset To y - offset - 1
        For j = offset To x - offset - 1
            m = 0
            For fi = -offset To offset
                For fj = -offset To offset
                    reda(fi + fj + 2 + m) = ImagePixels(0, i + fi, j + fj)
                    greena(fi + fj + 2 + m) = ImagePixels(1, i + fi, j + fj)
                    bluea(fi + fj + 2 + m) = ImagePixels(2, i + fi, j + fj)
                Next
                m = m + 2
            Next
            For mi = 0 To 8
                For mj = mi To 7
                    If reda(mj) > reda(mj + 1) Then
                    temp = reda(mj)
                    reda(mj) = reda(mj + 1)
                    reda(mj + 1) = reda(mj)
                    End If
                    If greena(mj) > greena(mj + 1) Then
                    temp = greena(mj)
                    greena(mj) = greena(mj + 1)
                    greena(mj + 1) = greena(mj)
                    End If
                    If bluea(mj) > bluea(mj + 1) Then
                    temp = bluea(mj)
                    bluea(mj) = bluea(mj + 1)
                    bluea(mj + 1) = bluea(mj)
                    End If
                 Next
            Next
            If takeabs = True Then
            red = Abs(reda(4) + filterbias)
            green = Abs(greena(4) + filterbias)
            blue = Abs(bluea(4) + filterbias)
            Else
            red = (reda(4) + filterbias)
            green = (greena(4) + filterbias)
            blue = (bluea(4) + filterbias)
            If red > 255 Then
            red = 255
            Else
            If red < 0 Then red = 0
            End If
            If green > 255 Then
            green = 255
            Else
            If green < 0 Then green = 0
            End If
            If blue > 255 Then
            blue = 255
            Else
            If blue < 0 Then blue = 0
            End If
            End If
            SetPixelV Picture1.hdc, j, i, RGB(red, green, blue)
        Next
DoEvents
    Next

End Sub
Public Sub processmediankc(ByRef Picture1 As PictureBox, ByVal filterbias As Integer, ByVal takeabs As Boolean)
    readit Picture1
    
    offset = 1
    For i = offset To y - offset - 1
        For j = offset To x - offset - 1
            m = 0
            For fi = -offset To offset
                For fj = -offset To offset
                    reda(fi + fj + 2 + m) = ImagePixels(0, i + fi, j + fj)
                    greena(fi + fj + 2 + m) = ImagePixels(1, i + fi, j + fj)
                    bluea(fi + fj + 2 + m) = ImagePixels(2, i + fi, j + fj)
                Next
                m = m + 2
            Next
            For mi = 0 To 8
                sum(0, mi) = reda(mi) + greena(mi) + bluea(mi)
                sum(1, mi) = mi
            Next
            For mi = 0 To 8
                For mj = mi To 7
                    If sum(0, mj) > sum(0, mj + 1) Then
                    temp = sum(0, mj)
                    sum(0, mj) = sum(0, mj + 1)
                    sum(0, mj + 1) = sum(0, mj)
                    temp = sum(1, mj)
                    sum(1, mj) = sum(1, mj + 1)
                    sum(1, mj + 1) = sum(1, mj)
                    End If
                 Next
            Next
            If takeabs = True Then
            red = Abs(reda(sum(1, 4)) + filterbias)
            green = Abs(greena(sum(1, 4)) + filterbias)
            blue = Abs(bluea(sum(1, 4)) + filterbias)
            Else
            red = (reda(sum(1, 4)) + filterbias)
            green = (greena(sum(1, 4)) + filterbias)
            blue = (bluea(sum(1, 4)) + filterbias)
            If red > 255 Then
            red = 255
            Else
            If red < 0 Then red = 0
            End If
            If green > 255 Then
            green = 255
            Else
            If green < 0 Then green = 0
            End If
            If blue > 255 Then
            blue = 255
            Else
            If blue < 0 Then blue = 0
            End If
            End If
            SetPixelV Picture1.hdc, j, i, RGB(red, green, blue)
        Next
    DoEvents
Next

End Sub
Public Sub ProcessDiffuse(ByRef Picture1 As PictureBox, ByVal rndinto As Integer, ByVal rndminus As Integer)

Dim Rx As Integer, Ry As Integer
   
    readit Picture1
    For i = 2 To y - 3
        For j = 2 To x - 3
            Rx = Rnd * rndinto - rndminus '4 - 2
            Ry = Rnd * rndinto - rndminus '4 - 2
            red = ImagePixels(0, i + Rx, j + Ry)
            green = ImagePixels(1, i + Rx, j + Ry)
            blue = ImagePixels(2, i + Rx, j + Ry)
            SetPixelV Picture1.hdc, j, i, RGB(red, green, blue)
        Next
        DoEvents
    Next

End Sub

Public Sub ProcessEmboss(ByRef Picture1 As PictureBox, ByVal filterbias As Integer)

Dim Dx As Integer, Dy As Integer
    readit Picture1

    Dx = 1
    Dy = 1
    
    
    'T1 = Timer
    For i = 1 To y - 2
        For j = 1 To x - 2
            red = Abs(ImagePixels(0, i, j) - ImagePixels(0, i + Dx, j + Dy) + filterbias) '128)
            green = Abs(ImagePixels(1, i, j) - ImagePixels(1, i + Dx, j + Dy) + filterbias) '128)
            blue = Abs(ImagePixels(2, i, j) - ImagePixels(2, i + Dx, j + Dy) + filterbias) '128)
            SetPixelV Picture1.hdc, j, i, RGB(red, green, blue)
        Next
        DoEvents
    Next
    
End Sub

Public Sub ProcessPixelize(ByRef Picture1 As PictureBox, ByVal rndplus As Integer, ByVal rndminus As Integer, ByVal intoradius As Integer, ByVal minradius As Integer)

Dim Ypixel As Integer, Xpixel As Integer
Dim r As Integer
    readit Picture1

    'T1 = Timer
    Picture1.FillStyle = vbSolid
    For i = 1 To y / 3
        For j = 1 To x / 3
            Ypixel = Rnd * x + rndplus - rndminus ' 4 - 2
            Xpixel = Rnd * y + rndplus - rndminus '4 - 2
            r = Int(Rnd() * intoradius) + minradius '3'2
            red = ImagePixels(0, Xpixel, Ypixel)
            green = ImagePixels(1, Xpixel, Ypixel)
            blue = ImagePixels(2, Xpixel, Ypixel)
            Picture1.FillColor = RGB(red, green, blue)
            Picture1.Circle (Ypixel, Xpixel), r, RGB(red, green, blue)
        Next
        DoEvents
        Picture1.Refresh
    Next
    Picture1.FillStyle = vbTransparent

End Sub

Public Sub ProcessSharpen(ByRef Picture1 As PictureBox)
Dim Dx As Integer, Dy As Integer
    
    readit Picture1

    Dx = 1: Dy = 1
    
    For i = 1 To y - 2
        For j = 1 To x - 2
            red = ImagePixels(0, i, j) + 0.5 * (ImagePixels(0, i, j) - ImagePixels(0, i - Dx, j - Dy))
            green = ImagePixels(1, i, j) + 0.5 * (ImagePixels(1, i, j) - ImagePixels(1, i - Dx, j - Dy))
            blue = ImagePixels(2, i, j) + 0.5 * (ImagePixels(2, i, j) - ImagePixels(2, i - Dx, j - Dy))
            If red > 255 Then red = 255
            If red < 0 Then red = 0
            If green > 255 Then green = 255
            If green < 0 Then green = 0
            If blue > 255 Then blue = 255
            If blue < 0 Then blue = 0
            SetPixelV Picture1.hdc, j, i, RGB(red, green, blue)
        Next
        DoEvents
    Next
 
End Sub



Public Sub ProcessSmooth(ByRef Picture1 As PictureBox)
    
    readit Picture1

    For i = 1 To y - 2
        For j = 1 To x - 2
            red = ImagePixels(0, i - 1, j - 1) + ImagePixels(0, i - 1, j) + ImagePixels(0, i - 1, j + 1) + _
            ImagePixels(0, i, j - 1) + ImagePixels(0, i, j) + ImagePixels(0, i, j + 1) + _
            ImagePixels(0, i + 1, j - 1) + ImagePixels(0, i + 1, j) + ImagePixels(0, i + 1, j + 1)
            
            green = ImagePixels(1, i - 1, j - 1) + ImagePixels(1, i - 1, j) + ImagePixels(1, i - 1, j + 1) + _
            ImagePixels(1, i, j - 1) + ImagePixels(1, i, j) + ImagePixels(1, i, j + 1) + _
            ImagePixels(1, i + 1, j - 1) + ImagePixels(1, i + 1, j) + ImagePixels(1, i + 1, j + 1)
            
            blue = ImagePixels(2, i - 1, j - 1) + ImagePixels(2, i - 1, j) + ImagePixels(2, i - 1, j + 1) + _
            ImagePixels(2, i, j - 1) + ImagePixels(2, i, j) + ImagePixels(2, i, j + 1) + _
            ImagePixels(2, i + 1, j - 1) + ImagePixels(2, i + 1, j) + ImagePixels(2, i + 1, j + 1)
            
            SetPixelV Picture1.hdc, j, i, RGB(red / 9, green / 9, blue / 9)
        Next
        DoEvents
    Next
 
End Sub

Public Sub ProcessSolarize(ByRef Picture1 As PictureBox, ByVal ll As Integer, ByVal ul As Integer)
    readit Picture1
    For i = 1 To y - 2
        For j = 1 To x - 2
            red = ImagePixels(0, i, j)
            green = ImagePixels(1, i, j)
            blue = ImagePixels(2, i, j)
            If ((red < ll) Or (red > ul)) Then red = 255 - red
            If ((green < ll) Or (green > ul)) Then green = 255 - green
            If ((blue < ll) Or (blue > ul)) Then blue = 255 - blue
            SetPixelV Picture1.hdc, j, i, RGB(red, green, blue)
        Next
        DoEvents
    Next

End Sub







