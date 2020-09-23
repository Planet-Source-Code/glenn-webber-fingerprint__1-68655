Attribute VB_Name = "Module2"
Dim ip(0 To 6000, 0 To 6000) As Integer
Dim temp(0 To 2, 0 To 2) As Integer
Dim i As Integer
Dim j As Integer
Dim A As Integer
Dim b As Integer
Dim x As Integer
Dim y As Integer
Dim nf As Integer
Dim red As Integer
Dim pixel As Long
Dim cnt As Integer
Dim bc As Integer
Dim non As Integer
Dim c As Integer
Dim d As Integer
Dim f As Integer
Dim black As Boolean
Dim check1, check2, check3, check4 As Boolean
Dim c1, c2, c3, c4 As Boolean
Dim repeat As Integer
Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'Dim ctl As UserControl1
Public Sub thinit(Picture1 As PictureBox, ByVal w As Integer, ByVal h As Integer) 'Picture1 As PictureBox)

x = w
y = h
For i = 0 To x - 1
        For j = 0 To y - 1
            pixel = Picture1.Point(i, j)
            red = pixel& Mod 256
            If red = 0 Then
            ip(i, j) = 1
            '1==black
            Else
            ip(i, j) = 0
            '0==white
            End If
        Next j
    DoEvents
    Next i
thin Picture1
End Sub
Private Sub thin(Picture1 As PictureBox)

check1 = False
check2 = False
check3 = False
check4 = False

c1 = True
c2 = True
c3 = True
c4 = True
For repeat = 1 To 10

black = True
 If ((check1 = False) And (c1 = True)) Then
 For i = 1 To x - 1 Step 1     ' t to b
    For j = 1 To y - 1 Step 1
    If black = True Then
    If ((ip(i, j) = 1 And ip(i, j - 1) = 0)) Then
                 If (canthin() = 1) Then
                      ip(i, j) = 0
                      black = False
                      check1 = True
                 End If
    End If
    Else
    If ip(i, j) = 0 Then
    black = True
    End If
    End If
Next j
DoEvents
Next i
End If

' next iteration
black = True
If ((check2 = False) And (c2 = True)) Then
For i = 1 To x - 1  'b to t
    For j = y - 1 To 1 Step -1
      If black = True Then
           If (ip(i, j) = 1 And ip(i, j + 1) = 0) Then
                 If (canthin() = 1) Then
                      ip(i, j) = 0
                      black = False
                      check2 = True
                      End If
           End If
    Else
    If ip(i, j) = 0 Then
    black = True
    End If
    End If
Next j
DoEvents
Next i
End If

'next iteration
black = True
If ((check3 = False) And (c3 = True)) Then
For j = 1 To y - 1 Step 1 'l to r
    For i = 1 To x - 1
    If black = True Then
         If (ip(i, j) = 1 And ip(i - 1, j) = 0) Then
                          If (canthin() = 1) Then
                                 ip(i, j) = 0
                                 black = False
                                 check3 = True
                           End If
                    End If
    Else
    If ip(i, j) = 0 Then
    black = True
    End If
    End If
Next i
DoEvents
Next j
End If


'next iteration
black = True
If ((check4 = False) And (c4 = True)) Then
For j = 1 To y - 1 ' r to l
    For i = x - 1 To 1 Step -1
        If black = True Then
             If (ip(i + 1, j) = 0 And ip(i, j) = 1) Then
                     If (canthin() = 1) Then
                        ip(i, j) = 0
                        black = False
                        check4 = True
                    End If
            End If
     Else
    If ip(i, j) = 0 Then
    black = True
    End If
    End If
Next i
DoEvents
Next j
End If

If check1 = True Then
check1 = False
Else
c1 = False
End If

If check2 = True Then
check2 = False
Else
c2 = False
End If

If check3 = True Then
check3 = False
Else
c3 = False
End If

If check4 = True Then
check4 = False
Else
c4 = False
End If

If ((c1 = False) And (c2 = False) And (c3 = False) And (c4 = False)) Then
Exit For
End If

Next repeat

For i = 0 To x - 1
For j = 0 To y - 1
If ip(i, j) = 1 Then
            SetPixelV Picture1.hdc, i, j, vbBlack
Else
            SetPixelV Picture1.hdc, i, j, vbWhite
End If
Next j
DoEvents
Next i

End Sub
Function canthin() As Integer

For A = -1 To 1
    For b = -1 To 1
            temp(A + 1, b + 1) = ip(i + A, j + b)
            Next b
            Next A
temp(1, 1) = 0

For A = 0 To 2
    For b = 0 To 2
        If temp(A, b) = 1 Then
        temp(A, b) = 0
        non = 1
        For c = -1 To 1
          For d = -1 To 1
              If (((A + c) >= 0) And ((A + c) <= 2) And ((d + b) >= 0) And ((b + d) <= 2)) Then
                  If (temp(A + c, b + d) = 1) Then
                   non = 0
                  End If
              End If
           Next d
       Next c
    temp(A, b) = 1
       If (non) Then
            canthin = 0
            Exit Function
       End If
       End If
   Next b
   Next A
 canthin = 1
End Function



