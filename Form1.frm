VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fingerprint prog"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   902
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   8790
      ScaleHeight     =   350
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   286
      TabIndex        =   7
      Top             =   30
      Width           =   4320
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command7"
      Height          =   525
      Left            =   9120
      TabIndex        =   6
      Top             =   5580
      Width           =   1245
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Smooth"
      Height          =   525
      Left            =   4560
      TabIndex        =   5
      Top             =   6090
      Width           =   1245
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Thin"
      Height          =   525
      Left            =   5910
      TabIndex        =   4
      Top             =   5520
      Width           =   1245
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sharpen"
      Height          =   525
      Left            =   4560
      TabIndex        =   3
      Top             =   5520
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Minutiae 3*3"
      Height          =   525
      Left            =   7260
      TabIndex        =   2
      Top             =   5520
      Width           =   1245
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   4410
      ScaleHeight     =   350
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   286
      TabIndex        =   1
      Top             =   30
      Width           =   4320
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5280
      Left            =   30
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   288
      TabIndex        =   0
      Top             =   30
      Width           =   4320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   690
      TabIndex        =   8
      Top             =   5400
      Width           =   2925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command2_Click()

Dim A1, A2, A3, A4, A5, A6, A7, A8, A9 As Integer
    
    For y = 50 To (Picture2.ScaleHeight - 50)
    For x = 50 To (Picture2.ScaleWidth - 50)
        A1 = GetBW(Picture2.Point(x, y))
        A2 = GetBW(Picture2.Point(x + 1, y))
        A3 = GetBW(Picture2.Point(x + 2, y))
        
        A4 = GetBW(Picture2.Point(x, y + 1))
        A5 = GetBW(Picture2.Point(x + 1, y + 1))
        A6 = GetBW(Picture2.Point(x + 2, y + 1))
        
        A7 = GetBW(Picture2.Point(x, y + 2))
        A8 = GetBW(Picture2.Point(x + 1, y + 2))
        A9 = GetBW(Picture2.Point(x + 2, y + 2))
            
           
           ' 110
           ' 110
           ' 001
           If A1 = 1 And A2 = 1 And A3 = 0 _
            And A4 = 1 And A5 = 1 And A6 = 0 _
            And A7 = 0 And A8 = 0 And A9 = 1 Then
            Picture3.Circle (x, y), 2, vbRed
           End If
           
           If A1 = 0 And A2 = 1 And A3 = 1 _
           And A4 = 0 And A5 = 1 And A6 = 1 _
           And A7 = 1 And A8 = 0 And A9 = 0 Then
            Picture3.Circle (x, y), 2, vbRed
           End If
           
           If A1 = 0 And A2 = 0 And A3 = 1 _
           And A4 = 1 And A5 = 1 And A6 = 0 _
           And A7 = 1 And A8 = 1 And A9 = 0 Then
            Picture3.Circle (x, y), 2, vbRed
           End If
           
           If A1 = 1 And A2 = 0 And A3 = 0 _
           And A4 = 0 And A5 = 1 And A6 = 1 _
           And A7 = 0 And A8 = 1 And A9 = 1 Then
            Picture3.Circle (x, y), 2, vbRed
           End If
           
           ' 100
           ' 010
           ' 000
           If A1 = 1 And A2 = 0 And A3 = 0 _
           And A4 = 0 And A5 = 1 And A6 = 0 _
           And A7 = 0 And A8 = 0 And A9 = 0 Then
            Picture3.Circle (x, y), 2, vbBlue
           End If
           
           If A1 = 0 And A2 = 0 And A3 = 1 _
           And A4 = 0 And A5 = 1 And A6 = 0 _
           And A7 = 0 And A8 = 0 And A9 = 0 Then
            Picture3.Circle (x, y), 2, vbBlue
           End If
            
            If A1 = 0 And A2 = 0 And A3 = 0 _
           And A4 = 0 And A5 = 1 And A6 = 0 _
           And A7 = 0 And A8 = 0 And A9 = 1 Then
            Picture3.Circle (x, y), 2, vbBlue
           End If
            
            If A1 = 0 And A2 = 0 And A3 = 0 _
           And A4 = 0 And A5 = 1 And A6 = 0 _
           And A7 = 1 And A8 = 0 And A9 = 0 Then
            Picture3.Circle (x, y), 2, vbBlue
           End If

Next
        DoEvents
    Next


Label1.Caption = "OK"

End Sub

Private Sub Command3_Click()
    
    For y = 0 To Picture1.Height - 1
        For x = 0 To Picture1.Width - 1
            c = Picture1.Point(x, y)
            h = Right("000000" & Hex(c), 6)
            r = CInt("&h" & Mid(h, 1, 2))
            g = CInt("&h" & Mid(h, 3, 2))
            b = CInt("&h" & Mid(h, 5, 2))
            A = (r + g + b) / 3
            n = (255 - A) / 255
              
            If A > 195 Then
            Picture2.PSet (x, y)
            Else
    '        Picture2.PSet (x, y)
            End If
            
        Next
        DoEvents
    Next
    
Label1.Caption = "OK"

 End Sub

Private Sub Command4_Click()
Call thinit(Picture2, Picture2.Width, Picture2.Height)
Label1.Caption = "Thin OK"

End Sub

Private Sub Command5_Click()
ProcessSmooth Picture2
End Sub

Private Sub Command7_Click()
   
    For y = 0 To Picture2.Height - 1
        For x = 0 To Picture2.Width - 1
            c = Picture2.Point(x, y)
            h = Right("000000" & Hex(c), 6)
            r = CInt("&h" & Mid(h, 1, 2))
            g = CInt("&h" & Mid(h, 3, 2))
            b = CInt("&h" & Mid(h, 5, 2))
            A = (r + g + b) / 3
            n = (255 - A) / 255
              
            If n = 1 Then
            Picture3.PSet (x - 1, y), vbRed
            Picture3.PSet (x + 1, y), vbBlue
            Picture3.PSet (x, y - 1), vbGreen
            Picture3.PSet (x, y + 1), vbYellow
   '         Else
   '         Picture2.PSet (x, y)
            End If
            
        Next
        DoEvents
    Next
    
    Label1.Caption = "OK"

End Sub
Function GetBW(Gaw As Variant) As Variant
            
            c = Gaw
            h = Right("000000" & Hex(c), 6)
            r = CInt("&h" & Mid(h, 1, 2))
            g = CInt("&h" & Mid(h, 3, 2))
            b = CInt("&h" & Mid(h, 5, 2))
            A = (r + g + b) / 3
            n = (255 - A) / 255
            
           If n > 0 Then
              b = (n * 100) \ 3
              If b > 10 Then GetBW = 1
          Else
          GetBW = n
          End If
           
End Function

