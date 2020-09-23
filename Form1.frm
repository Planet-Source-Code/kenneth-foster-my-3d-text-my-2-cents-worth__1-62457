VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "3D Text"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1725
      TabIndex        =   21
      Top             =   4440
      Width           =   2955
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   75
      ScaleHeight     =   1605
      ScaleWidth      =   7245
      TabIndex        =   20
      Top             =   5460
      Width           =   7245
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save as bitmap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4755
      TabIndex        =   19
      Top             =   4440
      Width           =   1995
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   75
      ScaleHeight     =   1665
      ScaleWidth      =   7335
      TabIndex        =   15
      Top             =   90
      Width           =   7365
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   765
      MaxLength       =   20
      TabIndex        =   10
      Text            =   "THIS IS A TEST"
      Top             =   2340
      Width           =   2265
   End
   Begin VB.TextBox txtY 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3900
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "1"
      Top             =   3750
      Width           =   315
   End
   Begin VB.TextBox txtX 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3900
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "1"
      Top             =   3150
      Width           =   315
   End
   Begin VB.TextBox txtD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3915
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "1"
      Top             =   2490
      Width           =   315
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BottomLeft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   420
      TabIndex        =   3
      Top             =   3405
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TopRight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2025
      TabIndex        =   2
      Top             =   2760
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BottomRight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2025
      TabIndex        =   1
      Top             =   3405
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TopLeft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   420
      TabIndex        =   0
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   705
      TabIndex        =   22
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Background"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5145
      TabIndex        =   18
      Top             =   3255
      Width           =   1215
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5055
      TabIndex        =   17
      Top             =   3510
      Width           =   1305
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Not all color combinations work well."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   16
      Top             =   1860
      Width           =   3915
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "End Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5250
      TabIndex        =   14
      Top             =   2550
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5190
      TabIndex        =   13
      Top             =   1860
      Width           =   1005
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5055
      TabIndex        =   12
      Top             =   2115
      Width           =   1305
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5055
      TabIndex        =   11
      Top             =   2805
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Y offset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3810
      TabIndex        =   9
      Top             =   3555
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X offset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3810
      TabIndex        =   7
      Top             =   2940
      Width           =   675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Depth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   2265
      Width           =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************
'*
'*         Name of Program: 3D Text
'*                  Author: Ken Foster
'*                 Version: 1.0.0
'*                    Date: September 05,2005
'*                    Time: 03:54 PM
'*         No Copyrights claimed - Use as you like
'*
'***************************************************
'***************** Table of Procedures *************
'   Private Sub Form_Load
'   Private Sub Command1_Click
'   Private Sub Command2_Click
'   Private Sub Command3_Click
'   Private Sub Command4_Click
'   Private Sub Command5_Click
'   Private Sub Doit
'   Private Sub GetColor
'   Private Sub GetCase
'   Private Sub Label4_Click
'   Private Sub Label5_Click
'   Private Sub Label9_Click
'   Private Sub FileExists
'   Public Sub ShowColor
'   Public Sub Fillit
'***************** End of Table ********************
 Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
   
Private Type CHOOSECOLOR
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   rgbResult As Long
   lpCustColors As String
   flags As Long
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Dim CustomColors() As Byte
Dim CC As CHOOSECOLOR

Dim R As Integer
Dim G As Integer
Dim B As Integer
Dim SRed As Integer
Dim SGreen As Integer
Dim SBlue As Integer
Dim ERed As Integer
Dim EGreen As Integer
Dim EBlue As Integer
Dim ct As Integer                                             'keeps track of last button pressed
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub Form_Load()
   Fillit                                                     'For the custom colors in color dialog
End Sub

Private Sub Command1_Click()                                  'TopLeft
   If txtD.Text = "" Or txtD.Text = "0" Then txtD.Text = "1"
   ct = 1
   Doit (ct)
End Sub

Private Sub Command2_Click()                                  'BottomRight
   If txtD.Text = "" Or txtD.Text = "0" Then txtD.Text = "1"
   ct = 2
   Doit (ct)
End Sub

Private Sub Command3_Click()                                  'TopRight
   If txtD.Text = "" Or txtD.Text = "0" Then txtD.Text = "1"
   ct = 3
   Doit (ct)
End Sub

Private Sub Command4_Click()                                  'BottomLeft
   If txtD.Text = "" Or txtD.Text = "0" Then txtD.Text = "1"
   ct = 4
   Doit (ct)
End Sub

Private Sub Command5_Click()                                  'Save
   Dim iresponse As String
   Dim Fname As String
   
   BitBlt Picture1.hDC, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, pic1.hDC, 0, 0, vbSrcCopy
   Picture1.Picture = Picture1.Image
   If Text2.Text = "" Then
      MsgBox "No filename, try again."
      Exit Sub
   End If
   
   Fname = App.Path & "\" & Text2.Text & ".bmp"
   FileExists (Fname)

   If FileExists(Fname) = True Then
          iresponse = MsgBox("File Exists!! Do you want to overwrite file?", vbYesNo, "File Exists")
          If iresponse = vbNo Then Exit Sub
   End If

   SavePicture Picture1.Picture, App.Path & "\" & Text2.Text & ".bmp"
   MsgBox "Saved in appfolder as " & Text2.Text & ".bmp"
End Sub

Private Sub Doit(Index As Integer)
   Dim i As Integer
   Dim x As Single
   Dim y As Single
   Dim RChange As Integer
   Dim GChange As Integer
   Dim BChange As Integer
   
   On Error Resume Next
   pic1.Cls
   GetColor Label5.BackColor, 0, 0, 0, Label4.BackColor, 0, 0, 0
   For i = 0 To 254 Step txtD.Text
      
      Select Case ct
         Case 1                                              'TopLeft
            x = x - txtX.Text
            y = y - txtY.Text
         Case 2                                              'BottomRight
            x = x + txtX.Text
            y = y + txtY.Text
         Case 3                                              'TopRight
            x = x + txtX.Text
            y = y - txtY.Text
         Case 4                                              'BottomLeft
            x = x - txtX.Text
            y = y + txtY.Text
      End Select
      
      pic1.CurrentX = 520 + x                                'determines where text will print (Right/Left)
      pic1.CurrentY = 520 + y                                'also used to determine where text will print (Up/Down)
      
      RChange = RChange + (ERed - SRed) / 255                'start of gradient colors
      GChange = GChange + (EGreen - SGreen) / 255
      BChange = BChange + (EBlue - SBlue) / 255
      R = SRed + RChange
      G = SGreen + GChange
      B = SBlue + BChange
      pic1.ForeColor = RGB(R, G, B)                          'set text color
      
      If ct = 2 Or ct = 3 Or ct = 4 Then
         If i >= 226 Then pic1.ForeColor = Label4.BackColor  'adds a shadow effect with minor adjustment so all match TopLeft button
         If i >= 240 Then pic1.ForeColor = Label5.BackColor  'highlights start text
      Else
         If i >= 220 Then pic1.ForeColor = Label4.BackColor  'adds a shadow effect
         If i >= 240 Then pic1.ForeColor = Label5.BackColor  'highlights start text
      End If
      
      pic1.Print Text1.Text
      Next
   End Sub

Private Sub GetColor(ByVal LngCol As Long, R1 As Integer, G1 As Integer, B1 As Integer, LngCol1 As Long, R2 As Integer, G2 As Integer, B2 As Integer)
   R1 = LngCol Mod 256
   G1 = (LngCol And vbGreen) / 256
   B1 = (LngCol And vbBlue) / 65536
   
   R2 = LngCol1 Mod 256
   G2 = (LngCol1 And vbGreen) / 256
   B2 = (LngCol1 And vbBlue) / 65536
   
   SRed = R2
   SGreen = G2
   SBlue = B2
   ERed = R1
   EGreen = G1
   EBlue = B1
End Sub

Private Sub GetCase()
   Dim x As Integer
   
    For x = 1 To 4
      Select Case ct                                         'Get which button was pressed last
         Case 1
            Command1_Click                                   'TopLeft
         Case 2
            Command2_Click                                   'BottomRight
         Case 3
            Command3_Click                                   'TopRight
         Case 4
            Command4_Click                                   'BottomLeft
      End Select
    Next
End Sub

Private Sub Label4_Click()                                   'End color
   Dim Sure As Long
  
   Sure = ShowColor
   If Sure = -1 Then Exit Sub                                'Cancel was clicked
   Label4.BackColor = Sure
   GetCase                                                   'Get last button pressed
   End Sub

Private Sub Label5_Click()                                   'Start color
   Dim Sure As Long
   
   Sure = ShowColor
   If Sure = -1 Then Exit Sub                                'Cancel was clicked
   Label5.BackColor = Sure
   GetCase                                                   'Get last button pressed
   End Sub

Private Sub Label9_Click()                                   'picture background color
   Dim Sure As Long
  
   Sure = ShowColor
   If Sure = -1 Then Exit Sub                                'Cancel was clicked
   Label9.BackColor = Sure
   pic1.BackColor = Label9.BackColor
   GetCase                                                   'Get last button pressed
   End Sub
   
Public Function FileExists(Fname As String) As Boolean

If Fname = "" Or Right(Fname, 1) = "\" Then
  FileExists = False: Exit Function
End If

FileExists = (Dir(Fname) <> "")
End Function

 Public Function ShowColor() As Long
   
   CC.lStructSize = Len(CC)                                  'set the structure size
   CC.hwndOwner = Form1.hWnd                                 'Set the owner
   CC.hInstance = App.hInstance                              'set the application's instance
   CC.lpCustColors = StrConv(CustomColors, vbUnicode)        'set the custom colors (converted to Unicode)
                                                             'no extra flags
   CC.flags = 0                                              'set to 0 = define custom colors unselected.
                                                             '2= define custom colors selected
   If CHOOSECOLOR(CC) <> 0 Then                              'Show the 'Select Color'-dialog
      ShowColor = (CC.rgbResult)
      CustomColors = StrConv(CC.lpCustColors, vbFromUnicode)
   Else
      ShowColor = -1
   End If
End Function

Public Sub Fillit()
   Dim i As Integer
   
   ReDim CustomColors(0 To 16 * 4 - 1) As Byte
   
   For i = LBound(CustomColors) To UBound(CustomColors)
      CustomColors(i) = 0
   Next i
End Sub

Private Sub txtD_KeyPress(KeyAscii As Integer)
      Const Numbers$ = "123456789"
    If KeyAscii <> 8 Then
        If InStr(Numbers, Chr(KeyAscii)) = 0 Then
            MsgBox "error"
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
