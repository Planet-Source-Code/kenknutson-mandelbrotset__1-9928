VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000080&
   Caption         =   "Making a MandelBrotSet"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtColorStart 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      TabIndex        =   15
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtColorIncrement 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtCPY2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Text            =   "-2"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtCPX2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      TabIndex        =   9
      Text            =   "1"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtCPY1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      TabIndex        =   7
      Text            =   "2"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtCPX1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      TabIndex        =   5
      Text            =   "-2"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtColorStep 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtColorMax 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5400
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5100
      Left            =   120
      ScaleHeight     =   5040
      ScaleWidth      =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   5100
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Color Start(0+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Color Increment(1+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CP Y2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CP X2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CP Y1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CP X1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Color Step (1+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Color Max (0+)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuContextCopyPic2Clipboard 
         Caption         =   "&Copy Picture to Clipboard"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Plotting As Boolean

Dim Resizer As clsResize
    
Private Sub Form_Load()
    Set Resizer = New clsResize
    Resizer.SourceForm = Me
'    Resizer.SizeFormToScreen 75
End Sub

Private Sub Form_Resize()
  Resizer.ResizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'check if form is minimized or maximized
    '     and set it back to normal
    
    Set Resizer = Nothing
    
    Plotting = False
    
    If Me.WindowState > 0 Then
        Me.WindowState = 0
    End If
    a = 400
    B = 400


    For i = 1 To 30
        'A is X2
        'B is Y2
        a = a - 10
        B = B - 10
        'set shape
        SetWindowRgn hWnd, CreateEllipticRgn(0, 0, a, B), True
        Me.Refresh
    Next


    For j = 1 To 10
        Me.Left = Me.Left + 800
        Me.Top = Me.Top + 800
    Next
    End
End Sub

Private Sub Form_Click()
  'DrawMandelBrautSet Me
End Sub

Private Sub mnuContextCopyPic2Clipboard_Click()

   Clipboard.Clear
   If TypeOf Screen.ActiveControl Is TextBox Then
      Clipboard.SetText Screen.ActiveControl.SelText
   ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
      Clipboard.SetText Screen.ActiveControl.Text
   ElseIf TypeOf Screen.ActiveControl Is PictureBox _
         Then
      Clipboard.SetData Screen.ActiveControl.Image
   ElseIf TypeOf Screen.ActiveControl Is ListBox Then
      Clipboard.SetText Screen.ActiveControl.Text
   Else
      ' No action makes sense for the other controls.
   End If

  
End Sub

Private Sub Picture1_Click()
  Dim colormax As Long
  Dim colorstep As Long
  Dim cpx1 As Currency
  Dim cpy1 As Currency
  Dim cpx2 As Currency
  Dim cpy2 As Currency
  Dim colorinc As Integer
  Dim ColorStart As Long
  
  ColorStart = CLng(Val(txtColorStart.Text))
  colorinc = Int(Val(txtColorIncrement.Text))
  colorstep = CLng(Val(txtColorStep.Text))
  colormax = CLng(Val(txtColorMax.Text))
  cpx1 = Val(txtCPX1.Text)
  cpy1 = Val(txtCPY1.Text)
  cpx2 = Val(txtCPX2.Text)
  cpy2 = Val(txtCPY2.Text)
  DrawMandelBrautSet Picture1, colormax, colorstep, cpx1, cpy1, cpx2, cpy2, colorinc, ColorStart
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMenu mnuContext
  End If
End Sub
