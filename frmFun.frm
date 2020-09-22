VERSION 5.00
Begin VB.Form frmFun 
   BackColor       =   &H00000000&
   Caption         =   "KnightRider Fun"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLights 
      Caption         =   "Lights On"
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   225
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   360
      Left            =   5040
      TabIndex        =   38
      Top             =   3075
      Width           =   1110
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   0
      Left            =   3555
      TabIndex        =   19
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   255
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin VB.PictureBox picCruiser 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2370
      Left            =   255
      Picture         =   "frmFun.frx":0000
      ScaleHeight     =   2370
      ScaleWidth      =   3045
      TabIndex        =   1
      Top             =   135
      Width           =   3045
      Begin pKnight.KnightRider KnightRider9 
         Height          =   240
         Left            =   285
         TabIndex        =   46
         Top             =   1440
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   423
         ForeColor       =   16777215
         Speed           =   20
      End
      Begin pKnight.KnightRider KnightRider5 
         Height          =   255
         Left            =   630
         TabIndex        =   42
         Top             =   1440
         Visible         =   0   'False
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   450
         ForeColor       =   8438015
         Speed           =   15
         Tail            =   0
      End
      Begin pKnight.KnightRider KnightRider3 
         Height          =   120
         Left            =   1545
         TabIndex        =   40
         Top             =   1515
         Visible         =   0   'False
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   212
         ForeColor       =   255
         Speed           =   18
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider14 
         Height          =   135
         Left            =   870
         TabIndex        =   2
         Top             =   135
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   238
         ForeColor       =   255
         Speed           =   11
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider15 
         Height          =   135
         Left            =   1815
         TabIndex        =   3
         Top             =   135
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   238
         ForeColor       =   16711680
         Speed           =   11
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider4 
         Height          =   120
         Left            =   1275
         TabIndex        =   41
         Top             =   1515
         Visible         =   0   'False
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   212
         ForeColor       =   16711680
         Speed           =   18
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider6 
         Height          =   255
         Left            =   2340
         TabIndex        =   43
         Top             =   1440
         Visible         =   0   'False
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   450
         ForeColor       =   8438015
         Speed           =   15
         Tail            =   0
      End
      Begin pKnight.KnightRider KnightRider7 
         Height          =   255
         Left            =   2835
         TabIndex        =   44
         Top             =   1425
         Visible         =   0   'False
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   450
         ForeColor       =   8438015
         Speed           =   15
         Tail            =   0
      End
      Begin pKnight.KnightRider KnightRider8 
         Height          =   255
         Left            =   165
         TabIndex        =   45
         Top             =   1425
         Visible         =   0   'False
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   450
         ForeColor       =   8438015
         Speed           =   15
         Tail            =   0
      End
      Begin pKnight.KnightRider KnightRider10 
         Height          =   240
         Left            =   2445
         TabIndex        =   47
         Top             =   1440
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   423
         ForeColor       =   16777215
         Speed           =   20
      End
   End
   Begin VB.PictureBox picLights 
      BorderStyle     =   0  'None
      Height          =   105
      Left            =   45
      Picture         =   "frmFun.frx":2C95
      ScaleHeight     =   105
      ScaleWidth      =   2040
      TabIndex        =   0
      Top             =   2775
      Width           =   2040
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   2
         Left            =   390
         TabIndex        =   6
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   3
         Left            =   540
         TabIndex        =   7
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   4
         Left            =   690
         TabIndex        =   8
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   5
         Left            =   840
         TabIndex        =   9
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   6
         Left            =   990
         TabIndex        =   10
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   7
         Left            =   1140
         TabIndex        =   11
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   8
         Left            =   1290
         TabIndex        =   12
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   9
         Left            =   1440
         TabIndex        =   13
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   10
         Left            =   1590
         TabIndex        =   14
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   11
         Left            =   1740
         TabIndex        =   15
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
      Begin pKnight.KnightRider KnightRider1 
         Height          =   75
         Index           =   12
         Left            =   1890
         TabIndex        =   16
         Top             =   15
         Width           =   120
         _ExtentX        =   212
         _ExtentY        =   132
         Enabled         =   -1  'True
         Effect          =   0
         Tail            =   1
      End
   End
   Begin pKnight.KnightRider KnightRider1 
      Height          =   75
      Index           =   13
      Left            =   885
      TabIndex        =   17
      Top             =   2790
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   132
      Enabled         =   -1  'True
      Speed           =   15
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   1
      Left            =   3705
      TabIndex        =   20
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   16777215
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   2
      Left            =   3855
      TabIndex        =   21
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   255
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   3
      Left            =   4005
      TabIndex        =   22
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   16777215
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   4
      Left            =   4155
      TabIndex        =   23
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   255
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   5
      Left            =   4305
      TabIndex        =   24
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   16777215
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   6
      Left            =   4455
      TabIndex        =   25
      Top             =   2280
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   212
      ForeColor       =   255
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   7
      Left            =   4590
      TabIndex        =   26
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   16777215
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   8
      Left            =   4740
      TabIndex        =   27
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   255
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   9
      Left            =   4890
      TabIndex        =   28
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   16777215
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   10
      Left            =   5040
      TabIndex        =   29
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   255
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   11
      Left            =   5190
      TabIndex        =   30
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   16777215
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   12
      Left            =   5340
      TabIndex        =   31
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   255
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   13
      Left            =   5490
      TabIndex        =   32
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   16777215
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   14
      Left            =   5640
      TabIndex        =   33
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   255
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   15
      Left            =   5790
      TabIndex        =   34
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   16777215
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   16
      Left            =   5940
      TabIndex        =   35
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   255
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin pKnight.KnightRider KnightRider2 
      Height          =   120
      Index           =   17
      Left            =   6090
      TabIndex        =   36
      Top             =   2280
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   212
      ForeColor       =   16777215
      Enabled         =   -1  'True
      Effect          =   0
      Speed           =   3
      Tail            =   0
   End
   Begin VB.Label lblHappy 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Happy Holidays!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3840
      TabIndex        =   37
      Top             =   1905
      Width           =   2580
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make sure you are not looking at this in your rear view mirror this Christmas!!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1140
      Left            =   3825
      TabIndex        =   18
      Top             =   720
      Width           =   2445
   End
End
Attribute VB_Name = "frmFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkLights_Click()
     If chkLights Then
          chkLights.Caption = "Lights Off" '
          KnightRider3.Visible = True
          KnightRider4.Visible = True
          KnightRider5.Visible = True
          KnightRider6.Visible = True
          KnightRider7.Visible = True
          KnightRider8.Visible = True
          KnightRider9.Visible = True
          KnightRider10.Visible = True
          KnightRider3.Enabled = True
          KnightRider4.Enabled = True
          KnightRider5.Enabled = True
          KnightRider6.Enabled = True
          KnightRider7.Enabled = True
          KnightRider8.Enabled = True
          KnightRider9.Enabled = True
          KnightRider10.Enabled = True
          KnightRider14.Enabled = True
          KnightRider15.Enabled = True
     Else
          chkLights.Caption = "Lights On"
          KnightRider3.Visible = False
          KnightRider4.Visible = False
          KnightRider5.Visible = False
          KnightRider6.Visible = False
          KnightRider7.Visible = False
          KnightRider8.Visible = False
          KnightRider9.Visible = False
          KnightRider10.Visible = False
          KnightRider3.Enabled = False
          KnightRider4.Enabled = False
          KnightRider5.Enabled = False
          KnightRider6.Enabled = False
          KnightRider7.Enabled = False
          KnightRider8.Enabled = False
          KnightRider9.Enabled = False
          KnightRider10.Enabled = False
          KnightRider14.Enabled = False
          KnightRider15.Enabled = False
     End If
End Sub

Private Sub cmdExit_Click()
     Unload Me
End Sub

Private Sub KnightRider1_TripComplete(Index As Integer)
     Dim lLoop As Long
     Dim lColor As Long
     lColor = RGB((255 * Rnd) + 1, (255 * Rnd) + 1, (255 * Rnd) + 1)
     For lLoop = 0 To 13
          KnightRider1(lLoop).ForeColor = lColor
     Next
End Sub

