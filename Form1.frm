VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Collision Detector/animation/dialog"
   ClientHeight    =   4140
   ClientLeft      =   2040
   ClientTop       =   2010
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6450
   Begin VB.Timer Rundown 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   2820
      Top             =   3180
   End
   Begin VB.Timer RunUp 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   4080
      Top             =   3180
   End
   Begin VB.Timer Runleft 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   3660
      Top             =   3180
   End
   Begin VB.Timer Runright 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   3240
      Top             =   3180
   End
   Begin VB.Label DialogChoice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5415
      TabIndex        =   31
      Top             =   3645
      Width           =   615
   End
   Begin VB.Label DialogChoice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4875
      TabIndex        =   30
      Top             =   3645
      Width           =   585
   End
   Begin VB.Label Codenumber 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0/20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   29
      Top             =   2010
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Totalskillnumber 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Hacksnumber 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   27
      Top             =   1545
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Exploitsnumber 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   1770
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label CodeLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2805
      TabIndex        =   25
      Top             =   2010
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label AreaLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Warehouse Alley"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3270
      TabIndex        =   24
      Top             =   2310
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label CurrentAreaLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Area:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2790
      TabIndex        =   23
      Top             =   2325
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label ExploitsLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Exploits:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2805
      TabIndex        =   22
      Top             =   1770
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label HacksLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Hacks:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2805
      TabIndex        =   21
      Top             =   1545
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label TotalSkillLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Skill:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2805
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Charlevellabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4305
      TabIndex        =   19
      Top             =   975
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Charhealthlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4305
      TabIndex        =   18
      Top             =   585
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label MenuStatusCharLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3615
      TabIndex        =   17
      Top             =   975
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label MenuStatusCharHealth 
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3615
      TabIndex        =   16
      Top             =   585
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label MenuStatusCharName 
      BackStyle       =   0  'Transparent
      Caption         =   "Demo Char"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3615
      TabIndex        =   15
      Top             =   375
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image Character01 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   2730
      Picture         =   "Form1.frx":0000
      Top             =   375
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label MenuStatusClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4755
      TabIndex        =   14
      Top             =   150
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image DialogBox 
      Height          =   2445
      Index           =   27
      Left            =   2625
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Image DialogBox 
      Height          =   1995
      Index           =   35
      Left            =   2565
      Picture         =   "Form1.frx":1F56
      Stretch         =   -1  'True
      Top             =   345
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image DialogBox 
      Height          =   2025
      Index           =   34
      Left            =   4950
      Picture         =   "Form1.frx":279C
      Stretch         =   -1  'True
      Top             =   345
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image DialogBox 
      Height          =   60
      Index           =   33
      Left            =   2910
      Picture         =   "Form1.frx":2FE2
      Stretch         =   -1  'True
      Top             =   2580
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Image DialogBox 
      Height          =   60
      Index           =   32
      Left            =   2895
      Picture         =   "Form1.frx":30FC
      Stretch         =   -1  'True
      Top             =   75
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   31
      Left            =   2565
      Picture         =   "Form1.frx":3216
      Top             =   2325
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   30
      Left            =   2565
      Picture         =   "Form1.frx":3934
      Top             =   45
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   29
      Left            =   4665
      Picture         =   "Form1.frx":4052
      Top             =   45
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   28
      Left            =   4665
      Picture         =   "Form1.frx":4770
      Top             =   2325
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label MenuLoadGame 
      BackStyle       =   0  'Transparent
      Caption         =   "Load Game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   1020
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label MenuStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   780
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label MenuSounds 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4440
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label MenuSoundsLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Sounds:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3120
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label MenuSaveGame 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   1260
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label MenuItems 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label MenuConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Config"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   1500
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label MenuCloseMenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Close Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   1740
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image DialogBox 
      Height          =   1935
      Index           =   13
      Left            =   5175
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   14
      Left            =   5115
      Picture         =   "Form1.frx":4E8E
      Top             =   1785
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label MenuConfigSpeedLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Run Speed:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3105
      TabIndex        =   8
      Top             =   330
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label MenuConfigSpeed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3030
      TabIndex        =   7
      Top             =   570
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Image MenuConfigSpeedDown 
      Height          =   75
      Left            =   4230
      Picture         =   "Form1.frx":55AC
      Top             =   705
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image MenuConfigSpeedUp 
      Height          =   75
      Left            =   4230
      Picture         =   "Form1.frx":565A
      Top             =   585
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label MenuConfigClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   4755
      TabIndex        =   6
      Top             =   150
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image DialogBox 
      Height          =   1200
      Index           =   18
      Left            =   2970
      Stretch         =   -1  'True
      Top             =   135
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   26
      Left            =   4665
      Picture         =   "Form1.frx":5708
      Top             =   1080
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   25
      Left            =   4665
      Picture         =   "Form1.frx":5E26
      Top             =   45
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   24
      Left            =   2910
      Picture         =   "Form1.frx":6544
      Top             =   45
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   23
      Left            =   2910
      Picture         =   "Form1.frx":6C62
      Top             =   1080
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   60
      Index           =   22
      Left            =   3195
      Picture         =   "Form1.frx":7380
      Stretch         =   -1  'True
      Top             =   75
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image DialogBox 
      Height          =   60
      Index           =   21
      Left            =   3195
      Picture         =   "Form1.frx":749A
      Stretch         =   -1  'True
      Top             =   1335
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Image DialogBox 
      Height          =   960
      Index           =   20
      Left            =   4950
      Picture         =   "Form1.frx":75B4
      Stretch         =   -1  'True
      Top             =   345
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image DialogBox 
      Height          =   765
      Index           =   19
      Left            =   2910
      Picture         =   "Form1.frx":7DFA
      Stretch         =   -1  'True
      Top             =   345
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image DialogBox 
      Height          =   1500
      Index           =   12
      Left            =   5115
      Picture         =   "Form1.frx":8640
      Stretch         =   -1  'True
      Top             =   330
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image DialogBox 
      Height          =   1530
      Index           =   11
      Left            =   6360
      Picture         =   "Form1.frx":8E86
      Stretch         =   -1  'True
      Top             =   330
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image DialogBox 
      Height          =   60
      Index           =   10
      Left            =   5460
      Picture         =   "Form1.frx":96CC
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image DialogBox 
      Height          =   60
      Index           =   9
      Left            =   5460
      Picture         =   "Form1.frx":97E6
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label dialog1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use Numpad Keys, click to close this box"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   4905
      TabIndex        =   2
      Top             =   3000
      Width           =   1050
   End
   Begin VB.Image DialogBox 
      Height          =   60
      Index           =   8
      Left            =   5115
      Picture         =   "Form1.frx":9900
      Stretch         =   -1  'True
      Top             =   2835
      Width           =   645
   End
   Begin VB.Image DialogBox 
      Height          =   60
      Index           =   7
      Left            =   5160
      Picture         =   "Form1.frx":9A1A
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   660
   End
   Begin VB.Image DialogBox 
      Height          =   405
      Index           =   6
      Left            =   6060
      Picture         =   "Form1.frx":9B34
      Stretch         =   -1  'True
      Top             =   3165
      Width           =   60
   End
   Begin VB.Image DialogBox 
      Height          =   405
      Index           =   5
      Left            =   4800
      Picture         =   "Form1.frx":A37A
      Stretch         =   -1  'True
      Top             =   3195
      Width           =   60
   End
   Begin VB.Image DialogBox 
      Height          =   825
      Index           =   4
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   2925
      Width           =   1185
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   3
      Left            =   4815
      Picture         =   "Form1.frx":ABC0
      Top             =   3585
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   2
      Left            =   5790
      Picture         =   "Form1.frx":B2DE
      Top             =   3555
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   1
      Left            =   5760
      Picture         =   "Form1.frx":B9FC
      Top             =   2820
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   0
      Left            =   4800
      Picture         =   "Form1.frx":C11A
      Top             =   2820
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Impact area 0001 , Figure cannot move within this box"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   4140
      Width           =   6615
   End
   Begin VB.Image Stillright 
      Height          =   480
      Left            =   1110
      Picture         =   "Form1.frx":C838
      Top             =   435
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Stillleft 
      Height          =   480
      Left            =   840
      Picture         =   "Form1.frx":D102
      Top             =   420
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Stillback 
      Height          =   480
      Left            =   525
      Picture         =   "Form1.frx":D9CC
      Top             =   420
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Stillforward 
      Height          =   480
      Left            =   210
      Picture         =   "Form1.frx":E296
      Top             =   405
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image MainFigure 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   2190
      Picture         =   "Form1.frx":EB60
      Top             =   2835
      Width           =   480
   End
   Begin VB.Image label3 
      Height          =   630
      Left            =   3930
      Picture         =   "Form1.frx":F42A
      Top             =   2535
      Width           =   630
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   17
      Left            =   5115
      Picture         =   "Form1.frx":10AC4
      Top             =   30
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   16
      Left            =   6075
      Picture         =   "Form1.frx":111E2
      Top             =   30
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image DialogBox 
      Height          =   345
      Index           =   15
      Left            =   6075
      Picture         =   "Form1.frx":11900
      Top             =   1785
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   -15
      Picture         =   "Form1.frx":1201E
      Top             =   -345
      Width           =   6465
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Impact area 0002 , Figure cannot move within this box"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6615
   End
   Begin VB.Image Puddle 
      Height          =   1140
      Left            =   630
      Picture         =   "Form1.frx":1B2EF
      Top             =   2385
      Width           =   1305
   End
   Begin VB.Image Image2 
      Height          =   3225
      Left            =   -120
      Picture         =   "Form1.frx":1BAA0
      Top             =   1725
      Width           =   9570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub dialog1_Click()
On Error Resume Next
'if the player clicks the dialog box, close it
dialog1.Visible = False
Dim errr As Integer
For errr = 0 To 8
DialogBox(errr).Visible = False
Next
DialogChoice(1).Visible = False
DialogChoice(0).Visible = False
End Sub

Private Sub DialogBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MenuItems.ForeColor = vbWhite
MenuCloseMenu.ForeColor = vbWhite
MenuConfig.ForeColor = vbWhite
MenuSaveGame.ForeColor = vbWhite
MenuStatus.ForeColor = vbWhite
End Sub

Private Sub DialogChoice_Click(Index As Integer)
Select Case Index
Case 0
ChoiceReturn = 0
If Gameexit = True Then
Unload Me
End
Else
End If
Case 1
ChoiceReturn = 1
If Gameexit = True Then
dialog1_Click
Else
End If
Gameexit = False
End Select
End Sub

Private Sub DialogChoice_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
DialogChoice(Index).ForeColor = &HC0FFC0
If Index = 1 Then
DialogChoice(0).ForeColor = vbWhite
ElseIf Index = 0 Then
DialogChoice(1).ForeColor = vbWhite
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ab:
Dim tmm As Integer
Select Case KeyCode

Case 107
For erer = 9 To 17
DialogBox(erer).Visible = True
Next erer
MenuCloseMenu.Visible = True
MenuItems.Visible = True
MenuConfig.Visible = True
MenuSaveGame.Visible = True
MenuLoadGame.Visible = True
MenuStatus.Visible = True
Case 27 'esc key
Gameexit = True
Arrange_dialog (Me.ScaleWidth / 2) - 1100, (Me.ScaleHeight / 2) - 500, 2000, 500, "Are you sure you want to exit?", True, "Yes", "No!"



Case vbKeyRight '(numpad right)
Runright.Enabled = True
RunUp.Enabled = False
Rundown.Enabled = False
Runleft.Enabled = False

Case vbKeyLeft '(numpad left)
Runleft.Enabled = True
RunUp.Enabled = False
Rundown.Enabled = False
Runright.Enabled = False

Case vbKeyUp '(numpad up)
RunUp.Enabled = True
Runright.Enabled = False
Runleft.Enabled = False
Rundown.Enabled = False

Case vbKeyDown '(numpad down)
Runright.Enabled = False
Runleft.Enabled = False
RunUp.Enabled = False
Rundown.Enabled = True


Case 96 '(numpad-0)(Close dialog box key)
On Error Resume Next
dialog1.Visible = False
DialogChoice(1).Visible = False
DialogChoice(0).Visible = False
Dim errr As Integer
For errr = 0 To 8
DialogBox(errr).Visible = False
Next
End Select

Exit Sub
ab:
Select Case Err.Number
Case 23
Exit Sub
e100e = 0
e104e = 0
e98e = 0
e102e = 0
RunUp.Enabled = False
Rundown.Enabled = False
Runright.Enabled = False
Runleft.Enabled = False
WalkOnWater = False
Sndd1 = ""
Sndd2 = ""
Case Else
MsgBox Err.Number & " " & Err.Description & "   This unexpected error occured!, please restart the program"
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyRight '(numpad right)
Runright.Enabled = False
MainFigure.Picture = Stillright.Picture
Case vbKeyLeft '(numpad left)
Runleft.Enabled = False
MainFigure.Picture = Stillleft.Picture
Case vbKeyUp '(numpad up)
RunUp.Enabled = False
MainFigure.Picture = Stillback.Picture
Case vbKeyDown '(numpad down)
Rundown.Enabled = False
MainFigure.Picture = Stillforward.Picture
End Select
End Sub

Private Sub Form_Load()
DWn = 1
P = 1
E1 = 1
I = 1
AA = 1
DWn1 = 1
speed = 40
Gameexit = False
DialogBox(27).Picture = LoadPicture(App.Path & "\data\images\menu\menugradient.bmx")
DialogBox(13).Picture = LoadPicture(App.Path & "\data\images\menu\menugradient.bmx")
DialogBox(4).Picture = LoadPicture(App.Path & "\data\images\menu\menugradient.bmx")
DialogBox(18).Picture = LoadPicture(App.Path & "\data\images\menu\menugradient.bmx")

MenuSoundsEnabled = True
Arrange_dialog Rnd * 2000, Rnd * 800, 3000, 1200, "Welcome the the new version! Sound FX Added in new version! Code optimized! NEW CONTROLS: Arrow keys now control movement, for more classic gameplay. Numpad + for menu, Numpad 0 to close any open dialog boxes (like this one!). Esc to exit.", False, "", ""
ChoiceReturn = 3
End Sub

Private Sub Image1_Click()
Arrange_dialog (Rnd * 1000), (Rnd * 1000), 1000, 400, "Back of Warehouse Building", False, "", ""
End Sub



Private Sub Image4_Click()

End Sub


Private Sub label3_Click()

Arrange_dialog (Rnd * 1000), (Rnd * 1000), 1000, 300, "Some Barrels", False, "", ""
End Sub


Private Sub Label6_Click()

End Sub

Private Sub MainFigure_Click()
Arrange_dialog 1000, 300, 900, 415, "Yes, that is you! kinda small huh?", False, "", ""
End Sub

Private Sub MenuCloseMenu_Click()
On Error Resume Next
Dim erer As Integer
For erer = 9 To DialogBox.Count
DialogBox(erer).Visible = False
Next
MenuCloseMenu.Visible = False
MenuItems.Visible = False
MenuConfig.Visible = False
MenuConfigClose.Visible = False
MenuConfigSpeedUp.Visible = False
MenuConfigSpeedDown.Visible = False
MenuConfigSpeed.Visible = False
MenuConfigSpeedLabel.Visible = False
MenuSaveGame.Visible = False
MenuSoundsLabel.Visible = False
MenuSounds.Visible = False
MenuLoadGame.Visible = False
MenuStatus.Visible = False
MenuStatusClose.Visible = False
Character01.Visible = False
MenuStatusCharName.Visible = False
MenuStatusCharHealth.Visible = False
MenuStatusCharLevel.Visible = False
Charlevellabel.Visible = False
Charhealthlabel.Visible = False

TotalSkillLabel.Visible = False
HacksLabel.Visible = False
ExploitsLabel.Visible = False
CodeLabel.Visible = False
CurrentAreaLabel.Visible = False
AreaLabel.Visible = False

Totalskillnumber.Visible = False
Hacksnumber.Visible = False
Exploitsnumber.Visible = False
Codenumber.Visible = False
End Sub

Private Sub MenuCloseMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MenuSoundsEnabled = True Then
If MenuCloseMenu.ForeColor = vbWhite Then
PlayMidi (App.Path & "\data\sfx\" & "menu.wav")
Else
End If
Else
End If
MenuItems.ForeColor = vbWhite
MenuCloseMenu.ForeColor = &HC0FFC0
MenuConfig.ForeColor = vbWhite
MenuSaveGame.ForeColor = vbWhite
MenuStatus.ForeColor = vbWhite
End Sub


Private Sub MenuConfig_Click()
On Error Resume Next
Dim erer As Integer
For erer = 18 To 26
DialogBox(erer).Visible = True
Next
MenuConfigClose.Visible = True
MenuConfigSpeedUp.Visible = True
MenuConfigSpeedDown.Visible = True
MenuConfigSpeed.Visible = True
MenuConfigSpeedLabel.Visible = True
MenuSoundsLabel.Visible = True
MenuSounds.Visible = True
For erer = 27 To 35
DialogBox(erer).Visible = False
Next
MenuStatusClose.Visible = False
Character01.Visible = False
MenuStatusCharName.Visible = False
MenuStatusCharHealth.Visible = False
MenuStatusCharLevel.Visible = False
Charlevellabel.Visible = False
Charhealthlabel.Visible = False

TotalSkillLabel.Visible = False
HacksLabel.Visible = False
ExploitsLabel.Visible = False
CodeLabel.Visible = False
CurrentAreaLabel.Visible = False
AreaLabel.Visible = False

Totalskillnumber.Visible = False
Hacksnumber.Visible = False
Exploitsnumber.Visible = False
Codenumber.Visible = False
End Sub

Private Sub MenuConfig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MenuSoundsEnabled = True Then
If MenuConfig.ForeColor = vbWhite Then
PlayMidi (App.Path & "\data\sfx\" & "menu.wav")
Else
End If
Else
End If
MenuItems.ForeColor = vbWhite
MenuCloseMenu.ForeColor = vbWhite
MenuConfig.ForeColor = &HC0FFC0
MenuSaveGame.ForeColor = vbWhite
MenuStatus.ForeColor = vbWhite
End Sub


Private Sub MenuConfigClose_Click()
On Error Resume Next

For erer = 26 To 18 Step -1
DialogBox(erer).Visible = False
Next
MenuConfigClose.Visible = False
MenuConfigSpeedUp.Visible = False
MenuConfigSpeedDown.Visible = False
MenuConfigSpeed.Visible = False
MenuConfigSpeedLabel.Visible = False
MenuSoundsLabel.Visible = False
MenuSounds.Visible = False
End Sub

Private Sub MenuConfigSpeed_Change()
If MenuConfigSpeed.Caption >= 65 Then
MenuConfigSpeed.Caption = 64
Else
End If

If MenuConfigSpeed.Caption <= 10 Then
MenuConfigSpeed.Caption = 11
Else
End If

speed = MenuConfigSpeed.Caption
End Sub

Private Sub MenuConfigSpeedDown_Click()
MenuConfigSpeed = MenuConfigSpeed - 1
End Sub

Private Sub MenuConfigSpeedUp_Click()
MenuConfigSpeed = MenuConfigSpeed + 1
End Sub

Private Sub MenuItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

MenuConfig.ForeColor = vbWhite
MenuCloseMenu.ForeColor = vbWhite
MenuSaveGame.ForeColor = vbWhite
If MenuSoundsEnabled = True Then
If MenuItems.ForeColor = vbWhite Then
PlayMidi (App.Path & "\data\sfx\" & "menu.wav")
Else
End If
Else
End If
MenuItems.ForeColor = &HC0FFC0
End Sub

Private Sub MenuSaveGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MenuSoundsEnabled = True Then
If MenuSaveGame.ForeColor = vbWhite Then
PlayMidi (App.Path & "\data\sfx\" & "menu.wav")
Else
End If
Else
End If
MenuItems.ForeColor = vbWhite
MenuCloseMenu.ForeColor = vbWhite
MenuConfig.ForeColor = vbWhite
MenuSaveGame.ForeColor = &HC0FFC0
End Sub

Private Sub MenuSounds_Click()
Select Case LCase(MenuSounds.Caption)
Case "yes"
MenuSoundsEnabled = False
MenuSounds.Caption = "No"
Case "no"
MenuSoundsEnabled = True
MenuSounds.Caption = "Yes"
End Select
End Sub

Private Sub MenuStatus_Click()
On Error Resume Next
Dim erer As Integer
For erer = 26 To 18 Step -1
DialogBox(erer).Visible = False
Next
MenuConfigClose.Visible = False
MenuConfigSpeedUp.Visible = False
MenuConfigSpeedDown.Visible = False
MenuConfigSpeed.Visible = False
MenuConfigSpeedLabel.Visible = False
MenuSoundsLabel.Visible = False
MenuSounds.Visible = False
erer = 0
For erer = 27 To 35
DialogBox(erer).Visible = True
Next
MenuStatusClose.Visible = True
Character01.Visible = True
MenuStatusCharName.Visible = True
MenuStatusCharHealth.Visible = True
MenuStatusCharLevel.Visible = True
Charlevellabel.Visible = True
Charhealthlabel.Visible = True


TotalSkillLabel.Visible = True
HacksLabel.Visible = True
ExploitsLabel.Visible = True
CodeLabel.Visible = True
CurrentAreaLabel.Visible = True
AreaLabel.Visible = True

Totalskillnumber.Visible = True
Hacksnumber.Visible = True
Exploitsnumber.Visible = True
Codenumber.Visible = True

End Sub

Private Sub MenuStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MenuSoundsEnabled = True Then
If MenuStatus.ForeColor = vbWhite Then
PlayMidi (App.Path & "\data\sfx\" & "menu.wav")
Else
End If
Else
End If
MenuItems.ForeColor = vbWhite
MenuStatus.ForeColor = &HC0FFC0
MenuCloseMenu.ForeColor = vbWhite
MenuConfig.ForeColor = vbWhite
MenuSaveGame.ForeColor = vbWhite
End Sub

Private Sub MenuStatusClose_Click()
MenuStatusClose.Visible = False
Character01.Visible = False
MenuStatusCharName.Visible = False
MenuStatusCharHealth.Visible = False
MenuStatusCharLevel.Visible = False
Charlevellabel.Visible = False
Charhealthlabel.Visible = False

TotalSkillLabel.Visible = False
HacksLabel.Visible = False
ExploitsLabel.Visible = False
CodeLabel.Visible = False
CurrentAreaLabel.Visible = False
AreaLabel.Visible = False

Totalskillnumber.Visible = False
Hacksnumber.Visible = False
Exploitsnumber.Visible = False
Codenumber.Visible = False

Dim erer As Integer
For erer = 27 To 35
DialogBox(erer).Visible = False
Next
End Sub

Private Sub Puddle_Click()
Arrange_dialog 340, 600, 900, 400, "Looks like a puddle", False, "", ""
End Sub

Private Sub Rundown_Timer()

' this is the code for running DOWN (south)
'there are 15 frames in the animation, so
'if the frame gets to 16, aa(frame) will go
'back to one if the frame is less than 16,
'the frame is increased by one until it = 15

If AA <> 16 Then
MainFigure.Picture = LoadPicture(App.Path & "\data\images\characters\char01\rundown\run" & AA & ".ico")
MainFigure.Move MainFigure.Left, MainFigure.Top + speed


If WalkOnWater = True Then
Sndd1 = "footsteps(water).wav"
Sndd2 = "footsteps(water1).wav"
Else
Sndd1 = "footstep.wav"
Sndd2 = "footstep1.wav"
End If

If AA = 7 Then
PlayMidi (App.Path & "\data\sfx\" & Sndd1)
ElseIf AA = 15 Then
PlayMidi (App.Path & "\data\sfx\" & Sndd2)
End If
'collision Detection for running down
Call TestForCollision("down")

AA = AA + 1
Else
AA = 1
End If
End Sub

Private Sub Runleft_Timer()

' this is the code for running LEFT (west)
'there are 15 frames in the animation, so
'if the frame gets to 16, aa(frame) will go
'back to one if the frame is less than 16,
'the frame is increased by one until it = 15

If E1 <> 16 Then
MainFigure.Picture = LoadPicture(App.Path & "\data\images\characters\char01\runleft\run" & E1 & ".ico")
MainFigure.Move MainFigure.Left - speed

'test for collision
Call TestForCollision("left")


If WalkOnWater = True Then
Sndd1 = "footsteps(water).wav"
Sndd2 = "footsteps(water1).wav"
Else
Sndd1 = "footstep.wav"
Sndd2 = "footstep1.wav"
End If


If E1 = 7 Then
PlayMidi (App.Path & "\data\sfx\" & Sndd1)
ElseIf E1 = 15 Then
PlayMidi (App.Path & "\data\sfx\" & Sndd2)
End If

E1 = E1 + 1
Else
E1 = 1
End If
End Sub

Private Sub Runright_Timer()

'this is the code for running RIGHT (East)
'there are 15 frames in the animation, so
'if the frame gets to 16, aa(frame) will go
'back to one if the frame is less than 16,
'the frame is increased by one until it = 15

If I <> 16 Then
MainFigure.Picture = LoadPicture(App.Path & "\data\images\characters\char01\runright\run" & I & ".ico")
MainFigure.Move MainFigure.Left + speed

'test for a collision
Call TestForCollision("right")


If WalkOnWater = True Then
Sndd1 = "footsteps(water).wav"
Sndd2 = "footsteps(water1).wav"
Else
Sndd1 = "footstep.wav"
Sndd2 = "footstep1.wav"
End If

If I = 7 Then
PlayMidi (App.Path & "\data\sfx\" & Sndd1)
ElseIf I = 15 Then
PlayMidi (App.Path & "\data\sfx\" & Sndd2)
End If

I = I + 1
Else
I = 1
End If
End Sub

Private Sub RunUp_Timer()

' this is the code for running UP (North)
'there are 15 frames in the animation, so
'if the frame gets to 16, aa(frame) will go
'back to one if the frame is less than 16,
'the frame is increased by one until it = 15

If P <> 16 Then

MainFigure.Picture = LoadPicture(App.Path & "\data\images\characters\char01\runup\run" & P & ".ico")
MainFigure.Move MainFigure.Left, MainFigure.Top - speed

'collision detection for running UP
Call TestForCollision("up")


If WalkOnWater = True Then
Sndd1 = "footsteps(water).wav"
Sndd2 = "footsteps(water1).wav"
Else
Sndd1 = "footstep.wav"
Sndd2 = "footstep1.wav"
End If


If P = 7 Then
PlayMidi (App.Path & "\data\sfx\" & Sndd1)
ElseIf P = 15 Then
PlayMidi (App.Path & "\data\sfx\" & Sndd2)
End If


P = P + 1
Else
P = 1
End If
End Sub

Private Sub Timer1_Timer()

End Sub
