VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   6795
   StartUpPosition =   3  '����ȱʡ
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdExport 
      Appearance      =   0  'Flat
      Caption         =   "����"
      Height          =   495
      Left            =   3480
      TabIndex        =   85
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtDebug 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   5775
      IMEMode         =   1  'ON
      Left            =   4920
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScrollBars      =   2  'Vertical
      TabIndex        =   87
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdBackTrackSolve 
      Appearance      =   0  'Flat
      Caption         =   "�������"
      Height          =   495
      Left            =   5160
      TabIndex        =   83
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdImport 
      Appearance      =   0  'Flat
      Caption         =   "����"
      Height          =   495
      Left            =   3480
      TabIndex        =   84
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   86
      Text            =   "Main.frx":0000
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton cmdLogicSolve 
      Appearance      =   0  'Flat
      Caption         =   "�߼����"
      Height          =   495
      Left            =   5160
      TabIndex        =   82
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   80
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   81
      Top             =   4200
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   79
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   80
      Text            =   "8"
      Top             =   4200
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   78
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   79
      Text            =   "3"
      Top             =   4200
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   77
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   78
      Text            =   "6"
      Top             =   4200
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   76
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   77
      Top             =   4200
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   75
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   76
      Top             =   4200
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   74
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   75
      Text            =   "1"
      Top             =   4200
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   73
      Left            =   600
      MaxLength       =   1
      TabIndex        =   74
      Top             =   4200
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   72
      Left            =   120
      MaxLength       =   1
      TabIndex        =   73
      Text            =   "5"
      Top             =   4200
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   71
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   72
      Top             =   3720
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   70
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   71
      Text            =   "9"
      Top             =   3720
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   69
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   70
      Top             =   3720
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   68
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   69
      Text            =   "1"
      Top             =   3720
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   67
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   68
      Top             =   3720
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   66
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   67
      Top             =   3720
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   65
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   66
      Top             =   3720
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   64
      Left            =   600
      MaxLength       =   1
      TabIndex        =   65
      Text            =   "4"
      Top             =   3720
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   63
      Left            =   120
      MaxLength       =   1
      TabIndex        =   64
      Top             =   3720
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   62
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   63
      Top             =   3240
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   61
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   62
      Top             =   3240
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   60
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   61
      Text            =   "4"
      Top             =   3240
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   59
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   60
      Top             =   3240
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   58
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   59
      Text            =   "5"
      Top             =   3240
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   57
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   58
      Text            =   "8"
      Top             =   3240
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   56
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   57
      Top             =   3240
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   55
      Left            =   600
      MaxLength       =   1
      TabIndex        =   56
      Top             =   3240
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   54
      Left            =   120
      MaxLength       =   1
      TabIndex        =   55
      Text            =   "9"
      Top             =   3240
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   53
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   54
      Text            =   "6"
      Top             =   2640
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   52
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   53
      Top             =   2640
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   51
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   52
      Top             =   2640
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   50
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   51
      Top             =   2640
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   49
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   50
      Top             =   2640
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   48
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   49
      Text            =   "7"
      Top             =   2640
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   47
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   48
      Top             =   2640
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   46
      Left            =   600
      MaxLength       =   1
      TabIndex        =   47
      Top             =   2640
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   45
      Left            =   120
      MaxLength       =   1
      TabIndex        =   46
      Text            =   "3"
      Top             =   2640
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   44
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   45
      Top             =   2160
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   43
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   44
      Text            =   "4"
      Top             =   2160
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   42
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   43
      Text            =   "5"
      Top             =   2160
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   41
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   42
      Top             =   2160
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   40
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   41
      Text            =   "9"
      Top             =   2160
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   39
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   40
      Text            =   "6"
      Top             =   2160
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   38
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   39
      Text            =   "8"
      Top             =   2160
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   37
      Left            =   600
      MaxLength       =   1
      TabIndex        =   38
      Text            =   "7"
      Top             =   2160
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   36
      Left            =   120
      MaxLength       =   1
      TabIndex        =   37
      Top             =   2160
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   35
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   36
      Text            =   "8"
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   34
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   35
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   33
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   34
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   32
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   33
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   31
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   32
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   30
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   31
      Text            =   "4"
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   29
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   30
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   28
      Left            =   600
      MaxLength       =   1
      TabIndex        =   29
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   27
      Left            =   120
      MaxLength       =   1
      TabIndex        =   28
      Text            =   "2"
      Top             =   1680
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   26
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   27
      Text            =   "1"
      Top             =   1080
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   25
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   26
      Top             =   1080
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   24
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   25
      Top             =   1080
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   23
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   24
      Text            =   "9"
      Top             =   1080
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   22
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   23
      Text            =   "6"
      Top             =   1080
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   21
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   22
      Top             =   1080
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   20
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   21
      Text            =   "3"
      Top             =   1080
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   19
      Left            =   600
      MaxLength       =   1
      TabIndex        =   20
      Top             =   1080
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   18
      Left            =   120
      MaxLength       =   1
      TabIndex        =   19
      Top             =   1080
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   17
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   18
      Top             =   600
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   16
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   17
      Text            =   "3"
      Top             =   600
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   15
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   16
      Top             =   600
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   14
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   15
      Top             =   600
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   13
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   14
      Top             =   600
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   12
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "5"
      Top             =   600
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   11
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   12
      Top             =   600
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   10
      Left            =   600
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "1"
      Top             =   600
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   9
      Left            =   120
      MaxLength       =   1
      TabIndex        =   10
      Top             =   600
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   8
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "4"
      Top             =   120
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   7
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   8
      Top             =   120
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   6
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "2"
      Top             =   120
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   5
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   6
      Top             =   120
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   4
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   5
      Top             =   120
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   3
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "1"
      Top             =   120
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   2
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "9"
      Top             =   120
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   1
      Left            =   600
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "5"
      Top             =   120
      Width           =   500
   End
   Begin VB.TextBox txtTable 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Index           =   0
      Left            =   120
      MaxLength       =   1
      TabIndex        =   1
      Top             =   120
      Width           =   500
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "���"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   5040
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' string$
' int% long&
' single! double#
' currency@

' �����߼���
' 1. �����̱���������
'    -- �������롢�ַ�������
' 2. ���
'    -- ����һ���ҽ�һ��LoadQuestion���Ѵ�֮ǰ�����������Ϊ��Ϣ��
'    -- Ȼ�����������
' 3. �ص���һ��
'    -- ��հ�ť�����밴ť



' ��¼ÿ����������ܳ��ֵ�����
Private storage(80) As possibility

' �Ƿ���load��Ŀ
Private isQuestionLoad As Boolean


Private Sub LoadQuestion()
    ' ������Ŀ����ʼ����ر���
    If isQuestionLoad = False Then
        Dim i%, j%
        For i = 0 To 80
            With txtTable(i)
                If .Text = "" Then ' û���������ݵĸ�Ҳ���������
                    .Locked = False
                    .BackColor = vbWhite
                    PossibilitySetAll storage(i), True
                Else ' �������ݵ���Ϣ��
                    .Locked = True
                    .BackColor = &HDDDDDD
                    PossibilitySetAll storage(i), False
                End If
            End With
        Next i
        
        isQuestionLoad = True
    End If
End Sub

Private Sub ClearQuestion()
    isQuestionLoad = False
    Dim i%
    For i = 0 To 80
        With txtTable(i)
            .Text = ""
            .BackColor = vbWhite
            .Locked = False
        End With
    Next i
End Sub

Private Sub BuildPossibility()
    Dim i%, j%, index%, row%, column%, tempNum%
    For index = 0 To 80
        PossibilitySetAll storage(index), True
    Next index
    For index = 0 To 80
        If txtTable(index) <> "" Then ' ����Ѿ���������
            column = index Mod 9
            row = index \ 9
            tempNum = Val(txtTable(index)) - 1
            
            PossibilitySetAll storage(index), False ' �Լ�������п�����
            
            For i = 0 To 8 ' �������е�����Ԫ�����������
                storage(row * 9 + i).p(tempNum) = False
            Next i
            
            For i = 0 To 8 ' �������е�����Ԫ�����������
                storage(i * 9 + column).p(tempNum) = False
            Next i
            
            For i = (row \ 3) * 3 To (row \ 3) * 3 + 2 ' ���ڷ����ڵ�����Ԫ�����������
                For j = (column \ 3) * 3 To (column \ 3) * 3 + 2
                    storage(i * 9 + j).p(tempNum) = False
                Next j
            Next i
            
        End If
    Next index
End Sub

Private Function BackTrack() As Boolean
    Dim index%, tempNum%
    
    For index = 0 To 80
        If txtTable(index) = "" Then
            For tempNum = 0 To 8
                If storage(index).p(tempNum) = True Then
                    txtTable(index).Text = tempNum + 1
                    BuildPossibility
                    If BackTrack() = True Then
                        '�ҵ���
                        'BackTrack = True
                        'Exit Function
                        BackTrack = True
                        DebugInfo "���н⣺"
                        DebugInfo Export
                    End If
                    txtTable(index).Text = ""
                    BuildPossibility
                End If
            Next tempNum
            '�����������ֶ����ȥ������ʧ��
            BackTrack = False
            Exit Function
        End If
    Next index
    BackTrack = True
End Function

Private Sub cmdBackTrackSolve_Click()
    '���ݷ�������
    
    LoadQuestion
    BuildPossibility
    BackTrack
End Sub

Private Sub cmdClear_Click()
    ClearQuestion
End Sub

Private Sub cmdExport_Click()
    DebugInfo Export, , True
End Sub

Private Sub cmdImport_Click()
    ClearQuestion
    Dim length&, txtpos&, txt$, tablepos%, tempchar$
    txt = txtInput.Text
    length = Len(txt)
    txtpos = 1
    tablepos = 0
    Do While txtpos <= length
        tempchar = Mid(txt, txtpos, 1)
        Select Case tempchar
        Case "1" To "9" ' ����ֱ�Ӷ���
            txtTable(tablepos) = tempchar
            tablepos = tablepos + 1
        Case "", vbCr, vbLf, vbCrLf, vbTab ' ���� Tab��������
            ' ����
        Case Else ' �����ַ���ռλ��
            txtTable(tablepos) = ""
            tablepos = tablepos + 1
        End Select
        If tablepos > 80 Then Exit Do
        txtpos = txtpos + 1
    Loop
End Sub

Private Function Export$()
    Dim index%
    For index = 0 To 80
        If txtTable(index).Text = "" Then
            Export = Export & "0"
        Else
            Export = Export & txtTable(index).Text
        End If
        If index Mod 9 = 8 Then Export = Export & vbNewLine
    Next index
End Function

Private Sub cmdLogicSolve_Click()
    ' �߼��ƶϷ���������
    
    LoadQuestion
    
    Dim i%, j%, k%
    Dim index%, row%, column%, tempNum%
    Dim actionCount%, sumPossibility%, lastPossibleNum%
    Dim numCount%(8), numLastOccur%(8)
    
    Do
        ' ���˵һ��/��/������ֻ��һ����ĳ����������һ������Ǹ���
        BuildPossibility
        actionCount = 0
        
        For i = 0 To 8 ' ������
            For k = 0 To 8 ' ��ʼ��
                numCount(k) = 0
            Next k
            For j = 0 To 8 ' ͳ��һ����ĳ�����ִ���
                For k = 0 To 8
                    If storage(i * 9 + j).p(k) = True Then
                        numCount(k) = numCount(k) + 1
                        numLastOccur(k) = i * 9 + j
                    End If
                Next k
            Next j
            For k = 0 To 8 ' ����Ƿ�ֻ��һ����ĳ����
                If numCount(k) = 1 Then
                    txtTable(numLastOccur(k)) = k + 1
                    actionCount = actionCount + 1
                End If
            Next k
        Next i
                
        For j = 0 To 8 ' ������
            For k = 0 To 8 ' ��ʼ��
                numCount(k) = 0
            Next k
            For i = 0 To 8 ' ͳ��һ����ĳ�����ִ���
                For k = 0 To 8
                    If storage(i * 9 + j).p(k) = True Then
                        numCount(k) = numCount(k) + 1
                        numLastOccur(k) = i * 9 + j
                    End If
                Next k
            Next i
            For k = 0 To 8 ' ����Ƿ�ֻ��һ����ĳ����
                If numCount(k) = 1 Then
                    txtTable(numLastOccur(k)) = k + 1
                    actionCount = actionCount + 1
                End If
            Next k
        Next j
        
        For row = 0 To 8 Step 3 ' ��������
            For column = 0 To 8 Step 3
                For k = 0 To 8 ' ��ʼ��
                    numCount(k) = 0
                Next k
                For i = row To row + 2 ' ͳ�Ʒ�����ĳ�����ִ���
                    For j = column To column + 2
                        For k = 0 To 8
                            If storage(i * 9 + j).p(k) = True Then
                                numCount(k) = numCount(k) + 1
                                numLastOccur(k) = i * 9 + j
                            End If
                        Next k
                    Next j
                Next i
                For k = 0 To 8 ' ����Ƿ�ֻ��һ����ĳ����
                    If numCount(k) = 1 Then
                        txtTable(numLastOccur(k)) = k + 1
                        actionCount = actionCount + 1
                    End If
                Next k
            Next column
        Next row
        
    Loop While actionCount <> 0 ' ��û�н�չʱ����
End Sub

Private Function possibility$(index%)
    ' ��������Ϣ
    
    Dim i%
    For i = 0 To 8
        If storage(index).p(i) = True Then
            possibility = possibility & (i + 1) & " "
        Else
            possibility = possibility & "_" & " "
        End If
    Next i
End Function

Private Sub DebugInfo(ByRef message$, Optional ByRef split$ = vbNewLine, Optional needRefresh As Boolean = False)
    ' ��txtDebug�д�ӡ������Ϣ��
    ' needRefresh�����Ƿ��ͷ��ʼ��ӡ��
    
    If needRefresh Then
        txtDebug.Text = message & split
    Else
        txtDebug.Text = txtDebug.Text & message & split
    End If
End Sub

Private Sub Form_Load()
    isQuestionLoad = False
    'EnableHighDPI Me
End Sub

Private Sub txtTable_KeyPress(index As Integer, KeyAscii As Integer)
    '   �����ַ�ʱ�Ĵ���ʽ
    ' 1. �����������1~9�����ǻ��Ԫ�����滻ԭ��Ԫ�����ݣ�Ȼ���л����㵽��һ����Ԫ��
    ' 2. �����������0�����ǻ��Ԫ����������ݣ�Ȼ���л����㵽��һ��
    ' 3. ��������˸�������ǻ��Ԫ���������ݣ���ɾ��ԭ��Ԫ�����ݣ����л�����
    '    ����ֱ���л����㵽��һ������û����һ��������ԭ��
    ' 4. ����Enter�����л����㵽��һ�У����û����һ�У����л�����ⰴť��
    ' 9. ���������ֱ���л����㵽��һ��
    '   �л����㷽����
    ' ���û����һ����Ԫ���ˣ����л�����ⰴť��
    
    Select Case KeyAscii
    Case 8 ' �˸��
        If Not txtTable(index).Locked And txtTable(index) <> "" Then
            txtTable(index) = ""
        Else
            If index > 0 Then txtTable(index - 1).SetFocus
        End If
        
    Case 13 ' �س���
        If index < 72 Then
            txtTable(index + 9).SetFocus
        Else
            cmdLogicSolve.SetFocus
        End If
        
    Case 49 To 57 ' ���� 1~9
        If Not txtTable(index).Locked Then
            txtTable(index) = "" ' ����ı����Զ�ȡ����
        End If
        If index < 80 Then
            txtTable(index + 1).SetFocus
        Else
            cmdLogicSolve.SetFocus
        End If
    
    Case 48 ' ����0
        KeyAscii = 0 ' �����ź�
        If Not txtTable(index).Locked Then
            txtTable(index) = ""
        End If
        If index < 80 Then
            txtTable(index + 1).SetFocus
        Else
            cmdLogicSolve.SetFocus
        End If
        
    Case Else
        KeyAscii = 0 ' �����ź�
        If index < 80 Then
            txtTable(index + 1).SetFocus
        Else
            cmdLogicSolve.SetFocus
        End If
        
    End Select

End Sub

Private Sub txtTable_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTable(index).ToolTipText = possibility(index)
End Sub
