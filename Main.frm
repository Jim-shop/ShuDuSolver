VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "����"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8385
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
   ScaleHeight     =   6240
   ScaleWidth      =   8385
   StartUpPosition =   3  '����ȱʡ
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdImport 
      Appearance      =   0  'Flat
      Caption         =   "����"
      Height          =   495
      Left            =   4200
      TabIndex        =   85
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtInput 
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
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   84
      Top             =   4920
      Width           =   3855
   End
   Begin VB.CommandButton cmdSolve 
      Appearance      =   0  'Flat
      Caption         =   "���"
      Height          =   495
      Left            =   6720
      TabIndex        =   82
      Top             =   960
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
      Left            =   6720
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblDebug 
      Height          =   3975
      Left            =   6120
      TabIndex        =   83
      Top             =   2160
      Width           =   2175
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

' ��¼һ����������ܳ��ֵ����֣��Ѿ�ȷ������������ȫfalse��
Private Type Possibility
    p(9) As Boolean
End Type

' ��¼ÿ����������ܳ��ֵ�����
Private storage(81) As Possibility

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
                    For j = 0 To 8
                        storage(i).p(j) = True
                    Next j
                Else ' �������ݵ���Ϣ��
                    .Locked = True
                    .BackColor = &HDDDDDD
                    For j = 0 To 8
                        storage(i).p(j) = False
                    Next j
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

Private Sub cmdClear_Click()
    ClearQuestion
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
        If tablepos >= 80 Then Exit Do
        txtpos = txtpos + 1
    Loop
End Sub

Private Sub cmdSolve_Click()
    ' ��������
    LoadQuestion
    Dim Index%, row%, column%, count%, i%, j%, tempNum%, sumPossibility%, lastPossibleNum%
    DebugInfo "", "", True
    Do
        For Index = 0 To 80
            If txtTable(Index) <> "" Then ' ����Ѿ���������
                column = Index Mod 9
                row = Index \ 9
                tempNum = Val(txtTable(Index)) - 1
                
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
        Next Index
        
        ' ��û�����������Եĵ���д����
        count = 0
        For Index = 0 To 80
            sumPossibility = 0
            For i = 0 To 8
                If storage(Index).p(i) = True Then
                    sumPossibility = sumPossibility + 1
                    lastPossibleNum = i
                End If
            Next i
            If sumPossibility = 1 Then
              count = count + 1
                txtTable(Index) = lastPossibleNum + 1
            End If
        Next Index
        DebugInfo "�Ƶ���" & count & "��λ��"
    Loop While count <> 0 ' ��û�н�չʱ����
End Sub

Private Sub PrintPossibility(Index%)
    ' ��ӡ��������Ϣ
    
    DebugInfo "txtTable(" & Index & "):", , True
    
    Dim i%
    For i = 0 To 8
        If storage(Index).p(i) = True Then
            DebugInfo (i + 1), " "
        Else
            DebugInfo " ", " "
        End If
        If i Mod 3 = 2 Then DebugInfo ""
    Next i
End Sub

Private Sub DebugInfo(ByRef message$, Optional split$ = vbNewLine, Optional needRefresh As Boolean = False)
    ' ��lblDebug�д�ӡ������Ϣ��
    ' needRefresh�����Ƿ��ͷ��ʼ��ӡ��
    
    If needRefresh Then
        lblDebug.Caption = message & split
    Else
        lblDebug.Caption = lblDebug.Caption & message & split
    End If
End Sub

Private Sub Form_Load()
    isQuestionLoad = False
End Sub

Private Sub txtTable_KeyPress(Index As Integer, KeyAscii As Integer)
    '   �����ַ�ʱ�Ĵ���ʽ
    ' 1. ����������֣����ǻ��Ԫ�����滻ԭ��Ԫ�����ݣ�Ȼ���л����㵽��һ����Ԫ��
    ' 2. ��������˸�������ǻ��Ԫ���������ݣ���ɾ��ԭ��Ԫ�����ݣ����л�����
    '    ����ֱ���л����㵽��һ������û����һ��������ԭ��
    ' 3. ����Enter�����л����㵽��һ�У����û����һ�У����л�����ⰴť��
    ' 9. ���������ֱ���л����㵽��һ��
    '   �л����㷽����
    ' ���û����һ����Ԫ���ˣ����л�����ⰴť��
    
    Select Case KeyAscii
    Case 8 ' �˸��
        If Not txtTable(Index).Locked And txtTable(Index) <> "" Then
            txtTable(Index) = ""
        Else
            If Index > 0 Then txtTable(Index - 1).SetFocus
        End If
        
    Case 13 ' �س���
        If Index < 72 Then
            txtTable(Index + 9).SetFocus
        Else
            cmdSolve.SetFocus
        End If
        
    Case 49 To 57 ' ���� 1~9
        If Not txtTable(Index).Locked Then
            txtTable(Index) = "" ' ����ı����Զ�ȡ����
        End If
        If Index < 80 Then
            txtTable(Index + 1).SetFocus
        Else
            cmdSolve.SetFocus
        End If
        
    Case Else
        KeyAscii = 0 ' �����ź�
        If Index < 80 Then
            txtTable(Index + 1).SetFocus
        Else
            cmdSolve.SetFocus
        End If
        
    End Select

End Sub

Private Sub txtTable_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PrintPossibility Index
End Sub
