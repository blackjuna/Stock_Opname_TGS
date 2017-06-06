VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form Progress SO"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20160
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh Data"
      Height          =   495
      Left            =   10440
      TabIndex        =   60
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Laporan"
      Height          =   495
      Left            =   10440
      TabIndex        =   59
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   10440
      Top             =   7080
   End
   Begin VB.Label ltr_total_tgs 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   85
      Top             =   7440
      Width           =   1740
   End
   Begin VB.Label lc_jkt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   84
      Top             =   6120
      Width           =   1740
   End
   Begin VB.Label lti_jkt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   83
      Top             =   6120
      Width           =   1740
   End
   Begin VB.Label ltr_jkt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   82
      Top             =   6120
      Width           =   1740
   End
   Begin VB.Label lc_sby 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   81
      Top             =   5160
      Width           =   1740
   End
   Begin VB.Label lti_sby 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   80
      Top             =   5160
      Width           =   1740
   End
   Begin VB.Label ltr_sby 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   79
      Top             =   5160
      Width           =   1740
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "SIP Surabaya"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   78
      Top             =   5160
      Width           =   1470
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tag Blank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   77
      Top             =   6120
      Width           =   1065
   End
   Begin VB.Label ltr_gaj 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   76
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label ltr_itj 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   75
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label lti_gaj 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   74
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lti_itj 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   73
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label lc_gaj 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   72
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lc_itj 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   71
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "GA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   70
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "IT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   69
      Top             =   3480
      Width           =   210
   End
   Begin VB.Label lc_itt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   68
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Label lc_gat 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   67
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Label lti_itt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   66
      Top             =   4680
      Width           =   1740
   End
   Begin VB.Label lti_gat 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   65
      Top             =   4200
      Width           =   1740
   End
   Begin VB.Label ltr_itt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   64
      Top             =   4680
      Width           =   1740
   End
   Begin VB.Label ltr_gat 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   63
      Top             =   4200
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   62
      Top             =   4680
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   61
      Top             =   4200
      Width           =   360
   End
   Begin VB.Label lbl_jam 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12000
      TabIndex        =   58
      Top             =   6480
      Width           =   2010
   End
   Begin VB.Label lbl_tanggal 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12120
      TabIndex        =   57
      Top             =   5760
      Width           =   2010
   End
   Begin VB.Label ltr_total_jfi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   56
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label lti_total_jfi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   55
      Top             =   4920
      Width           =   1740
   End
   Begin VB.Label lc_total_jfi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   54
      Top             =   4920
      Width           =   1545
   End
   Begin VB.Line Line9 
      X1              =   10320
      X2              =   17760
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label111 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   53
      Top             =   4920
      Width           =   1485
   End
   Begin VB.Line Line8 
      X1              =   10320
      X2              =   17760
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label110 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10560
      TabIndex        =   52
      Top             =   600
      Width           =   1305
   End
   Begin VB.Label Label109 
      AutoSize        =   -1  'True
      Caption         =   "% Complete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   15960
      TabIndex        =   51
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label Label108 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14040
      TabIndex        =   50
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label107 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Released"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12120
      TabIndex        =   49
      Top             =   600
      Width           =   1770
   End
   Begin VB.Label Label106 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "PT. JEIL FAJAR INDONESIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12735
      TabIndex        =   48
      Top             =   0
      Width           =   3645
   End
   Begin VB.Label Label105 
      AutoSize        =   -1  'True
      Caption         =   "Gudang JFI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   47
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label103 
      AutoSize        =   -1  'True
      Caption         =   "SWG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   46
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label Label102 
      AutoSize        =   -1  'True
      Caption         =   "Fluoroplastic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   45
      Top             =   2520
      Width           =   1350
   End
   Begin VB.Line Line7 
      X1              =   10320
      X2              =   10320
      Y1              =   480
      Y2              =   5400
   End
   Begin VB.Line Line6 
      X1              =   10320
      X2              =   17760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      X1              =   12000
      X2              =   12000
      Y1              =   480
      Y2              =   5400
   End
   Begin VB.Line Line4 
      X1              =   10320
      X2              =   17760
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   13920
      X2              =   13920
      Y1              =   480
      Y2              =   5400
   End
   Begin VB.Label ltr_fl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   44
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label ltr_swg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   43
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label ltr_gc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   42
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lti_fl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   41
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lti_swg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   40
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lti_gc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   39
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Line Line2 
      X1              =   15840
      X2              =   15840
      Y1              =   480
      Y2              =   5400
   End
   Begin VB.Line Line1 
      X1              =   17760
      X2              =   17760
      Y1              =   480
      Y2              =   5400
   End
   Begin VB.Label lc_fl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   38
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lc_swg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   37
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lc_gc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   36
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lti_total_tgs 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   35
      Top             =   7440
      Width           =   1740
   End
   Begin VB.Label lc_total_tgs 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   34
      Top             =   7440
      Width           =   1545
   End
   Begin VB.Line Line18 
      X1              =   1680
      X2              =   10080
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Label Label74 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   33
      Top             =   7440
      Width           =   1845
   End
   Begin VB.Line Line17 
      X1              =   1680
      X2              =   10080
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label73 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   32
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label72 
      AutoSize        =   -1  'True
      Caption         =   "% Complete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8400
      TabIndex        =   31
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label Label71 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6240
      TabIndex        =   30
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tag Released"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   29
      Top             =   600
      Width           =   1770
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "PT. TRIGRAHA SEALISINDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   28
      Top             =   0
      Width           =   8235
   End
   Begin VB.Label Label68 
      AutoSize        =   -1  'True
      Caption         =   "Gudang TGS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   27
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label Label65 
      AutoSize        =   -1  'True
      Caption         =   "Gudang F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   26
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label64 
      AutoSize        =   -1  'True
      Caption         =   "Gudang G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   25
      Top             =   1560
      Width           =   1125
   End
   Begin VB.Line Line16 
      X1              =   1680
      X2              =   1680
      Y1              =   480
      Y2              =   7920
   End
   Begin VB.Label Label61 
      AutoSize        =   -1  'True
      Caption         =   "Gland Packing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   24
      Top             =   2040
      Width           =   1545
   End
   Begin VB.Label Label60 
      AutoSize        =   -1  'True
      Caption         =   "Mech Seal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   23
      Top             =   2520
      Width           =   1125
   End
   Begin VB.Label Label59 
      Caption         =   "EJ - M  Flexible Hose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1800
      TabIndex        =   22
      Top             =   3000
      Width           =   1560
   End
   Begin VB.Label Label57 
      AutoSize        =   -1  'True
      Caption         =   "EJ - F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   21
      Top             =   3720
      Width           =   630
   End
   Begin VB.Line Line15 
      X1              =   1680
      X2              =   10080
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line14 
      X1              =   3840
      X2              =   3840
      Y1              =   480
      Y2              =   7920
   End
   Begin VB.Line Line13 
      X1              =   1680
      X2              =   10080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line12 
      X1              =   6000
      X2              =   6000
      Y1              =   480
      Y2              =   7920
   End
   Begin VB.Label ltr_ejm 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   20
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label ltr_ms 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   19
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label ltr_gp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   18
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label ltr_gg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12120
      TabIndex        =   17
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label ltr_gf 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   16
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label ltr_gA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   15
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label ltr_ejf 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   14
      Top             =   3720
      Width           =   1740
   End
   Begin VB.Label lti_ejm 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   13
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lti_ms 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   12
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lti_gp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   11
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lti_gg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14040
      TabIndex        =   10
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lti_gf 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   9
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lti_ga 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   8
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lti_ejf 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6240
      TabIndex        =   7
      Top             =   3720
      Width           =   1740
   End
   Begin VB.Line Line11 
      X1              =   8160
      X2              =   8160
      Y1              =   480
      Y2              =   7920
   End
   Begin VB.Line Line10 
      X1              =   10080
      X2              =   10080
      Y1              =   480
      Y2              =   7920
   End
   Begin VB.Label lc_ejm 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   6
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lc_ms 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   5
      Top             =   2520
      Width           =   1740
   End
   Begin VB.Label lc_gp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   4
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label lc_gg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15960
      TabIndex        =   3
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lc_gf 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   2
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label lc_ga 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8280
      TabIndex        =   1
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lc_ejf 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8400
      TabIndex        =   0
      Top             =   3720
      Width           =   1500
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub current_progress_jfi()
    'Gudang C
    qry_input_gc = "select count(tag_no) AS input_gc from tag_stock_opname_tgs where left(tag_no,3)='GJF' and status='OK'"
    Set rs_so = conn.Execute(qry_input_gc)
    lti_gc.Caption = Val(rs_so!input_gc)
    qry_release_gc = "select count(tag_no) AS release_gc from tag_stock_opname_tgs where left(tag_no,3)='GJF'"
    Set rs_so = conn.Execute(qry_release_gc)
    ltr_gc.Caption = Val(rs_so!release_gc)
    If Val(rs_so!release_gc) = 0 Then lc_gc.Caption = 0 Else _
        lc_gc.Caption = Round((Val(lti_gc.Caption) / Val(ltr_gc.Caption) * 100), 2)
'
'    'Gudang E
'    qry_input_ge = "select count(tag_no) AS input_ge from tag_stock_opname_tgs where left(tag_no,2)='GE' and status='OK'"
'    Set rs_so = conn.Execute(qry_input_ge)
'    lti_ge.Caption = Val(rs_so!input_ge)
'    qry_release_ge = "select count(tag_no) AS release_ge from tag_stock_opname_tgs where left(tag_no,2)='GE'"
'    Set rs_so = conn.Execute(qry_release_ge)
'    ltr_ge.Caption = Val(rs_so!release_ge)
'    If Val(rs_so!release_ge) = 0 Then lc_ge.Caption = 0 Else _
'        lc_ge.Caption = Round((Val(lti_ge.Caption) / Val(ltr_ge.Caption) * 100), 2)
'
    'Gudang G
    qry_input_gg = "select count(tag_no) AS input_gg from tag_stock_opname_tgs where left(tag_no,2)='GG' and status='OK'"
    Set rs_so = conn.Execute(qry_input_gg)
    lti_gg.Caption = Val(rs_so!input_gg)
    qry_release_gg = "select count(tag_no) AS release_gg from tag_stock_opname_tgs where left(tag_no,2)='GG'"
    Set rs_so = conn.Execute(qry_release_gg)
    ltr_gg.Caption = Val(rs_so!release_gg)
    If Val(rs_so!release_gg) = 0 Then lc_gg.Caption = 0 Else _
        lc_gg.Caption = Round((Val(lti_gg.Caption) / Val(ltr_gg.Caption) * 100), 2)

    
    'Flouroplastic
    qry_input_fl = "select count(tag_no) AS input_fl from tag_stock_opname_tgs where left(tag_no,2)='FP' and status='OK'"
    Set rs_so = conn.Execute(qry_input_fl)
    lti_fl.Caption = Val(rs_so!input_fl)
    qry_release_fl = "select count(tag_no) AS release_fl from tag_stock_opname_tgs where left(tag_no,2)='FP'"
    Set rs_so = conn.Execute(qry_release_fl)
    ltr_fl.Caption = Val(rs_so!release_fl)
    If Val(rs_so!release_fl) = 0 Then lc_fl.Caption = 0 Else _
        lc_fl.Caption = Round((Val(lti_fl.Caption) / Val(ltr_fl.Caption) * 100), 2)
    
    'SWG
    qry_input_sw = "select count(tag_no) AS input_sw from tag_stock_opname_tgs where left(tag_no,2)='SW' and status='OK'"
    Set rs_so = conn.Execute(qry_input_sw)
    lti_swg.Caption = Val(rs_so!input_sw)
    qry_release_sw = "select count(tag_no) AS release_sw from tag_stock_opname_tgs where left(tag_no,2)='SW'"
    Set rs_so = conn.Execute(qry_release_sw)
    ltr_swg.Caption = Val(rs_so!release_sw)
    If Val(rs_so!release_sw) = 0 Then lc_swg.Caption = 0 Else _
        lc_swg.Caption = Round((Val(lti_swg.Caption) / Val(ltr_swg.Caption) * 100), 2)
        
          'GA - JFI
    qry_input_gaj = "select count(tag_no) AS input_gaj from tag_stock_opname_tgs where left(tag_no,3)='GAJ' and status='OK'"
    Set rs_so = conn.Execute(qry_input_gaj)
    lti_gaj.Caption = Val(rs_so!input_gaj)
    qry_release_gaj = "select count(tag_no) AS release_gaj from tag_stock_opname_tgs where left(tag_no,3)='GAJ'"
    Set rs_so = conn.Execute(qry_release_gaj)
    ltr_gaj.Caption = Val(rs_so!release_gaj)
    If Val(rs_so!release_gaj) = 0 Then lc_gaj.Caption = 0 Else _
    lc_gaj.Caption = Round((Val(lti_gaj.Caption) / Val(ltr_gaj.Caption) * 100), 2)
    
            'IT - JFI
    qry_input_itj = "select count(tag_no) AS input_itj from tag_stock_opname_tgs where left(tag_no,3)='ITJ' and status='OK'"
    Set rs_so = conn.Execute(qry_input_itj)
    lti_itj.Caption = Val(rs_so!input_itj)
    qry_release_itj = "select count(tag_no) AS release_itj from tag_stock_opname_tgs where left(tag_no,3)='ITJ'"
    Set rs_so = conn.Execute(qry_release_itj)
    ltr_itj.Caption = Val(rs_so!release_itj)
    If Val(rs_so!release_itj) = 0 Then lc_itj.Caption = 0 Else _
    lc_itj.Caption = Round((Val(lti_itj.Caption) / Val(ltr_itj.Caption) * 100), 2)
    
    'Total
    ltr_total_jfi.Caption = Val(ltr_gc.Caption) + Val(ltr_gg.Caption) + Val(ltr_swg.Caption) + _
        Val(ltr_fl.Caption) + Val(ltr_itj.Caption) + Val(ltr_gaj.Caption)
    lti_total_jfi.Caption = Val(lti_gc.Caption) + Val(lti_gg.Caption) + Val(lti_swg.Caption) + _
        Val(lti_fl.Caption) + Val(lti_itj.Caption) + Val(lti_gaj.Caption)
    If Val(lti_total_jfi.Caption) = 0 Then lc_total_jfi.Caption = 0 Else _
        lc_total_jfi.Caption = Round(((Val(lti_total_jfi) / Val(ltr_total_jfi) * 100)), 2)
    
End Sub
Sub current_progress_tgs()
    'Gudang A
    qry_input_ga = "select count(tag_no) AS input_ga from tag_stock_opname_tgs where left(tag_no,3)='GTS' and status='OK'"
    Set rs_so = conn.Execute(qry_input_ga)
    lti_ga.Caption = Val(rs_so!input_ga)
    qry_release_ga = "select count(tag_no) AS release_ga from tag_stock_opname_tgs where left(tag_no,3)='GTS'"
    Set rs_so = conn.Execute(qry_release_ga)
    ltr_gA.Caption = Val(rs_so!release_ga)
    If Val(rs_so!release_ga) = 0 Then lc_ga.Caption = 0 Else _
        lc_ga.Caption = Round((Val(lti_ga.Caption) / Val(ltr_gA.Caption) * 100), 2)
'
'    'Gudang B
'    qry_input_gb = "select count(tag_no) AS input_gb from tag_stock_opname_tgs where left(tag_no,2)='GB' and status='OK'"
'    Set rs_so = conn.Execute(qry_input_gb)
'    lti_gb.Caption = Val(rs_so!input_gb)
'    qry_release_gb = "select count(tag_no) AS release_gb from tag_stock_opname_tgs where left(tag_no,2)='GB'"
'    Set rs_so = conn.Execute(qry_release_gb)
'    ltr_gb.Caption = Val(rs_so!release_gb)
'    If Val(rs_so!release_gb) = 0 Then lc_gb.Caption = 0 Else _
'        lc_gb.Caption = Round((Val(lti_gb.Caption) / Val(ltr_gb.Caption) * 100), 2)
'
'    'Gudang D
'    qry_input_gd = "select count(tag_no) AS input_gd from tag_stock_opname_tgs where left(tag_no,2)='GD' and status='OK'"
'    Set rs_so = conn.Execute(qry_input_gd)
'    lti_gd.Caption = Val(rs_so!input_gd)
'    qry_release_gd = "select count(tag_no) AS release_gd from tag_stock_opname_tgs where left(tag_no,2)='GD'"
'    Set rs_so = conn.Execute(qry_release_gd)
'    ltr_gd.Caption = Val(rs_so!release_gd)
'    If Val(rs_so!release_gd) = 0 Then lc_gd.Caption = 0 Else _
'        lc_gd.Caption = Round((Val(lti_gd.Caption) / Val(ltr_gd.Caption) * 100), 2)
'
    'Gudang F
    qry_input_gf = "select count(tag_no) AS input_gf from tag_stock_opname_tgs where left(tag_no,2)='GF' and status='OK'"
    Set rs_so = conn.Execute(qry_input_gf)
    lti_gf.Caption = Val(rs_so!input_gf)
    qry_release_gf = "select count(tag_no) AS release_gf from tag_stock_opname_tgs where left(tag_no,2)='GF'"
    Set rs_so = conn.Execute(qry_release_gf)
    ltr_gf.Caption = Val(rs_so!release_gf)
    If Val(rs_so!release_gf) = 0 Then lc_gf.Caption = 0 Else _
        lc_gf.Caption = Round((Val(lti_gf.Caption) / Val(ltr_gf.Caption) * 100), 2)
        
    'Pencelupan PTFE
'    qry_input_pp = "select count(tag_no) AS input_pp from tag_stock_opname_tgs where left(tag_no,2)='PP' and status='OK'"
'    Set rs_so = conn.Execute(qry_input_pp)
'    lti_pp.Caption = Val(rs_so!input_pp)
'    qry_release_pp = "select count(tag_no) AS release_pp from tag_stock_opname_tgs where left(tag_no,2)='PP'"
'    Set rs_so = conn.Execute(qry_release_pp)
'    ltr_pp.Caption = Val(rs_so!release_pp)
'    If Val(rs_so!release_pp) = 0 Then ltr_pp.Caption = 0 Else _
'        lc_pp.Caption = Round((Val(lti_pp.Caption) / Val(ltr_pp.Caption) * 100), 2)
'
'    'Man Hole
'    qry_input_mh = "select count(tag_no) AS input_mh from tag_stock_opname_tgs where left(tag_no,2)='MH' and status='OK'"
'    Set rs_so = conn.Execute(qry_input_mh)
'    lti_mh.Caption = Val(rs_so!input_mh)
'    qry_release_mh = "select count(tag_no) AS release_mh from tag_stock_opname_tgs where left(tag_no,2)='MH'"
'    Set rs_so = conn.Execute(qry_release_mh)
'    ltr_mh.Caption = Val(rs_so!release_mh)
'    If Val(rs_so!release_mh) = 0 Then lc_mh.Caption = 0 Else _
'        lc_mh.Caption = Round((Val(lti_mh.Caption) / Val(ltr_mh.Caption) * 100), 2)
    
    'Gland Packing
    qry_input_gp = "select count(tag_no) AS input_gp from tag_stock_opname_tgs where left(tag_no,2)='GP' and status='OK'"
    Set rs_so = conn.Execute(qry_input_gp)
    lti_gp.Caption = Val(rs_so!input_gp)
    qry_release_gp = "select count(tag_no) AS release_gp from tag_stock_opname_tgs where left(tag_no,2)='GP'"
    Set rs_so = conn.Execute(qry_release_gp)
    ltr_gp.Caption = Val(rs_so!release_gp)
    If Val(rs_so!release_gp) = 0 Then lc_gp.Caption = 0 Else _
        lc_gp.Caption = Round((Val(lti_gp.Caption) / Val(ltr_gp.Caption) * 100), 2)
    
    'Mech Seal
    qry_input_ms = "select count(tag_no) AS input_ms from tag_stock_opname_tgs where left(tag_no,2)='MS' and status='OK'"
    Set rs_so = conn.Execute(qry_input_ms)
    lti_ms.Caption = Val(rs_so!input_ms)
    qry_release_ms = "select count(tag_no) AS release_ms from tag_stock_opname_tgs where left(tag_no,2)='MS'"
    Set rs_so = conn.Execute(qry_release_ms)
    ltr_ms.Caption = Val(rs_so!release_ms)
    If Val(rs_so!release_ms) = 0 Then lc_ms.Caption = 0 Else _
        lc_ms.Caption = Round((Val(lti_ms.Caption) / Val(ltr_ms.Caption) * 100), 2)
    
    'EJ - Metal
    qry_input_em = "select count(tag_no) AS input_em from tag_stock_opname_tgs where left(tag_no,2)='EM' and status='OK'"
    Set rs_so = conn.Execute(qry_input_em)
    lti_ejm.Caption = Val(rs_so!input_em)
    qry_release_em = "select count(tag_no) AS release_em from tag_stock_opname_tgs where left(tag_no,2)='EM'"
    Set rs_so = conn.Execute(qry_release_em)
    ltr_ejm.Caption = Val(rs_so!release_em)
    If Val(rs_so!release_em) = 0 Then lc_ejm.Caption = 0 Else _
        lc_ejm.Caption = Round((Val(lti_ejm.Caption) / Val(ltr_ejm.Caption) * 100), 2)
    
    'Flexible Hose
'    qry_input_fh = "select count(tag_no) AS input_fh from tag_stock_opname_tgs where left(tag_no,2)='FH' and status='OK'"
'    Set rs_so = conn.Execute(qry_input_fh)
'    lti_fh.Caption = Val(rs_so!input_fh)
'    qry_release_fh = "select count(tag_no) AS release_fh from tag_stock_opname_tgs where left(tag_no,2)='FH'"
'    Set rs_so = conn.Execute(qry_release_fh)
'    ltr_fh.Caption = Val(rs_so!release_fh)
'    If Val(rs_so!release_fh) = 0 Then lc_fh.Caption = 0 Else _
'    lc_fh.Caption = Round((Val(lti_fh.Caption) / Val(ltr_fh.Caption) * 100), 2)
    
    'EJ - Fabric
    qry_input_ef = "select count(tag_no) AS input_ef from tag_stock_opname_tgs where left(tag_no,2)='EF' and status='OK'"
    Set rs_so = conn.Execute(qry_input_ef)
    lti_ejf.Caption = Val(rs_so!input_ef)
    qry_release_ef = "select count(tag_no) AS release_ef from tag_stock_opname_tgs where left(tag_no,2)='EF'"
    Set rs_so = conn.Execute(qry_release_ef)
    ltr_ejf.Caption = Val(rs_so!release_ef)
    If Val(rs_so!release_ef) = 0 Then lc_ejf.Caption = 0 Else _
        lc_ejf.Caption = Round((Val(lti_ejf.Caption) / Val(ltr_ejf.Caption) * 100), 2)
    
        'GA - TGS
    qry_input_gat = "select count(tag_no) AS input_gat from tag_stock_opname_tgs where left(tag_no,3)='GAT' and status='OK'"
    Set rs_so = conn.Execute(qry_input_gat)
    lti_gat.Caption = Val(rs_so!input_gat)
    qry_release_gat = "select count(tag_no) AS release_gat from tag_stock_opname_tgs where left(tag_no,3)='GAT'"
    Set rs_so = conn.Execute(qry_release_gat)
    ltr_gat.Caption = Val(rs_so!release_gat)
    If Val(rs_so!release_gat) = 0 Then lc_gat.Caption = 0 Else _
    lc_gat.Caption = Round((Val(lti_gat.Caption) / Val(ltr_gat.Caption) * 100), 2)
    
                'IT - TGS
    qry_input_itt = "select count(tag_no) AS input_itt from tag_stock_opname_tgs where left(tag_no,3)='ITT' and status='OK'"
    Set rs_so = conn.Execute(qry_input_itt)
    lti_itt.Caption = Val(rs_so!input_itt)
    qry_release_itt = "select count(tag_no) AS release_itt from tag_stock_opname_tgs where left(tag_no,3)='ITT'"
    Set rs_so = conn.Execute(qry_release_itt)
    ltr_itt.Caption = Val(rs_so!release_itt)
    If Val(rs_so!release_itt) = 0 Then lc_itt.Caption = 0 Else _
    lc_itt.Caption = Round((Val(lti_itt.Caption) / Val(ltr_itt.Caption) * 100), 2)
    
                    'SIP - Surabaya
    qry_input_sby = "select count(tag_no) AS input_sby from tag_stock_opname_tgs where left(tag_no,3)='SBY' and status='OK'"
    Set rs_so = conn.Execute(qry_input_sby)
    lti_sby.Caption = Val(rs_so!input_sby)
    qry_release_SBY = "select count(tag_no) AS release_SBY from tag_stock_opname_tgs where left(tag_no,3)='SBY'"
    Set rs_so = conn.Execute(qry_release_SBY)
    ltr_sby.Caption = Val(rs_so!release_SBY)
    If Val(rs_so!release_SBY) = 0 Then lc_sby.Caption = 0 Else _
    lc_sby.Caption = Round((Val(lti_sby.Caption) / Val(ltr_sby.Caption) * 100), 2)
    
'                        'Konsinyasi
'    qry_input_ksy = "select count(tag_no) AS input_ksy from tag_stock_opname_tgs where left(tag_no,3)='KSY' and status='OK'"
'    Set rs_so = conn.Execute(qry_input_ksy)
'    lti_ksy.Caption = Val(rs_so!input_ksy)
'    qry_release_ksy = "select count(tag_no) AS release_ksy from tag_stock_opname_tgs where left(tag_no,3)='KSY'"
'    Set rs_so = conn.Execute(qry_release_ksy)
'    ltr_ksy.Caption = Val(rs_so!release_ksy)
'    If Val(rs_so!release_ksy) = 0 Then lc_ksy.Caption = 0 Else _
'    lc_ksy.Caption = Round((Val(lti_ksy.Caption) / Val(ltr_ksy.Caption) * 100), 2)
    
                        'TAG BLANK
    qry_input_tb = "select count(tag_no) AS input_tb from tag_stock_opname_tgs where left(tag_no,2)='TB' and status='OK'"
    Set rs_so = conn.Execute(qry_input_tb)
    lti_jkt.Caption = Val(rs_so!input_tb)
    qry_release_tb = "select count(tag_no) AS release_TB from tag_stock_opname_tgs where left(tag_no,2)='TB'"
    Set rs_so = conn.Execute(qry_release_tb)
    ltr_jkt.Caption = Val(rs_so!release_TB)
    If Val(rs_so!release_TB) = 0 Then lc_jkt.Caption = 0 Else _
    lc_jkt.Caption = Round((Val(lti_jkt.Caption) / Val(ltr_jkt.Caption) * 100), 2)
    
    'Total
    ltr_total_tgs.Caption = Val(ltr_gA.Caption) + Val(ltr_gf.Caption) _
        + Val(ltr_gp.Caption) + Val(ltr_ms.Caption) + Val(ltr_ejm.Caption) _
        + Val(ltr_ejf.Caption) + Val(ltr_itt.Caption) + Val(ltr_gat.Caption) + Val(ltr_jkt.Caption) + _
        Val(ltr_sby.Caption)
    'lti_total_tgs.Caption = Val(lti_ga.Caption) + Val(lti_gb.Caption) + Val(lti_gd.Caption) + Val(lti_gf.Caption) + Val(lti_pp.Caption) + Val(lti_mh.Caption) + Val(lti_gp.Caption) + Val(lti_ms.Caption) + Val(lti_ejm.Caption) + Val(lti_fh.Caption) + Val(lti_ejf.Caption)
    lti_total_tgs.Caption = Val(lti_ga.Caption) + Val(lti_gf.Caption) _
        + Val(lti_gp.Caption) + Val(lti_ms.Caption) + Val(lti_ejm.Caption) _
        + Val(lti_ejf.Caption) + Val(lti_itt.Caption) + Val(lti_gat.Caption) + Val(lti_jkt.Caption) _
        + Val(lti_sby.Caption)
    If Val(lti_total_tgs.Caption) = 0 Then lc_total_tgs.Caption = 0 Else _
        lc_total_tgs.Caption = Round(((Val(lti_total_tgs) / Val(ltr_total_tgs) * 100)), 2)
    
    'lcttl.Caption=ltr
    
End Sub


Private Sub Command1_Click()
    Printer.PaperSize = vbPRPSLegal
    Printer.Orientation = vbPRORLandscape
    Form2.PrintForm
End Sub

Private Sub Command2_Click()
Call current_progress_jfi
Call current_progress_tgs

End Sub

Private Sub Form_Load()
Timer1.Interval = 500
Timer1.Enabled = True
Call db
If rs_so.State = 1 Then rs_so.Close
rs_so.Open "Select * from tag_stock_opname_tgs", conn
Set rscompletion_slip = Nothing
Call current_progress_jfi
Call current_progress_tgs
lbl_tanggal = Date
End Sub

Private Sub Timer1_Timer()
lbl_jam.Caption = Format(Time, "hh:mm:ss")
End Sub
