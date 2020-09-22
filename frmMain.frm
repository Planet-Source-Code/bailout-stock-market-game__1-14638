VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                So... You Wanna Make Some Cash?"
   ClientHeight    =   5760
   ClientLeft      =   915
   ClientTop       =   1560
   ClientWidth     =   18750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   18750
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   825
      TabIndex        =   113
      Top             =   4320
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   17040
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   17040
      Top             =   1560
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   12240
      TabIndex        =   109
      Top             =   0
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   17040
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Over"
      Height          =   375
      Left            =   8160
      TabIndex        =   107
      Top             =   0
      Width           =   4095
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "Sell"
      Height          =   375
      Left            =   12240
      TabIndex        =   57
      Top             =   3990
      Width           =   4095
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Height          =   375
      Left            =   8160
      TabIndex        =   56
      Top             =   3990
      Width           =   4095
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   30000
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   30000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   44
      Top             =   7680
      Width           =   150
   End
   Begin VB.Frame Frame1 
      Height          =   4320
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      Begin VB.PictureBox picTitleBar 
         Height          =   645
         Left            =   110
         ScaleHeight     =   585
         ScaleWidth      =   7875
         TabIndex        =   35
         Top             =   145
         Width           =   7935
         Begin VB.CommandButton cmdCommand1 
            Height          =   195
            Left            =   4140
            TabIndex        =   45
            Top             =   675
            Width           =   375
         End
         Begin VB.CommandButton cmdGetQuote 
            Caption         =   "&Get Quote"
            Height          =   310
            Left            =   1770
            TabIndex        =   41
            Top             =   180
            Width           =   1032
         End
         Begin VB.CommandButton cmdViewGraph 
            Caption         =   "&View Graph"
            Height          =   310
            Left            =   6390
            TabIndex        =   40
            Top             =   135
            Width           =   1455
         End
         Begin VB.ComboBox cboGraph 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   39
            ToolTipText     =   "Select the type of graph you would like"
            Top             =   180
            Width           =   1512
         End
         Begin VB.ComboBox cboSymbol 
            Height          =   315
            Left            =   640
            TabIndex        =   38
            Text            =   "intc"
            ToolTipText     =   "Enter you stock symbol here"
            Top             =   180
            Width           =   1092
         End
         Begin VB.CommandButton cmdAddSymbol 
            Caption         =   "+"
            Height          =   312
            Left            =   2850
            TabIndex        =   37
            ToolTipText     =   "Add symbol to list"
            Top             =   180
            Width           =   372
         End
         Begin VB.CommandButton cmdRemoveSymbol 
            Caption         =   "-"
            Height          =   312
            Left            =   3270
            TabIndex        =   36
            ToolTipText     =   "Remove symbol from list"
            Top             =   180
            Width           =   372
         End
         Begin VB.Label Label20 
            Caption         =   "Lookup:"
            Height          =   255
            Left            =   30
            TabIndex        =   43
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Graph Type:"
            Height          =   255
            Left            =   3750
            TabIndex        =   42
            Top             =   225
            Width           =   915
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Market Capitilization:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   14
         Left            =   3960
         TabIndex        =   11
         Top             =   3225
         Width           =   2895
      End
      Begin VB.Label lblConnection 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "You are currently connected to the internet "
         Height          =   255
         Left            =   90
         TabIndex        =   48
         Top             =   3960
         Width           =   5055
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5220
         TabIndex        =   47
         Top             =   3960
         Width           =   1365
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6660
         TabIndex        =   46
         Top             =   3960
         Width           =   1365
      End
      Begin VB.Label lblDateTime 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   4905
         TabIndex        =   1
         Top             =   945
         Width           =   3075
      End
      Begin VB.Label lblLastGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   25
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Label lblSEGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6315
         TabIndex        =   2
         Top             =   3540
         Width           =   1725
      End
      Begin VB.Label lblPERGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   18
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label lblOpenGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   24
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label lblHighGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   23
         Top             =   1965
         Width           =   1215
      End
      Begin VB.Label lblLowGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   22
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblVolumeGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   21
         Top             =   2595
         Width           =   1215
      End
      Begin VB.Label lblPSEGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   20
         Top             =   2910
         Width           =   1215
      End
      Begin VB.Label lblPSOGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2835
         TabIndex        =   19
         Top             =   3225
         Width           =   1215
      End
      Begin VB.Label lblChangeGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6780
         TabIndex        =   9
         Top             =   1335
         Width           =   1245
      End
      Begin VB.Label lblChangepGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6795
         TabIndex        =   8
         Top             =   1650
         Width           =   1245
      End
      Begin VB.Label lbl52HighGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6810
         TabIndex        =   7
         Top             =   1965
         Width           =   1245
      End
      Begin VB.Label lbl52LowGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6795
         TabIndex        =   6
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label lblBidGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6795
         TabIndex        =   5
         Top             =   2595
         Width           =   1245
      End
      Begin VB.Label lblAskGet 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6825
         TabIndex        =   4
         Top             =   2910
         Width           =   1245
      End
      Begin VB.Label lblMCGet 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6795
         TabIndex        =   3
         Top             =   3225
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Stock Exchange:"
         Height          =   255
         Index           =   15
         Left            =   4050
         TabIndex        =   10
         Top             =   3540
         Width           =   2775
      End
      Begin VB.Label lblSymbol 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   90
         TabIndex        =   34
         Top             =   855
         Width           =   7935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Last:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   33
         Top             =   1335
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Open:"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   32
         Top             =   1650
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "High:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   31
         Top             =   1965
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Low:"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   30
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Volume:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   29
         Top             =   2595
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Per Share Earning:"
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   28
         Top             =   2910
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Per Share Outstanding:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   6
         Left            =   90
         TabIndex        =   27
         Top             =   3225
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "P/E Ratio:"
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   26
         Top             =   3540
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Change:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   8
         Left            =   4050
         TabIndex        =   17
         Top             =   1335
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Change(%):"
         Height          =   255
         Index           =   9
         Left            =   4050
         TabIndex        =   16
         Top             =   1650
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "52 Week High:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   10
         Left            =   4050
         TabIndex        =   15
         Top             =   1965
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "52 Week Low:"
         Height          =   255
         Index           =   11
         Left            =   4050
         TabIndex        =   14
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Bid:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   12
         Left            =   4035
         TabIndex        =   13
         Top             =   2595
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Ask:"
         Height          =   255
         Index           =   13
         Left            =   4050
         TabIndex        =   12
         Top             =   2910
         Width           =   2775
      End
   End
   Begin InetCtlsObjects.Inet inetQuotes 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   9120
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13920
      TabIndex        =   112
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Total Holdings:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   111
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Cash to Spend:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   110
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblBudget 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   108
      Top             =   480
      Width           =   1935
   End
   Begin VB.Line Line12 
      X1              =   8160
      X2              =   16320
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line10 
      X1              =   8160
      X2              =   16320
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line Line9 
      X1              =   8160
      X2              =   8160
      Y1              =   3960
      Y2              =   960
   End
   Begin VB.Line Line8 
      X1              =   16320
      X2              =   8160
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblValueChng 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   14760
      TabIndex        =   106
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblValueChng 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   14760
      TabIndex        =   105
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblValueChng 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   14760
      TabIndex        =   104
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblValueChng 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   14760
      TabIndex        =   103
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblValueChng 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   14760
      TabIndex        =   102
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblValueChng 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   14760
      TabIndex        =   101
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblValueChng 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   14760
      TabIndex        =   100
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   13320
      TabIndex        =   99
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   13320
      TabIndex        =   98
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   13320
      TabIndex        =   97
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   13320
      TabIndex        =   96
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   13320
      TabIndex        =   95
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   13320
      TabIndex        =   94
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   13320
      TabIndex        =   93
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblPaidPts 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   12480
      TabIndex        =   92
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblPaidPts 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   12480
      TabIndex        =   91
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblPaidPts 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   12480
      TabIndex        =   90
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblPaidPts 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   12480
      TabIndex        =   89
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblPaidPts 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   12480
      TabIndex        =   88
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblPaidPts 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   12480
      TabIndex        =   87
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblPaidPts 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   12480
      TabIndex        =   86
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblShares 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   11640
      TabIndex        =   85
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblShares 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   11640
      TabIndex        =   84
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblShares 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   11640
      TabIndex        =   83
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblShares 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   11640
      TabIndex        =   82
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblShares 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   11640
      TabIndex        =   81
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblShares 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   11640
      TabIndex        =   80
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblShares 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   11640
      TabIndex        =   79
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblPtChange 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   10680
      TabIndex        =   78
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblPtChange 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   10680
      TabIndex        =   77
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblPtChange 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   10680
      TabIndex        =   76
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblPtChange 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   10680
      TabIndex        =   75
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblPtChange 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   10680
      TabIndex        =   74
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblPtChange 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   10680
      TabIndex        =   73
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblPtChange 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   10680
      TabIndex        =   72
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   9240
      TabIndex        =   71
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9240
      TabIndex        =   70
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9240
      TabIndex        =   69
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   9240
      TabIndex        =   68
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9240
      TabIndex        =   67
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9240
      TabIndex        =   66
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9240
      TabIndex        =   65
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblSym 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   8280
      TabIndex        =   64
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label lblSym 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   8280
      TabIndex        =   63
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblSym 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8280
      TabIndex        =   62
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblSym 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   8280
      TabIndex        =   61
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblSym 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   60
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblSym 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   59
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblSym 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8280
      TabIndex        =   58
      Top             =   1440
      Width           =   735
   End
   Begin VB.Line Line7 
      X1              =   16320
      X2              =   16320
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Line Line6 
      X1              =   14640
      X2              =   14640
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Line Line5 
      X1              =   13200
      X2              =   13200
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Line Line4 
      X1              =   12360
      X2              =   12360
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      X1              =   11520
      X2              =   11520
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      X1              =   10560
      X2              =   10560
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      X1              =   9120
      X2              =   9120
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Value Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14760
      TabIndex        =   55
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13320
      TabIndex        =   54
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Paid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   53
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Shares"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11640
      TabIndex        =   52
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   51
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Current Price"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   50
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Symbol"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   49
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
' Set the following constant to True to load
' from a test data file instead of from
' the Web.
'Private Const USE_TEST_FILE = True
Private Const USE_TEST_FILE = False
Dim Simble As String
Dim stockman As Long
Dim Budget, budgeta As Double
Dim stock As Long
Dim symb As String
Dim AppName As String 'Used for registry
Dim Section As String 'Used for registry
Dim sCustKey As String * 50 'Used for registry
Dim sCustVal As String * 50 'Used for registry
Dim Counter1 As Long 'Used to enumerate the registry entries
Dim counter As Long  'Used to enumerate the registry entries
Dim j As String 'For Delete for loop
Dim j1 As Long 'For Delete for loop
Dim Mysettings1 As Variant, intSettings1 As Long 'For custlist
Dim intSet1 As Long 'For custlist
Dim symbols As String 'For GetQuote

Private Sub cmdBuy_Click()
Dim shares As Long
Dim tester As Long
shares = InputBox("How many shares would you like to buy of " & symb & "?")
If shares * lblLastGet.Caption > Budget Then
    shares = Budget / lblLastGet.Caption
End If
If shares <= 0 Then
    MsgBox "Sorry your BROKE!!!"
    GoTo Enderss
End If
If MsgBox("Have you updated this lately?", vbYesNo) = vbNo Then
    GoTo Enderss
Else
For tester = 0 To 6
    If symb = lblSym(tester).Caption And lblLastGet.Caption = lblPaidPts(tester).Caption Then
        YesNo = 1
        GoTo alreadyhave
    End If
Next
lblSym(stockman).Caption = symb
Write_Ini "Stock" & (stockman), "Symbol", symb
lblPrice(stockman).Caption = lblLastGet.Caption
lblPtChange(stockman).Caption = lblChangeGet
lblShares(stockman) = shares
Write_Ini "Stock" & (stockman), "Shares", shares
lblPaidPts(stockman) = lblLastGet.Caption
Write_Ini "Stock" & (stockman), "Paid", lblLastGet.Caption
lblValue(stockman) = (shares * lblLastGet.Caption)
Write_Ini "Stock" & (stockman), "Value", (shares * lblLastGet.Caption)
lblValueChng(stockman) = lblValue(stockman).Caption - (Read_Ini("Stock" & stockman, "Shares") * lblPaidPts(stockman).Caption)
Budget = Budget - (shares * lblLastGet.Caption)
Write_Ini "Main", "Budget", Budget
End If
If YesNo = 1 Then
alreadyhave:
Write_Ini "Stock" & tester, "Shares", (Read_Ini("Stock" & tester, "Shares") + shares)
Write_Ini "Main", "Budget", Budget - (shares * lblLastGet.Caption)
Update
Write_Ini "Main", "Budget", Budget - 8
End If
Enderss:
End Sub

Private Sub cmdLoad_Click()
counters = 0
For counters = 0 To 6
If Read_Ini("Stock" & counters, "Symbol") <> "" Then
    cmdrGetQuote Read_Ini("Stock" & counters, "Symbol")
    Update
End If
Next
counters = 0
End Sub

Private Sub cmdSell_Click()
Dim many As Long
Dim yes As Long
Dim symbolforsale As String
symbolforsale = InputBox("What symbol would you like to sell?")
many = InputBox("How many shares of that stock would you like to sell? (Type all 9s to sell all shares)")
If many = 0 Then GoTo no
    Do Until counters > stock
            If lblSym(counters).Caption = symbolforsale Then
                yes = 1
                GoTo sell
            End If
        counters = counters + 1
    Loop
If yes = 1 Then
sell:
Write_Ini "Main", "Budget", lblPrice(counters).Caption * lblShares(counters).Caption + Budget
If many = Read_Ini("Stock" & counters, "Shares") Then
    DeleteSection "Stock" & counters
    GoTo no:
End If
Write_Ini "Stock" & counters, "Shares", ((Read_Ini("Stock" & counters, "Shares") - many))
'counter
Write_Ini "Main", "Budget", Budget + 8

Else
MsgBox "dint work"
End If



no:
Update
End Sub

Private Sub Command1_Click()
Write_Ini "Main", "Budget", "2500.00"
End Sub

Private Sub Form_Load()
Dim eR As EIGCInternetConnectionState
Dim sName As String
Dim bConnected As Boolean
frmMain.Show
INISetup "C:\STOCKS.INI", 1000
If Read_Ini("Main", "Budget") = "" Then
    Write_Ini "Main", "Budget", "2500"
End If
stockman = 0
    cboSymbol.Text = Read_Ini("Stock0", "Symbol")
    If cboSymbol.Text <> "" Then
        cmdrGetQuote (Read_Ini("Stock0", "Symbol"))
    Else
        cboSymbol.Text = "INTC"
        cmdrGetQuote "INTC"
    End If
    frmMain.Left = (Screen.Width - Me.Width) / 2
    frmMain.Top = (Screen.Height - Me.Height) / 2

   bConnected = InternetConnected(eR, sName)
   
   If (eR And INTERNET_CONNECTION_MODEM) = INTERNET_CONNECTION_MODEM Then
     lblConnection.Caption = lblConnection.Caption & "via modem." & vbCrLf
   End If
   
   If (eR And INTERNET_CONNECTION_LAN) = INTERNET_CONNECTION_LAN Then
     lblConnection.Caption = lblConnection.Caption & "via LAN." & vbCrLf
   End If
   
   If (eR And INTERNET_CONNECTION_PROXY) = INTERNET_CONNECTION_PROXY Then
     lblConnection.Caption = lblConnection.Caption & "via Proxy." & vbCrLf
   End If
   
   If (eR And INTERNET_CONNECTION_OFFLINE) = INTERNET_CONNECTION_OFFLINE Then
     lblConnection.Caption = "You are currently not connected to the internet." & vbCrLf
   End If
   
        lblSymbol.Caption = ""
        lblTime.Caption = Time
        lblDate.Caption = Date

        'Registry Load Stuff
        SaveSetting AppName:="Header", Section:="CustList", Key:="0", setting:=".."
        Mysettings1 = GetAllSettings(AppName:="Header", Section:="CustList")

    For intSettings1 = LBound(Mysettings1, 1) To UBound(Mysettings1, 1)
        cboSymbol.AddItem LTrim(Mysettings1(intSettings1, 1))
    Next intSettings1
    'cmdLoad_Click
End Sub

Private Sub cmdGetQuote_Click()
cmdrGetQuote (cboSymbol.Text)
End Sub
Public Function cmdrGetQuote(Simble)
On Error GoTo ErrorHandler
Dim not_first_symbolb As Boolean
Dim symbolb As String
Dim query_url As String
Dim i As Integer
Dim response As Variant
Dim objhttp As New MSXML.XMLHTTPRequest
Dim compname, datetime1, last1, open1, high1, low1, high52, changeper1, volume1, exchange, change1, marketcap, ask1, bid1, low52, peratio1, pershareprofit1, shareoutstanding1 As String
On Error GoTo ErrorHandler
        symb = Simble
        symbol = Simble
        
    'If cboSymbol.Text = "Symbol" Or Simble = ".." Then
    '    MsgBox "You must enter a ticker symbol first", vbOKOnly + vbCritical, "Duh!"
    '    cboSymbol.SetFocus
    '    Exit Function
    'End If

    If lblConnection.Caption = "You are currently not connected to the internet." Then
        MsgBox "You need to connect to the internet before you can use this program", _
        vbOKOnly + vbCritical, "Not Connected!"
    Exit Function
    End If
        objhttp.Open "GET", "http://www.stockpoint.com/quote.asp?Exchange=US&Symbol=" & symbol & "&Company=&x=0&y=0", False
        'objhttp.Open "GET", "http://finance.yahoo.com/q?s=" & symbol & "&d=t"
        'objhttp.Open "GET", "http://quotes.nasdaq-amex.com/quote.dll?page=multi&mode=stock&symbol=" & symbol
        objhttp.send
        strResponse = objhttp.responseText
                    
        'compname
        retval1 = InStr(1, strResponse, "<FONT COLOR=WHITE><B>") + 21
        retval2 = InStr(1, strResponse, "&nbsp;") - 1
        compname = Mid(strResponse, retval1, Len(strResponse) - retval2)
        retval1 = InStr(1, compname, "&nbsp;")
        compname = Left(compname, retval1 - 1)
        lblSymbol.Caption = compname & "(" & symbol & ")"
        compname = ""
        
        'date-time
        retval1 = InStr(1, strResponse, "As of ") + 6
        retval2 = InStr(1, strResponse, "(E.T.)") + 6
        datetime1 = Mid(strResponse, retval1, retval2 - retval1)
        temp = datetime1
        retval1 = InStr(1, datetime1, "&nbsp;&nbsp;") - 1
        datetime1 = Left(datetime1, retval1)
        retval1 = retval1 + 12
        datetime1 = datetime1 & " " & Right(temp, retval1)
        temp = Right(datetime1, 11)
        datetime1 = Left(datetime1, Len(datetime1) - 11)
        datetime1 = datetime1 & " " & Right(temp, 6)
        lblDateTime.Caption = datetime1
        datetime1 = ""

        'last
        'retval1 = InStr(1, strResponse, "Last") + 4
        'retval2 = InStr(retval1, strResponse, "<B>") + 3
        'strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        'retval1 = InStr(1, strResponse, "</B>") - 1
        'last1 = Left(strResponse, retval1)
        'lblLastGet.Caption = last1
        'last1 = ""



    txtQuotes.Text = ""
    DoEvents
    
    ' Prepare a URL to get the quotes.
    If USE_TEST_FILE Then
        response = LoadTestFile
    Else
        query_url = "http://quote.yahoo.com/q?s="
        For i = txtSymbol.LBound To txtSymbol.UBound    'LOOP
            symbolb = LCase$(Trim$(txtSymbol(i).Text))   'FORMAT CORRECT
            If Len(symbolb) > 0 Then                     'IS THERE A symbolb THERE?
                If not_first_symbolb Then _
                    query_url = query_url & "%2C"
                    query_url = query_url & symbolb
                    not_first_symbolb = True
            End If
        Next i
        query_url = query_url & "&d=v1"
        response = inetQuotes.OpenURL(query_url)
    End If
    Dim arse As String
    Dim sText() As String
    
    arse = ParseResponse(CStr(response))
    sText() = Split(arse, ", ", 500, vbTextCompare)
    Dim siblesa() As String
    siblesa() = Split(sText(0), ": ", 500, vbTextCompare)
    sText(0) = siblesa(1)
    'sText() = "Blah"
    lblLastGet.Caption = sText(1)
    'txtQuotes.Text = sText(0) & vbCrLf & sText(1) & vbCrLf & sText(2) & vbCrLf & sText(3) & vbCrLf & sText(4) & vbCrLf & sText(5) & vbCrLf
    'txtQuotes.Text = txtQuotes.Text & sText(0) & sText(1)
        'open
        retval1 = InStr(retval1, strResponse, "Open") + 4
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        open1 = Left(strResponse, retval1)
        Text1.Text = open1
        open1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblOpenGet.Caption = open1
          
        'high
        retval1 = InStr(retval1, strResponse, "high1") + 4
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        high1 = Left(strResponse, retval1)
        Text1.Text = high1
        high1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblHighGet.Caption = high1
        
        'low
        retval1 = InStr(retval1, strResponse, "low1") + 6
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        low1 = Left(strResponse, retval1)
        Text1.Text = low1
        low1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblLowGet.Caption = low1
           
        'volume
        retval1 = InStr(retval1, strResponse, "volume") + 6
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        volume1 = Left(strResponse, retval1)
        Text1.Text = volume1
        volume1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblVolumeGet.Caption = volume1
        
        'Per share earning
        retval1 = InStr(retval1, strResponse, "P/Share") + 7
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        pershareprofit1 = Left(strResponse, retval1)
        Text1.Text = pershareprofit1
        pershareprofit1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblPSEGet.Caption = pershareprofit1
                
        'per share outstanding
        retval1 = InStr(retval1, strResponse, "Outstanding") + 11
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        shareoutstanding1 = Left(strResponse, retval1)
        Text1.Text = shareoutstanding1
        shareoutstanding1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblPSOGet.Caption = shareoutstanding1
                
        'P/E ratio
        retval1 = InStr(retval1, strResponse, "Ratio") + 5
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        peratio1 = Left(strResponse, retval1)
        Text1.Text = peratio1
        peratio1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblPERGet.Caption = peratio1
        
        'change
        'retval1 = 0
        'retval1 = InStr(1, strResponse, "#339933")
        'If retval1 <> 0 Then
        'retval1 = retval1 + 9
        'ElseIf InStr(1, strResponse, "RED") <> 0 Then
        'retval1 = InStr(1, strResponse, "RED") + 6
        'ans = 1
        'Else
        'retval1 = InStr(1, strResponse, "SIZE=-1>") + 8
        'ans = 2
        'End If
        'strResponse = Mid(strResponse, retval1, Len(strResponse) - retval1)
        'retval2 = InStr(1, strResponse, "</FONT>")
        'change1 = Left(strResponse, retval2 - 1)
        'If ans = 2 Then change1 = Mid(change1, 3, Len(change1) - 3)
        'If ans = 1 Then change1 = "-" & change1
        'lblChangeGet.Caption = change1
        'Heh leaving this here for now... might need it later
        'and then my new one line way LOL!!
        lblChangeGet.Caption = sText(2)
        
        '% change
        'retval1 = 0
        'retval1 = InStr(1, strResponse, "#339933")
        'If retval1 <> 0 Then
        'retval1 = retval1 + 9
        'ElseIf InStr(1, strResponse, "RED") <> 0 Then
        'retval1 = InStr(1, strResponse, "RED") + 6
        'ans = 1
        'Else
        'retval1 = InStr(1, strResponse, "SIZE=-1>") + 8
        'ans = 2
        'End If
        'strResponse = Mid(strResponse, retval1, Len(strResponse) - retval1)
        'retval2 = InStr(1, strResponse, "</FONT>")
        'changeper1 = Left(strResponse, retval2 - 1)
        'Text1.Text = changeper1
        'If ans = 2 Then changeper1 = Mid(changeper1, 3, Len(changeper1) - 3)
        'If ans = 1 Then
        'changeper1 = "-" & changeper1
        'ans = 0
        'End If
        'lblChangepGet.Caption = changeper1
        'again my new one liner
        lblChangepGet.Caption = sText(3)
                
        '52high
        retval1 = InStr(retval1, strResponse, "High") + 4
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        high52 = Left(strResponse, retval1)
        Text1.Text = high52
        high52 = Right(Text1.Text, Len(Text1.Text) - 2)
        lbl52HighGet.Caption = high52
        
        '52low
        retval1 = InStr(retval1, strResponse, "Low") + 3
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        low52 = Left(strResponse, retval1)
        Text1.Text = low52
        low52 = Right(Text1.Text, Len(Text1.Text) - 2)
        lbl52LowGet.Caption = low52
        
        'bid
        retval1 = InStr(retval1, strResponse, "bid1") + 3
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        bid1 = Left(strResponse, retval1)
        Text1.Text = bid1
        bid1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblBidGet.Caption = bid1
        
        'ask
        retval1 = InStr(retval1, strResponse, "ask1") + 3
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        ask1 = Left(strResponse, retval1)
        Text1.Text = ask1
        ask1 = Right(Text1.Text, Len(Text1.Text) - 2)
        lblAskGet.Caption = ask1
        
        'yesterdays close
        retval1 = InStr(retval1, strResponse, "yesterday Close ") + 15
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        temp = Left(strResponse, retval1)
        Text1.Text = temp
        temp = Right(Text1.Text, Len(Text1.Text) - 2)
        
        'Market Cap
        retval1 = InStr(retval1, strResponse, "Capitalization") + 14
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        marketcap = Left(strResponse, retval1)
        Text1.Text = marketcap
        marketcap = Right(Text1.Text, Len(Text1.Text) - 2)
        lblMCGet.Caption = marketcap
        
        'exchange
        retval1 = InStr(retval1, strResponse, "Exchange") + 8
        retval2 = InStr(retval1, strResponse, "-1>") + 3
        strResponse = Mid(strResponse, retval2, Len(strResponse) - retval2)
        retval1 = InStr(1, strResponse, "</FONT>") - 1
        exchange = Left(strResponse, retval1)
        Text1.Text = exchange
        exchange = Right(Text1.Text, Len(Text1.Text) - 2)
        lblSEGet.Caption = exchange
        
        
        Dim url As String
        Dim Bite() As Byte
        
On Error GoTo ErrorHandler
Top:

symbol = LCase(cboSymbol)
        
    url = "http://a676.g.akamaitech.net/f/676/838/1h/nasdaq.com/logos/" & symbol & ".GIF"
        Bite() = Inet1.OpenURL(url, icByteArray) ' Download picture.s = Bilden()
        X = Bite()
        
    If Len(X) <> 75 Then
        Open "C:\graph.gif" For Binary Access Write As #1 ' Save the file.
        Put #1, , Bite()
        Close #1
    Else
    End If
    
        Picture1.Picture = LoadPicture("C:\graph.gif")

        Unload frmWait
        Exit Function
        
ErrorHandler:
        'MsgBox Err.Description & vbCrLf & Err.Number
        'Unload frmWait
        Resume Next
End Function

Public Sub GetGraph(GraphType As String, symbol As String)
Dim url As String
Dim Bite() As Byte
On Error GoTo ErrorHandler
Top:

symbol = LCase(cboSymbol)
        
    If cboGraph.Text = "1 year big" Then
        url = "http://chart.yahoo.com/c/1y/" & Left(symbol, 1) & "/" & symbol & ".gif"
    ElseIf cboGraph.Text = "2 year big" Then
        url = "http://chart.yahoo.com/c/2y/" & Left(symbol, 1) & "/" & symbol & ".gif"
    ElseIf cboGraph.Text = "3 months big" Then
        url = "http://chart.yahoo.com/c/3m/" & Left(symbol, 1) & "/" & symbol & ".gif"
    ElseIf cboGraph.Text = "6 months small" Then
        url = "http://chart.yahoo.com/c/0b/" & Left(symbol, 1) & "/" & symbol & ".gif"
    ElseIf cboGraph.Text = "1 Day" Then
        url = "http://ichart.yahoo.com/b?s=" & symbol
    ElseIf cobgraph.Text = "5 Day" Then
        url = "http://ichart.yahoo.com/w?s=" & symbol
    End If

        Bite() = Inet1.OpenURL(url, icByteArray) ' Download picture.s = Bilden()
        X = Bite()
        
    If Len(X) <> 75 Then
        Open "C:\graph.gif" For Binary Access Write As #1 ' Save the file.
        Put #1, , Bite()
        Close #1
    Else
    End If
    
        frmGraph.Picture1.Picture = LoadPicture("C:\graph.gif")
        frmGraph.Show
        
ErrorHandler:
    Resume Next
    If Err.Number = 35764 Then
        MsgBox "Still executiong last request", vbOKOnly, "Oops"
    GoTo Top
    End If
    
End Sub

Private Sub cmdViewGraph_Click()
    If cboSymbol.Text = "" Or cboSymbol = "Symbol" Or IsNumeric(cboSymbol.Text) = True Or cboGraph.Text = "" Then
        MsgBox "Could not process your request.  Check that you have entered a" & vbCrLf _
        & "valid ticker symbol and that you have selected a valid graph type.", vbOKOnly + vbCritical, "Oops"
        Exit Sub
    Else
        GetGraph cboGraph.Text, cboSymbol.Text
    End If

    
End Sub


Private Sub mnuExit_Click()
        Unload Me
End Sub
'Private Sub GraphDownloadCompleted(filename As String)
        'frmGraph.Picture1.Picture = LoadPicture("C:\graph.gif")
        'frmGraph.Show
'End Sub
Private Sub cmdAddSymbol_Click()
    If cboSymbol.Text = "Symbol" Or cboSymbol.Text = ".." Then
        MsgBox "You must enter a ticker symbol first", vbOKOnly + vbCritical, "Duh!"
        cboSymbol.SetFocus
    Exit Sub
    End If
        'Get Customer List form registry
        Mysettings1 = GetAllSettings(AppName:="Header", Section:="CustList")
        'Set the array upper and lower parameters
        
    For intSettings1 = LBound(Mysettings1, 1) To UBound(Mysettings1, 1)
        counter = Mysettings1(intSettings1, 0)
    Next intSettings1
    
        sCustKey = "CustList"
        sCustVal = RTrim(cboSymbol.Text)
        intSet1 = counter + 1
        'Saves combo text in CustList Registry
        SaveSetting AppName:="Header", Section:="CustList", Key:=intSet1, setting:=RTrim(sCustVal)
        Counter1 = Counter1 + 1
        cboSymbol.Clear
        'Clears then fills CboSymbol with new list
        Mysettings1 = GetAllSettings(AppName:="Header", Section:="CustList")
        
    For intSettings1 = LBound(Mysettings1, 1) To UBound(Mysettings1, 1)
        cboSymbol.AddItem Mysettings1(intSettings1, 1)
    Next intSettings1
    
        cboSymbol.ListIndex = 0
End Sub

Private Sub cmdRemoveSymbol_Click()

    If cboSymbol.Text = "Symbol" Or cboSymbol.Text = ".." Then
        MsgBox "You must enter a ticker symbol first", vbOKOnly + vbCritical, "Duh!"
        cboSymbol.SetFocus
    Exit Sub
    End If
    
    'Loop through the registry looking for a match to the cboSymbol.Text
    'If it is found delete it from the registry
    
    For intSettings1 = LBound(Mysettings1, 1) To UBound(Mysettings1, 1)
        j = Mysettings1(intSettings1, 1)
        j1 = Mysettings1(intSettings1, 0)
        
    If j = cboSymbol.Text Then GoTo Del
    
    Next intSettings1

Del:
        DeleteSetting "Header", "CustList", j1
        cboSymbol.Clear 'Clear before re-loading combo with new values
        Mysettings1 = GetAllSettings(AppName:="Header", Section:="CustList")
        'Re-Read the registry and fill cboSymbol
        
   For intSettings1 = LBound(Mysettings1, 1) To UBound(Mysettings1, 1)
        cboSymbol.AddItem LTrim(Mysettings1(intSettings1, 1))
   Next intSettings1
   
        cboSymbol.ListIndex = 0
    
End Sub

Private Sub Timer1_Timer()
Dim blah As Long
Dim crap(6) As Long
Budget = Read_Ini("Main", "Budget")
    crap(0) = Val(lblValue(0).Caption)
    crap(1) = Val(lblValue(1).Caption)
    crap(2) = Val(lblValue(2).Caption)
    crap(3) = Val(lblValue(3).Caption)
    crap(4) = Val(lblValue(4).Caption)
    crap(5) = Val(lblValue(5).Caption)
    crap(6) = Val(lblValue(6).Caption)
    budgeta = (crap(0) + crap(1) + crap(2) + crap(3) + crap(4) + crap(5) + crap(6)) + Budget
Label11.Caption = "$" & budgeta
lblBudget.Caption = "$" & Budget
End Sub

Private Sub Timer2_Timer()
'***********************   This Timer finds the interger stock   ****************************
If lblSym(0).Caption = "" Then
    stockman = 0
    GoTo nospace
ElseIf lblSym(1).Caption = "" Then
    stockman = 1
    GoTo nospace
ElseIf lblSym(2).Caption = "" Then
    stockman = 2
    GoTo nospace
ElseIf lblSym(3).Caption = "" Then
    stockman = 3
    GoTo nospace
ElseIf lblSym(4).Caption = "" Then
    stockman = 4
    GoTo nospace
ElseIf lblSym(5).Caption = "" Then
    stockman = 5
    GoTo nospace
ElseIf lblSym(6).Caption = "" Then
    stockman = 6
    GoTo nospace
End If
nospace:
'Labela.Caption = stockman
End Sub

Private Sub Timer3_Timer()
cmdGetQuote_Click
End Sub
'******************************************************'
'                                                      '
'                                                      '
'                     Stuff                            '
'                                                      '
'                                                      '
'******************************************************'


' Get a row from the response string.
Private Function GetRow(response As String) As String
Dim pos As Integer
Dim symbolb As String
Dim last_time As String
Dim last_price As String
Dim change_amount As String
Dim change_percent As String

    ' Find the "<tr" starting the row.
    pos = InStr(response, "<tr")
    If pos = 0 Then
        response = ""
        GetRow = ""
        Exit Function
    End If
    response = Mid$(response, pos)
    symbolb = GetRowItem(response)
    last_time = GetRowItem(response)
    If InStr(last_time, "No such ticker symbol.") > 0 Then
        GetRow = "No such ticker symbol."
        Exit Function
    End If
    last_price = GetRowItem(response)
    change_amount = GetRowItem(response)
    change_percent = GetRowItem(response)
    GetRow = symbolb & ": " & _
        last_time & ", " & _
        last_price & ", " & _
        change_amount & ", " & _
        change_percent
End Function
' Get the next table item from the table.
Private Function GetRowItem(response As String) As String
Dim start_pos As Integer
Dim end_pos As Integer
Dim pos As Integer
Dim count As Integer
Dim ch As String
Dim txt As String

    ' Find the "<td" and "</td" that bracket
    ' the item.
    start_pos = InStr(response, "<td")
    end_pos = InStr(start_pos, response, "</td")

    ' Save characters between these where the
    ' outstanding brackets match.
    count = 1
    For pos = start_pos + 1 To end_pos
        ch = Mid$(response, pos, 1)
        If ch = "<" Then
            count = count + 1
        ElseIf ch = ">" Then
            count = count - 1
        Else
            If count = 0 Then txt = txt & ch
        End If
    Next pos

    GetRowItem = txt
    response = Mid$(response, end_pos)
End Function
Private Function LoadTestFile() As String
Dim fname As String
Dim fnum As Integer

    fname = App.Path & "\quote.htm"
    fnum = FreeFile
    Open fname For Input As fnum
    LoadTestFile = Input$(LOF(fnum), #fnum)
    Close fnum
End Function

Private Function ParseResponse(ByVal response As String) As String
Dim start_pos As Integer
Dim end_pos As Integer
Dim i As Integer
Dim quotes As String
Dim new_row As String

    ' Find the table that contains the
    ' interesting information.
    start_pos = InStr(response, "Last Trade")
    If start_pos = 0 Then
        ParseResponse = "Error parsing response."
        Exit Function
    End If
    
    ' See where the table ends.
    end_pos = InStr(start_pos, response, "</table>")
    response = Mid$(response, start_pos, end_pos - start_pos + Len("</table>"))

    ' Parse the rows from the table.
    Do
        new_row = GetRow(response)
        If Len(new_row) = 0 Then Exit Do
        
        quotes = quotes & new_row & vbCrLf
    Loop

    ParseResponse = quotes
End Function
Private Sub cmdGetQuotes_Click()
Dim not_first_symbolb As Boolean
Dim symbolb As String
Dim query_url As String
Dim i As Integer
Dim response As Variant
    txtQuotes.Text = ""
    DoEvents
    
    ' Prepare a URL to get the quotes.
    If USE_TEST_FILE Then
        response = LoadTestFile
    Else
        query_url = "http://quote.yahoo.com/q?s="
        For i = txtSymbol.LBound To txtSymbol.UBound    'LOOP
            symbolb = LCase$(Trim$(txtSymbol(i).Text))   'FORMAT CORRECT
            If Len(symbolb) > 0 Then                     'IS THERE A symbolb THERE?
                If not_first_symbolb Then _
                    query_url = query_url & "%2C"
                    query_url = query_url & symbolb
                    not_first_symbolb = True
            End If
        Next i
        query_url = query_url & "&d=v1"
        response = inetQuotes.OpenURL(query_url)
    End If
    txtQuotes.Text = ParseResponse(CStr(response))
End Sub

