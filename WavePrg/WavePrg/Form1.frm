VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WaveForm"
   ClientHeight    =   9180
   ClientLeft      =   495
   ClientTop       =   4620
   ClientWidth     =   16680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   16680
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   16635
      TabIndex        =   19
      Top             =   990
      Width           =   16665
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000012&
         ForeColor       =   &H000000FF&
         Height          =   1170
         Left            =   1215
         ScaleHeight     =   74
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1022
         TabIndex        =   37
         Top             =   315
         Width           =   15390
         Begin VB.Timer TimerstopPlay 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   630
            Top             =   0
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   45
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000C000&
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   135
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   1020
            X2              =   1020
            Y1              =   0
            Y2              =   135
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1050
         TabIndex        =   36
         Text            =   "?"
         Top             =   45
         Width           =   420
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2220
         TabIndex        =   35
         Text            =   "?"
         Top             =   45
         Width           =   915
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3960
         TabIndex        =   34
         Text            =   "?"
         Top             =   45
         Width           =   555
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5445
         TabIndex        =   33
         Text            =   "?"
         Top             =   45
         Width           =   330
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6330
         TabIndex        =   32
         Text            =   "?"
         Top             =   45
         Width           =   285
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7320
         TabIndex        =   31
         Text            =   "?"
         Top             =   45
         Width           =   195
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0E0FF&
         Height          =   240
         Left            =   8205
         TabIndex        =   30
         Text            =   "?"
         Top             =   45
         Width           =   195
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   9465
         TabIndex        =   29
         Text            =   "?"
         Top             =   45
         Width           =   600
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10950
         TabIndex        =   28
         Text            =   "?"
         Top             =   45
         Width           =   600
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   12600
         TabIndex        =   27
         Text            =   "?"
         Top             =   45
         Width           =   195
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0E0FF&
         Height          =   240
         Left            =   13815
         TabIndex        =   26
         Text            =   "?"
         Top             =   45
         Width           =   285
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   14565
         TabIndex        =   25
         Text            =   "?"
         Top             =   45
         Width           =   420
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   15705
         TabIndex        =   24
         Text            =   "?"
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton CommandsaveL 
         Caption         =   "saveL(cu)"
         Height          =   255
         Left            =   45
         TabIndex        =   23
         ToolTipText     =   "save CuData to file (after preemphasis and set 0)"
         Top             =   1215
         Width           =   1095
      End
      Begin VB.CommandButton CommandplayL 
         Caption         =   "playL(cu)"
         Height          =   255
         Left            =   45
         TabIndex        =   22
         ToolTipText     =   "play CuData (after preemphasis and set 0)"
         Top             =   945
         Width           =   1095
      End
      Begin VB.CommandButton CommandTransfer 
         Caption         =   "transfer"
         Height          =   255
         Left            =   45
         TabIndex        =   21
         Top             =   585
         Width           =   1095
      End
      Begin VB.CommandButton Commandopen 
         Caption         =   "open"
         Height          =   255
         Left            =   45
         TabIndex        =   20
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "wRiffFormatTag"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   50
         Top             =   45
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "wfdataSize "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1575
         TabIndex        =   49
         Top             =   45
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "wFormatTag"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3240
         TabIndex        =   48
         Top             =   45
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "wFormatName"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4635
         TabIndex        =   47
         Top             =   45
         Width           =   780
      End
      Begin VB.Label Label5 
         Caption         =   "wCsize"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5895
         TabIndex        =   46
         Top             =   45
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "wWavefmt"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6705
         TabIndex        =   45
         Top             =   45
         Width           =   555
      End
      Begin VB.Label Label7 
         Caption         =   "wChannels"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7605
         TabIndex        =   44
         Top             =   45
         Width           =   600
      End
      Begin VB.Label Label8 
         Caption         =   "wSamplesPerSec "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8505
         TabIndex        =   43
         Top             =   45
         Width           =   915
      End
      Begin VB.Label Label9 
         Caption         =   "wData"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   14175
         TabIndex        =   42
         Top             =   45
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "wBitsPerSample"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12870
         TabIndex        =   41
         Top             =   45
         Width           =   870
      End
      Begin VB.Label Label11 
         Caption         =   "wBytePerSample"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11655
         TabIndex        =   40
         Top             =   45
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "wBytePerSec"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10170
         TabIndex        =   39
         Top             =   45
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "wDataSize"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   15075
         TabIndex        =   38
         Top             =   45
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   0
      ScaleHeight     =   2775
      ScaleWidth      =   16635
      TabIndex        =   18
      Top             =   2520
      Width           =   16665
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   2685
         Left            =   1215
         ScaleHeight     =   175
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1021
         TabIndex        =   51
         Top             =   45
         Width           =   15375
      End
      Begin VB.OptionButton PIC2Option 
         Caption         =   "WaveForm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   57
         Top             =   1755
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton PIC2Option 
         Caption         =   "FFT"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   56
         Top             =   810
         Width           =   615
      End
      Begin VB.OptionButton PIC2Option 
         Caption         =   "Spectrum"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   55
         Top             =   1035
         Width           =   1095
      End
      Begin VB.CheckBox CheckHaming 
         Caption         =   "haming"
         Height          =   255
         Left            =   195
         TabIndex        =   54
         Top             =   1290
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Checkbefore 
         Caption         =   "preemphasis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   53
         Top             =   90
         Value           =   1  'Checked
         Width           =   1260
      End
      Begin VB.CheckBox Checkset0 
         Caption         =   "set 0"
         Height          =   255
         Left            =   45
         TabIndex        =   52
         Top             =   555
         Value           =   1  'Checked
         Width           =   720
      End
      Begin VB.Label TextPreEmphasis 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0.95"
         Height          =   195
         Left            =   585
         TabIndex        =   72
         Top             =   315
         Width           =   495
      End
      Begin VB.Label LabelFFTsize 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1024"
         Height          =   195
         Left            =   810
         TabIndex        =   58
         Top             =   1530
         Width           =   405
      End
      Begin VB.Label Label18 
         Caption         =   "FFTsize"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   59
         Top             =   1530
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2505
      ScaleWidth      =   16635
      TabIndex        =   17
      Top             =   5310
      Width           =   16665
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   2430
         Left            =   1215
         ScaleHeight     =   2370
         ScaleWidth      =   15330
         TabIndex        =   60
         Top             =   45
         Width           =   15390
      End
      Begin VB.CheckBox CheckShortTimeAvecrossPer 
         Caption         =   "STAcross.01"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   240
         Left            =   45
         TabIndex        =   64
         Tag             =   ".01"
         Top             =   1530
         Width           =   1200
      End
      Begin VB.CheckBox CheckShortTimeAveEnergy 
         Caption         =   "STAEnergy"
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   45
         TabIndex        =   63
         Top             =   1305
         Width           =   1200
      End
      Begin VB.CheckBox CheckShortTimeAveCross0 
         Caption         =   "STACross0"
         ForeColor       =   &H0000C0C0&
         Height          =   240
         Left            =   45
         TabIndex        =   62
         Top             =   1080
         Width           =   1200
      End
      Begin VB.CheckBox CheckShortTimeAveScope 
         Caption         =   "STAScope"
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   45
         TabIndex        =   61
         Top             =   855
         Width           =   1200
      End
      Begin VB.Label TextshortTime2MS 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   71
         Top             =   540
         Width           =   585
      End
      Begin VB.Label TextshortTimeN2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "256p"
         Height          =   195
         Left            =   45
         TabIndex        =   70
         Top             =   540
         Width           =   495
      End
      Begin VB.Label LabelshortTime 
         Caption         =   "ShortTime"
         Height          =   195
         Left            =   45
         TabIndex        =   67
         Top             =   90
         Width           =   735
      End
      Begin VB.Label TextshortTimeN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "256p"
         Height          =   195
         Left            =   45
         TabIndex        =   66
         Top             =   315
         Width           =   495
      End
      Begin VB.Label TextshortTimeMS 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   65
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   16635
      TabIndex        =   1
      Top             =   0
      Width           =   16665
      Begin VB.PictureBox Picture9 
         Height          =   885
         Left            =   11970
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   825
         ScaleWidth      =   315
         TabIndex        =   68
         Top             =   45
         Width           =   375
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   800
            Left            =   0
            ScaleHeight     =   795
            ScaleWidth      =   330
            TabIndex        =   69
            Top             =   0
            Width           =   330
         End
      End
      Begin VB.CommandButton CommandRecorD 
         Caption         =   "RecordDown"
         Enabled         =   0   'False
         Height          =   255
         Left            =   45
         TabIndex        =   5
         Top             =   405
         Width           =   1095
      End
      Begin VB.CheckBox Checknew 
         Caption         =   "new"
         Height          =   255
         Left            =   540
         TabIndex        =   11
         Top             =   675
         Width           =   615
      End
      Begin VB.CommandButton CommandView 
         Caption         =   "View"
         Height          =   255
         Left            =   45
         TabIndex        =   10
         Top             =   45
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   13725
         TabIndex        =   9
         Text            =   "16"
         Top             =   45
         Width           =   735
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   13725
         TabIndex        =   8
         Text            =   "11025"
         Top             =   337
         Width           =   735
      End
      Begin VB.CommandButton CommandVolume 
         Caption         =   "Volume"
         Height          =   255
         Left            =   15615
         TabIndex        =   7
         Top             =   675
         Width           =   960
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   15840
         TabIndex        =   6
         Text            =   "1024"
         Top             =   45
         Width           =   735
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   15840
         TabIndex        =   4
         Text            =   "100"
         Top             =   360
         Width           =   735
      End
      Begin VB.PictureBox scopeBox 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H00C0FFC0&
         Height          =   885
         Left            =   1215
         ScaleHeight     =   825
         ScaleWidth      =   10680
         TabIndex        =   3
         Top             =   45
         Width           =   10740
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Left            =   90
            Top             =   90
         End
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   13725
         TabIndex        =   2
         Text            =   "2"
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "wBitsPerSample"
         Height          =   270
         Left            =   12420
         TabIndex        =   16
         Top             =   45
         Width           =   1275
      End
      Begin VB.Label Label15 
         Caption         =   "wSamplesPerSec "
         Height          =   270
         Left            =   12420
         TabIndex        =   15
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label16 
         Caption         =   "BUsampleNumber"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   14535
         TabIndex        =   14
         Top             =   60
         Width           =   1275
      End
      Begin VB.Label Label17 
         Caption         =   "Timer Interval "
         Height          =   270
         Left            =   14535
         TabIndex        =   13
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label21 
         Caption         =   "NUM_BUFFERS"
         Height          =   270
         Left            =   12420
         TabIndex        =   12
         Top             =   615
         Width           =   1275
      End
   End
   Begin VB.TextBox Text18 
      Height          =   750
      Left            =   1665
      TabIndex        =   0
      Text            =   "Text18"
      Top             =   8505
      Width           =   1140
   End
   Begin VB.PictureBox Picture10 
      Height          =   7350
      Left            =   1260
      ScaleHeight     =   7290
      ScaleWidth      =   15300
      TabIndex        =   73
      Tag             =   "1"
      Top             =   1035
      Width           =   15360
      Begin VB.TextBox TextHn 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   345
         TabIndex        =   75
         Text            =   "1500"
         Top             =   3015
         Width           =   495
      End
      Begin VB.TextBox TextVn 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   345
         TabIndex        =   74
         Text            =   "300"
         Top             =   2700
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Hn"
         Height          =   255
         Left            =   60
         TabIndex        =   77
         Top             =   3015
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "Vn"
         Height          =   285
         Left            =   75
         TabIndex        =   76
         Top             =   2700
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ReadWaveFile(sFileName As String) As Byte
Dim i As Long
'ReadWaveFile = -1                    lio10   LIO10
Open sFileName For Binary As #1 '開啟來源檔
Get #1, 1, pHandle '讀出44Byte的檔頭
        Text1.Text = pHandle.wRiffFormatTag
        Text2.Text = pHandle.wfdataSize
        Text3.Text = pHandle.wFormatTag
        Text4.Text = pHandle.wFormatName
        Text5.Text = pHandle.wCsize
        Text6.Text = pHandle.wWavefmt
        Text7.Text = pHandle.wChannels
        Text8.Text = pHandle.wSamplesPerSec
        Text9.Text = pHandle.wBytePerSec
        Text10.Text = pHandle.wBytePerSample
        Text11.Text = pHandle.wBitsPerSample
        Text12.Text = pHandle.wData
        Text13.Text = pHandle.wDataSize
        
        Picture1.Scale (0, 1)-(1, 0)
        Picture1.Cls
        Picture1.CurrentX = 0
        Picture1.CurrentY = 1
        If pHandle.wSamplesPerSec * pHandle.wBytePerSample <> pHandle.wBytePerSec Then Picture1.Print "pHandle.wSamplesPerSec * pHandle.wBytePerSample <> pHandle.wBytePerSec"
        'Debug.Print Picture1.CurrentX
        'Debug.Print Picture1.CurrentY
        If pHandle.wBitsPerSample / 8 * pHandle.wChannels <> pHandle.wBytePerSample Then Picture1.Print "pHandle.wBitsPerSample/8*pHandle.wChannels<>pHandle.wBytePerSample"
        If pHandle.wDataSize + 36 <> pHandle.wfdataSize Then Picture1.Print "pHandle.wDataSize + 36 <> pHandle.wfdataSize"
        If LOF(1) - pHandle.wfdataSize <> 8 Then Picture1.Print "LOF(1) - pHandle.wfdataSize <> 8"
'If MsgBox("ok?", vbYesNo, "pHandle") <> vbYes Then Close: Exit Sub

'ReDim LongData(pHandle.wDataSize / (pHandle.wBitsPerSample) / 8 / pHandle.wChannels)

Dim N As Long, BB1 As Byte, BB2 As Byte, II1 As Integer, II2 As Integer
N = pHandle.wDataSize / (pHandle.wBitsPerSample / 8) / pHandle.wChannels - 1
ReDim LongData(N)
If pHandle.wChannels = 1 Then '這是8bit音檔
    If pHandle.wBitsPerSample = 8 Then
        For i = 0 To N
             Get #1, , BB1
             LongData(i) = BB1
             'Debug.Print LongData(i)
             'If LongData(i) < 0 Then MsgBox ""
        Next i
    ElseIf pHandle.wBitsPerSample = 16 Then
        'For i = 0 To N
             Get #1, , LongData
             'LongData(i) = II1
             'Debug.Print LongData(i)
        'Next i
    Else
            Close
            MsgBox "BitsPerSample<>8 or 16"
            Picture1.Print "BitsPerSample<>8 or 16"
            Exit Function
    End If
ElseIf pHandle.wChannels = 2 Then
    Dim K As Integer
kk:      K = InputBox("1-left,0-right,0.5-lr/2", "Channels = 2", 0.5)
    If K <> 0 And K <> 1 And K <> 0.5 Then GoTo kk
    If pHandle.wBitsPerSample = 8 Then '單聲道
        For i = 0 To N
             Get #1, , BB1
             Get #1, , BB2
             LongData(i) = BB1 * K + BB2 * (1 - K)
             'Debug.Print LongData(i)
        Next i
    ElseIf pHandle.wBitsPerSample = 16 Then '雙聲道
        For i = 0 To N
             Get #1, , II1
             Get #1, , II2
             LongData(i) = II1 * K + II2 * (1 - K)
             'Debug.Print LongData(i)
        Next i
    Else
            Close
            MsgBox "BitsPerSample<>8 or 16"
            Picture1.Print "BitsPerSample<>8 or 16"
            Exit Function
    End If
Else
    Close
    MsgBox "pHandle.wChannels<>1 or 2"
    Picture1.Print "pHandle.wChannels<>1 or 2"
    Exit Function
End If
Close
ReadWaveFile = 1
End Function
'iol10
'IOL10

Public Sub PaintWave()
        On Error Resume Next
        Dim N As Long, i As Long
        N = pHandle.wDataSize / (pHandle.wBitsPerSample / 8) / pHandle.wChannels - 1
        If N < 1 Then N = 1
        Picture1.Cls
        If pHandle.wBitsPerSample = 8 Then Picture1.Scale (0, 256)-(N, 0)    '設定繪圖區域座標
        If pHandle.wBitsPerSample = 16 Then Picture1.Scale (0, 32767)-(N, -32768)
                    For i = 0 To N 'Step Int(N / 629) + 1  'save time
                        Picture1.Line -(i, LongData(i)), QBColor(8)                                '+ Int(N / 629) * Rnd
                        'Picture1.PSet (I, LongData(I))
                    Next i
        If pHandle.wBitsPerSample = 8 Then Picture1.CurrentY = 256      '設定繪圖區域座標
        If pHandle.wBitsPerSample = 16 Then Picture1.CurrentY = 32767
        Picture1.CurrentX = 0
        Picture1.Print "N="; UBound(LongData); Int(1 / CSng(pHandle.wSamplesPerSec) * CSng(UBound(LongData)) * 1000000) / 1000; "ms "; " STEP="; Int(UBound(LongData) / 2000) + 1;
End Sub
Private Sub makeCu()
        makeCuM pHandle, LongData, startX, endX, Form1.Checkbefore.Value, Form1.TextPreEmphasis, Form1.Checkset0.Value
        TextshortTimeN_Change
        TextshortTimeN2_Change
        'Form1.Text8.Text wSamplesPerSecP
        'Form1.Checkbefore.Value PreEmphasisYesNo  －0 是没有检查（缺省值），1 为已检查，和 2 为变灰（变
        'Form1.TextPreEmphasis PreEmphasisK
        'Form1.Checkset0.Value set0YesNo

End Sub
Private Sub makeCu_old()  ' 放弃 已用waveCls的makeCuM 替代
        On Error Resume Next
        CuHandle = pHandle
        
        'CuHandle.wRiffFormatTag
        'CuHandle.wFormatTag
        'CuHandle.wFormatName
        'CuHandle.wCsize
        'CuHandle.wWavefmt
        CuHandle.wChannels = 1
            CuHandle.wSamplesPerSec = Text8.Text  'ccccccccccccccccccccccccccccccccccccccchang
        CuHandle.wBytePerSample = pHandle.wBitsPerSample / 8 * CuHandle.wChannels
        CuHandle.wBytePerSec = CuHandle.wSamplesPerSec * CuHandle.wBytePerSample
        'CuHandle.wBitsPerSample
        'CuHandle.wData
        CuHandle.wDataSize = (Abs(startX - endX) + 1) * pHandle.wBitsPerSample / 8
        CuHandle.wfdataSize = CuHandle.wDataSize + 36
        Dim i As Long, tyt As Long
        ReDim CuDataS(Abs(startX - endX))
        'ReDim CuDataL(UBound(LongData))
        
        tyt = endX - startX  '避免 startX - endX ＝0 死循环
        If tyt = 0 Then tyt = 1
        
        For i = startX To endX Step Sgn(tyt) '预加重
             If Checkbefore.Value = 1 Then
                 If pHandle.wBitsPerSample = 8 Then
                    CuDataS(Abs(i - startX)) = (LongData(i + 1) - 128) - TextPreEmphasis * (LongData(i) - 128) + 128
                 Else
                    CuDataS(Abs(i - startX)) = LongData(i + 1) - TextPreEmphasis * LongData(i)
                 End If
             Else
                CuDataS(Abs(i - startX)) = LongData(i)
             End If
        Next i
        If Checkset0.Value = 1 Then
                Dim set0HA As Double
                set0HA = 0
                For i = 0 To Abs(startX - endX)    'set 0
                        set0HA = set0HA + CuDataS(i)
                Next i
                set0HA = set0HA / (Abs(startX - endX) + 1)
                If pHandle.wBitsPerSample = 8 Then set0HA = set0HA - 128
                For i = 0 To Abs(startX - endX)  'set 0
                        CuDataS(i) = CuDataS(i) - set0HA
                Next i
        End If
        
        
'        For i = 0 To UBound(LongData)   '预加重       CuDataL
'            If Checkbefore.Value = 1 Then
'                 If pHandle.wBitsPerSample = 8 Then
'                    CuDataL(i) = (LongData(i + 1) - 128) - TextPreEmphasis * (LongData(i) - 128) + 128
'                 Else
'                    CuDataL(i) = LongData(i + 1) - TextPreEmphasis * LongData(i)
'                 End If
'             Else
'               CuDataL(i) = LongData(i)
'             End If
'        Next i
'        If Checkset0.Value = 1 Then
'                Dim set0HA As Double
'                set0HA = 0
'                For i = 0 To UBound(LongData)    'set 0
'                        set0HA = set0HA + CuDataL(i)
'                Next i
'                set0HA = set0HA / (UBound(LongData) + 1)
'                If pHandle.wBitsPerSample = 8 Then set0HA = set0HA - 128
'                For i = 0 To Abs(startX - endX)  'set 0
'                        CuDataL(i) = CuDataL(i) - set0HA
'                Next i
'        End If
End Sub

Private Sub Checkbefore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    makeCu
    PIC2Option_Click (9)
End Sub
Private Sub Checkset0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    makeCu
    PIC2Option_Click (9)
End Sub

Private Sub CheckHaming_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PIC2Option_Click (9)
End Sub

Private Sub CheckShortTimeAveCross0_Click()
    'TextshortTimeN_Change
     Picture3Draw
End Sub

Private Sub CheckShortTimeAvecrossPer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbRightButton Then
                If Shift = vbShiftMask Then
                        CheckShortTimeAvecrossPer.Tag = CheckShortTimeAvecrossPer.Tag - 0.01
                ElseIf Shift = vbCtrlMask Then
                        CheckShortTimeAvecrossPer.Tag = CheckShortTimeAvecrossPer.Tag + 0.1
                Else
                        CheckShortTimeAvecrossPer.Tag = CheckShortTimeAvecrossPer.Tag + 0.01
                End If
                If Val(CheckShortTimeAvecrossPer.Tag) < 0# Then CheckShortTimeAvecrossPer.Tag = 0.4
                If Val(CheckShortTimeAvecrossPer.Tag) > 0.4 Then CheckShortTimeAvecrossPer.Tag = 0#
                CheckShortTimeAvecrossPer.Caption = "STAcross" & CheckShortTimeAvecrossPer.Tag
        ElseIf Button = vbLeftButton Then
                CheckShortTimeAvecrossPer.Tag = 0.01
        End If
        Picture3Draw
End Sub

Private Sub CheckShortTimeAvecrossPer_Click()
    'TextshortTimeN_Change
    ' Picture3Draw
End Sub

Private Sub CheckShortTimeAveScope_Click()
    'TextshortTimeN_Change
     Picture3Draw
End Sub
Private Sub CheckShortTimeAveEnergy_Click()
    'TextshortTimeN_Change
     Picture3Draw
End Sub




Private Sub CommandRecorD_Click()
        If Recording = False Then
                CommandRecorD.Caption = "recording..."
                If Checknew.Value = 1 Then ReDim LongData(0) Else ReDim Preserve LongData(UBound(LongData))
                Recording = True
                RecordLongdataTOhandale
        Else
                Recording = False
                CommandRecorD.Caption = "RecorDown"
                    RecordLongdataTOhandale
                    
                    PaintWave
                        startX = Picture1.ScaleLeft
                        Line1.X1 = startX
                        Line1.X2 = startX
                        endX = Picture1.ScaleWidth + Picture1.ScaleLeft - 0
                        Line2.X1 = endX
                        Line2.X2 = endX
                        Form1.Caption = "Recorded wave"
                    makeCu
                    PIC2Option_Click (9)
        End If
End Sub
Public Sub RecordLongdataTOhandale()
                    pHandle.wRiffFormatTag = "RIFF"
                    pHandle.wFormatTag = "WAVE"
                    pHandle.wFormatName = "fmt "
                    pHandle.wData = "data"
                    pHandle.wCsize = 16
                    pHandle.wWavefmt = 1
                    pHandle.wChannels = 1
                    pHandle.wSamplesPerSec = Text15.Text
                    pHandle.wBitsPerSample = Text14.Text
                    pHandle.wBytePerSample = pHandle.wBitsPerSample / 8 * pHandle.wChannels
                    pHandle.wBytePerSec = pHandle.wBytePerSample * pHandle.wSamplesPerSec
                    pHandle.wDataSize = (UBound(LongData) + 1) * pHandle.wChannels * (pHandle.wBitsPerSample / 8)
                    pHandle.wfdataSize = pHandle.wDataSize + 36

                    Text1.Text = pHandle.wRiffFormatTag
                    Text2.Text = pHandle.wfdataSize
                    Text3.Text = pHandle.wFormatTag
                    Text4.Text = pHandle.wFormatName
                    Text5.Text = pHandle.wCsize
                    Text6.Text = pHandle.wWavefmt
                    Text7.Text = pHandle.wChannels
                    Text8.Text = pHandle.wSamplesPerSec
                    Text9.Text = pHandle.wBytePerSec
                    Text10.Text = pHandle.wBytePerSample
                    Text11.Text = pHandle.wBitsPerSample
                    Text12.Text = pHandle.wData
                    Text13.Text = pHandle.wDataSize
End Sub
Public Sub ReadRecorD()     '由录音数据写phandle

                '    RecordLongdataTOhandale
                '    PaintWave
                    
End Sub

Private Sub CommandsaveL_Click()
    makeCu
    WriteWave CuHandle, CuDataS
End Sub
Private Sub WriteWave(aHandle As PCMFORM, Arr)
    CommonDialog1.Filter = "声音文件 (*.wav)|*.wav"
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNOverwritePrompt
    On Error GoTo Errhandler
    Randomize
    CommonDialog1.FileName = aHandle.wChannels & "_" & aHandle.wBitsPerSample & "_" & aHandle.wSamplesPerSec & "_" & Int(Rnd * 1000)
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
    
    Dim LorI As Variant
    Open CommonDialog1.FileName For Binary As #2  '開啟來源檔
    Put #2, , aHandle '讀出44Byte的檔頭
    For Each LorI In Arr
        If aHandle.wBitsPerSample = 16 Then
            Put #2, , CInt(LorI)
        ElseIf aHandle.wBitsPerSample = 8 Then
            Put #2, , CByte(LorI)
        End If
    Next
    Close
Errhandler:
    Exit Sub
            
End Sub

   
Private Sub CommandTransfer_Click()     'y
        If pHandle.wChannels = 2 Then
                    pHandle.wChannels = 1
        End If
        
        Dim ta As Single, II As Long
        If pHandle.wBitsPerSample <> Text11.Text Then
            ta = 1
            If pHandle.wBitsPerSample = 8 And Text11.Text = 16 Then
                    ta = 256
                    pHandle.wBitsPerSample = Text11.Text
                    For II = 0 To UBound(LongData)
                        LongData(II) = (LongData(II) - 128) * 255    'no 256 fangzhi yichu
                    Next II
            End If
            If pHandle.wBitsPerSample = 16 And Text11.Text = 8 Then
                    ta = 1 / 256
                    pHandle.wBitsPerSample = Text11.Text
                    For II = 0 To UBound(LongData)
                        LongData(II) = LongData(II) / 256 + 128
                    Next II
            End If
        End If
        
    If pHandle.wSamplesPerSec <> Text8.Text Then
                Dim Nn1 As Long, Nn2 As Long, Xaa() As Double, Yaa() As Double, Yee() As Double, Ybb() As Double
                Nn1 = UBound(LongData)
                Nn2 = UBound(LongData) * CSng(Val(Text8.Text)) / CSng(pHandle.wSamplesPerSec)
                ReDim Xaa(Nn1), Yaa(Nn1), Yee(Nn1), Ybb(Nn2)
                For II = 0 To Nn1
                    Xaa(II) = II
                    Yaa(II) = LongData(II)
                Next II
                Call SPLINE(Xaa(), Yaa(), Nn1, 0, 0, Yee())
                ReDim LongData(Nn2)
                For II = 0 To Nn2
                        Call SPLINT(Xaa(), Yaa(), Yee(), Nn1, II * CSng(pHandle.wSamplesPerSec) / CSng(Val(Text8.Text)), Ybb(II))
                        If Ybb(II) > 32767 Then Ybb(II) = 32767
                        If Ybb(II) < -32768 Then Ybb(II) = -32768
                        LongData(II) = CInt(Ybb(II))
                Next II
                
                pHandle.wSamplesPerSec = Text8.Text
    End If
        
        
                    pHandle.wRiffFormatTag = "RIFF"
                    pHandle.wFormatTag = "WAVE"
                    pHandle.wFormatName = "fmt "
                    pHandle.wData = "data"
                    pHandle.wCsize = 16
                    pHandle.wWavefmt = 1
            'pHandle.wChannels = 1
            'pHandle.wBitsPerSample = Text11.Text
            'pHandle.wSamplesPerSec = Text8.Text
                    pHandle.wBytePerSample = pHandle.wBitsPerSample / 8 * pHandle.wChannels
                    pHandle.wBytePerSec = pHandle.wBytePerSample * pHandle.wSamplesPerSec
                    pHandle.wDataSize = (UBound(LongData) + 1) * pHandle.wChannels * (pHandle.wBitsPerSample / 8)
                    pHandle.wfdataSize = pHandle.wDataSize + 36

                    Text1.Text = pHandle.wRiffFormatTag
                    Text2.Text = pHandle.wfdataSize
                    Text3.Text = pHandle.wFormatTag
                    Text4.Text = pHandle.wFormatName
                    Text5.Text = pHandle.wCsize
                    Text6.Text = pHandle.wWavefmt
                    Text7.Text = pHandle.wChannels
                    Text8.Text = pHandle.wSamplesPerSec
                    Text9.Text = pHandle.wBytePerSec
                    Text10.Text = pHandle.wBytePerSample
                    Text11.Text = pHandle.wBitsPerSample
                    Text12.Text = pHandle.wData
                    Text13.Text = pHandle.wDataSize
                    
                    PaintWave
                        startX = Picture1.ScaleLeft
                        Line1.X1 = startX
                        Line1.X2 = startX
                        endX = Picture1.ScaleWidth + Picture1.ScaleLeft - 0
                        Line2.X1 = endX
                        Line2.X2 = endX
                        
        makeCu
        PIC2Option_Click (9)
                        
                        
End Sub

Private Sub CommandTransfer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
         Exit Sub                      '忘了干嘛？
        On Error Resume Next
        If Button <> vbRightButton Then Exit Sub
        
        FFTsize = Val(LabelFFTsize.Caption)
        Haming = CheckHaming.Value
        ReDim FFTArr(FFTsize - 1 + 3)    '实傅利叶变换  算法 用刀后两哦位置
        Call aFFT(LongData(), startX, FFTsize / 2, Haming, FFTArr())
        ReDim Preserve FFTArr(FFTsize - 1)     '实傅利叶变换  算法 用刀后两哦位置
        '........            FFTArr()
                Dim i As Long, at As Double, TARR() As Double
                ReDim TARR(FFTsize - 1 + 3)
                For i = 0 To FFTsize - 1
                        TARR(i) = FFTArr(i)
                Next i
                ReDim FFTArr(FFTsize - 1 + 3)    '0
                For i = 0 To FFTsize - 1
                        'If at < Abs(FFTArr(i)) Then at = Abs(FFTArr(i))
                        'If i < FFTsize / 4 * 2 Then FFTArr(i) = 0
                        FFTArr(i) = TARR(i * 2)
                Next i
        
        
        ReDim Preserve FFTArr(FFTsize - 1 + 3)    '实傅利叶变换  算法 用刀后两哦位置
        Call aTFF(LongData(), startX, FFTsize / 2, Haming, FFTArr())

        PaintWave
        makeCu
        PIC2Option_Click (9)
End Sub

Private Sub CommandVolume_Click()
    Dim RetVal
    RetVal = Shell("sndvol32.exe", 1)
    RetVal = Shell("sndvol32.exe", 1)
    SendKeys "%p" & "r" & "{TAB}" & "{DOWN}" & "{ENTER}" ', True
End Sub
Private Sub Form_Click()

        
        
        Dim Arr(2, 1) As Integer
        Picture1.Print VarPtr(Arr(0, 0)); "(0, 0)"
        Picture1.Print VarPtr(Arr(1, 0)); "(1, 0)"
        Picture1.Print VarPtr(Arr(2, 0)); "(2, 0)"
        Picture1.Print VarPtr(Arr(0, 1)); "(0, 1)"
        Picture1.Print VarPtr(Arr(1, 1)); "(1, 1)"
        Picture1.Print VarPtr(Arr(2, 1)); "(2, 1)"
End Sub


Private Sub LabelFFTsize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbRightButton Then LabelFFTsize.Caption = Val(LabelFFTsize.Caption) * 2
        If Button = vbLeftButton Then LabelFFTsize.Caption = Val(LabelFFTsize.Caption) / 2
        If Val(LabelFFTsize.Caption) < 4 Then LabelFFTsize.Caption = 4
        If Val(LabelFFTsize.Caption) > 536870912 Then LabelFFTsize.Caption = 536870912
        PIC2Option_Click (9)
End Sub
Private Sub Commandopen_Click()

    CommonDialog1.Filter = "WAV (*.wav)|*.wav"
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then Exit Sub

    If ReadWaveFile(CommonDialog1.FileName) <> 1 Then Exit Sub
        
        Form1.Caption = CommonDialog1.FileName
        
            PaintWave
            startX = Picture1.ScaleLeft
            Line1.X1 = startX
            Line1.X2 = startX
            endX = Picture1.ScaleWidth + Picture1.ScaleLeft - 0
            Line2.X1 = endX
            Line2.X2 = endX
                    makeCu
                    PIC2Option_Click (9)
        

Errhandler:
    Exit Sub

End Sub

Private Sub Picture1_Click()
    '    Dim ii
    '    Picture1.Print "."     '???????????????????????????? bu ke shao
    '    ii = BitBlt(Picture1.hDC, 1, 0, Picture1.Width, Picture1.Height, Picture1.hDC, 0, 0, &HCC0021)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        'On Error Resume Next
        If Button = vbLeftButton Then startX = X: Line1.X1 = startX: Line1.X2 = startX
        If Button = vbRightButton Then endX = X: Line2.X1 = endX: Line2.X2 = endX
        
        'PaintWave
        makeCu
        PIC2Option_Click (9)
End Sub

Private Sub PIC2Option_Click(Index As Integer)
        If PIC2Option(0).Value = True Then PIC2WaveForm
        If PIC2Option(1).Value = True Then PIC2FFT
        If PIC2Option(2).Value = True Then PIC2Spectrum
        Picture3Draw
End Sub


Private Sub PIC2Spectrum()
        On Error Resume Next
        
        FFTsize = Val(LabelFFTsize.Caption)
        If FFTsize > 10000 Then Exit Sub: Beep
        ReDim FFTArr(FFTsize - 1 + 3)    '实傅利叶变换  算法 用刀后两哦位置
        Haming = CheckHaming.Value
        
        Dim N As Long, i As Long, J As Long, K As Long, STEPt As Long, STEPj As Long
        N = Abs(endX - startX)
        STEPt = Int(N / Val(TextHn.Text)) + 1 'PIC2 只画 stepT  个点
        Picture2.Cls
        Picture2.Scale (0, FFTsize / 2)-(N, 0)   '設定繪圖區域座標
        Dim tq As Double, tmax As Double
        Dim C1 As Integer, C2 As Integer
        'Debug.Print tmax
        For J = 0 To 9
            ReDim FFTArr(FFTsize - 1 + 3)    '实傅利叶变换  算法 用刀后两哦位置
           Call aFFT(LongData(), startX + Rnd * N, FFTsize / 2, Haming, FFTArr())
           'Call aFFT(CuDataS(), Rnd * N, FFTsize / 2, Haming, FFTArr())
            For i = 0 To FFTsize - 1
                tq = Log(Abs(FFTArr(i)))
                If tmax < tq Then tmax = tq
            Next i
            'Debug.Print tmax
        Next J
       STEPj = FFTsize / 2 / Val(TextVn.Text)
       For i = 0 To N Step STEPt 'save time
            ReDim FFTArr(FFTsize - 1 + 3)  '   此步必须！！！！！！！！！！  '实傅利叶变换  算法 用刀后两哦位置
            'Call aFFT(CuDataS(), i, FFTsize / 2, Haming, FFTArr())
            Call aFFT(LongData(), startX + i, FFTsize / 2, Haming, FFTArr())
            For J = 0 To FFTsize / 2 Step STEPj
                  tq = 0
                  For K = 0 To STEPj - 1
                        tq = tq + Log(Abs(FFTArr(J + K)))
                  Next K
                  tq = tq / STEPj / tmax * 455
                  'Debug.Print tq
                  If tq > 255 Then
                      C1 = 255
                      C2 = tq - 255
                  Else
                        C1 = tq
                        C2 = 0
                  End If
                  Picture2.PSet (i, J), RGB(C2, C2, C1)
            Next J
            DoEvents
        Next i
        
        Picture2.Line (Picture2.ScaleWidth / 300, 0#)-(Picture2.ScaleWidth / 300, FFTsize)

        For i = 0 To pHandle.wSamplesPerSec / 2 Step 500 * (Int(pHandle.wSamplesPerSec / 2 / 4 / 500) + 1)
               Picture2.Line (Picture2.ScaleWidth / 300, i / pHandle.wSamplesPerSec * 2 * FFTsize / 2)-(Picture2.ScaleWidth / 100, i / pHandle.wSamplesPerSec * 2 * FFTsize / 2)
               Picture2.Print Str(i)
        Next i
        'Picture2.Print "Hz"
        Picture2.CurrentY = FFTsize / 2
        Picture2.Print "stepT="; Int(N / STEPt) + 1; "StartX"; startX; "endX"; endX; " N="; N; Int(1 / CSng(CuHandle.wSamplesPerSec) * CSng(N) * 1000000) / 1000; "ms"
        
        For i = 0 To N / (FFTsize / 2)
                Picture2.Line (i * FFTsize, 0)-(i * FFTsize, FFTsize / 2 / 100 * (i Mod 10)), QBColor(4)
        Next i

End Sub
Private Sub PIC2FFT()
        On Error Resume Next
        FFTsize = Val(LabelFFTsize.Caption)
        ReDim FFTArr(FFTsize - 1 + 3)    '实傅利叶变换  算法 用刀后两哦位置
        Dim pFFTArr() As Double    '相位
        ReDim pFFTArr(FFTsize - 1 + 3)    '相位
        Haming = CheckHaming.Value
        Dim i As Long
        Dim tq As Double, tp As Double, C1 As Integer, C2 As Integer
        Picture2.Cls
         
        Call pFFT(LongData(), startX, FFTsize / 2, Haming, pFFTArr())  '位置以startX（绿杠）为准
        Call aFFT(LongData(), startX, FFTsize / 2, Haming, FFTArr())  '位置以startX（绿杠）为准
        For i = 0 To FFTsize / 2 - 1
                If tq < FFTArr(i) Then tq = FFTArr(i)
        Next i
        'Debug.Print tq
      

         
         '下边画幅谱
        'ReDim FFTArr(FFTsize - 1 + 3)    '实傅利叶变换  算法 用刀后两哦位置
        'Call aFFT(CuDataS(), FFTsize / 2, FFTsize / 2, Haming, FFTArr())  '位置以CuDataL()的FFTsize / 2为准
        
        'ReDim Preserve FFTArr(FFTsize / 2 - 1)       '减半
        'Debug.Print tq
        Picture2.Scale (0, tq)-(FFTsize / 2, 0)
        Picture2.CurrentX = 0
        Picture2.CurrentY = FFTArr(0)
        For i = 0 To FFTsize / 2 - 1
               Picture2.Line -(i, (FFTArr(i))), RGB(0, 60, 0)
        Next i
        For i = 0 To FFTsize / 2 - 1
               Picture2.PSet (i, (FFTArr(i))), QBColor(10)
        Next i
        Picture2.Line (0#, (tq) + Picture2.ScaleHeight / 100)-(FFTsize / 2, (tq) + Picture2.ScaleHeight / 100)
        
        For i = 0 To pHandle.wSamplesPerSec / 2 Step 500 * (Int(pHandle.wSamplesPerSec / 2 / 10 / 500) + 1)
               Picture2.Line (i / pHandle.wSamplesPerSec * 2 * FFTsize / 2, (tq) + Picture2.ScaleHeight / 100)-(i / pHandle.wSamplesPerSec * 2 * FFTsize / 2, (tq) + Picture2.ScaleHeight / 20)
               Picture2.Print Str(i)
        Next i
        Picture2.Print "Hz"
        Picture2.CurrentY = (tq) * 0.854
        Picture2.CurrentX = FFTsize / 2 * startX / UBound(LongData)
        Picture2.Print "StartX="; startX
        
        Picture2.Scale (0, Log(tq))-(FFTsize / 2, 0)
        Picture2.CurrentX = 0
        Picture2.CurrentY = Log(FFTArr(0))
        For i = 0 To FFTsize / 2 - 1
               Picture2.Line -(i, Log(FFTArr(i))), RGB(35, 35, 35)
               'Picture2.PSet (i, Log(FFTArr(i))), QBColor(2)
               'Debug.Print ; FFTArr(i)
        Next i
        
         'For i = 0 To FFTsize - 1
         '      tp = (Abs(FFTArr(i)) / tq * 510)
         '      If tp > 255 Then
         '          C1 = 255
         '          C2 = tp - 255
         '      Else
         '        C1 = tp
         '        C2 = 0
         '      End If
         '      Picture2.Line (i, Log(tq) + Picture2.ScaleHeight / 50)-(i, Log(tq)), RGB(C1, C2, C2)
         'Next i
        

        '下边画相位
        
        'ReDim Preserve FFTArr(FFTsize / 2 - 1)       '减半
        'Debug.Print tq
        Picture2.Scale (0, 3.14159265358979)-(FFTsize / 2, -3.14159265358979)
        Picture2.CurrentX = 0
        Picture2.CurrentY = pFFTArr(0)
        For i = 0 To FFTsize / 2 - 1
               'Picture2.Line -(i, (pFFTArr(i))), RGB(300 * FFTArr(i) / tq, 300 * FFTArr(i) / tq, 300 * FFTArr(i) / tq)
               Picture2.PSet (i, (pFFTArr(i))), RGB(666 * FFTArr(i) / tq, 26, 666 * FFTArr(i) / tq)
        Next i
        
        Picture2.Line (0#, 0)-(FFTsize / 2, 0), RGB(15, 15, 15)
        Picture2.Line (0#, 1.5707963267949)-(FFTsize / 2, 1.5707963267949), RGB(15, 15, 15)
        Picture2.Line (0#, -1.5707963267949)-(FFTsize / 2, -1.5707963267949), RGB(15, 15, 15)
  
End Sub
Private Sub aTFF(DataInt() As Integer, XX As Long, sizeFF As Long, HamingO As Boolean, ByRef DATA_FFT() As Double)
        '逆付利叶变换
        On Error Resume Next
        Call REALFT(DATA_FFT(), sizeFF, -1)  '-1
        Dim It As Long
        For It = 0 To sizeFF - 1
            If It + XX - sizeFF / 2 >= LBound(DataInt) Or It + XX - sizeFF / 2 <= UBound(DataInt) Then
                If HamingO = True Then
                    DATA_FFT(It) = DATA_FFT(It) / (0.54 - 0.46 * Cos(2 * It * 3.14159265358 / sizeFF))
                End If
                DataInt(It + XX - sizeFF / 2) = DATA_FFT(It) / sizeFF
            End If
            
        Next It
        

End Sub
Private Sub aFFT(DataInt() As Integer, XX As Long, sizeFF As Long, HamingO As Boolean, ByRef DATA_FFT() As Double)
        'XXXXXX   计算   幅度谱   XXXXXXXXXXX         sizeFF = FFTsize / 2
        On Error Resume Next
        Dim It As Long
        For It = 0 To sizeFF * 2 - 1
            If It + XX - sizeFF < LBound(DataInt) Or It + XX - sizeFF > UBound(DataInt) Then
                DATA_FFT(It) = 0#
            Else
                DATA_FFT(It) = CDbl(DataInt(It + XX - sizeFF))
                If HamingO = True Then
                    DATA_FFT(It) = DATA_FFT(It) * (0.54 - 0.46 * Cos(2 * 3.14159265358 * It / (sizeFF * 2)))
                    'Debug.Print DATA_FFT(It), (0.54 - 0.46 * Cos(2 * 3.14159265358 * It / sizeFF / 2))
                End If
            End If
            
        Next It
        Call REALFT(DATA_FFT(), sizeFF, 1)
        
        For It = 0 To sizeFF - 1
                 DATA_FFT(It) = Sqr(DATA_FFT(It * 2) * DATA_FFT(It * 2) + DATA_FFT(It * 2 + 1) * DATA_FFT(It * 2 + 1))
                 'DATA_FFT(It) = Atn(DATA_FFT(It * 2 + 1) / DATA_FFT(It * 2))
        Next It

End Sub
Private Sub pFFT(DataInt() As Integer, XX As Long, sizeFF As Long, HamingO As Boolean, ByRef DATA_FFT() As Double)
        'XXXXXX   计算   相位谱   XXXXXXXXXXX         sizeFF = FFTsize / 2
        On Error Resume Next
        Dim It As Long
        For It = 0 To sizeFF * 2 - 1
            If It + XX - sizeFF < LBound(DataInt) Or It + XX - sizeFF > UBound(DataInt) Then
                DATA_FFT(It) = 0#
            Else
                DATA_FFT(It) = CDbl(DataInt(It + XX - sizeFF))
                If HamingO = True Then
                    DATA_FFT(It) = DATA_FFT(It) * (0.54 - 0.46 * Cos(2 * 3.14159265358 * It / (sizeFF * 2)))
                    'Debug.Print DATA_FFT(It), (0.54 - 0.46 * Cos(2 * 3.14159265358 * It / sizeFF / 2))
                End If
            End If
            
        Next It
        Call REALFT(DATA_FFT(), sizeFF, 1)
        
        For It = 0 To sizeFF - 1
                 'DATA_FFT(It) = Sqr(DATA_FFT(It * 2) * DATA_FFT(It * 2) + DATA_FFT(It * 2 + 1) * DATA_FFT(It * 2 + 1))
                 DATA_FFT(It) = Atn(DATA_FFT(It * 2 + 1) / DATA_FFT(It * 2))
        Next It

End Sub

Private Sub PIC2WaveForm()
        On Error Resume Next
        Dim N As Long, i As Long, STEPt As Long
        N = UBound(CuDataS)
        STEPt = 2048   'PIC2 只画 stepT  个点
        Picture2.Cls
        If CuHandle.wBitsPerSample = 8 Then Picture2.Scale (0, 256)-(N, 0)     '設定繪圖區域座標
        If CuHandle.wBitsPerSample = 16 Then Picture2.Scale (0, 32767)-(N, -32768)
        If CuHandle.wBitsPerSample = 8 Then Picture2.Line (N, 128)-(0, 128), QBColor(7)
        If CuHandle.wBitsPerSample = 16 Then Picture2.Line (N, 0)-(0, 0), QBColor(7)
        For i = 0 To N 'Step Int(N / STEPt) + 1  'save time
            Picture2.Line -(i, CuDataS(i)), QBColor(2)
            'Picture2.Circle (I, CuDataS(I)), Picture2.ScaleWidth / 800
            'Picture2.Line (I, CuDataS(I) - Picture2.ScaleHeight / 100)-(I, CuDataS(I))
            'Picture2.Line (I - Picture2.ScaleWidth / 400, CuDataS(I))-(I, CuDataS(I))
        Next i
        If (N > STEPt) = 0 Then     'n>=STEPt
            For i = 0 To N 'Step Int(N / STEPt) + 1 'save time
                'Picture2.Circle (I, CuDataS(I)), Picture2.ScaleWidth / 800
                Picture2.PSet (i, CuDataS(i))
            Next i
        End If
        If CuHandle.wBitsPerSample = 8 Then Picture2.CurrentY = 256     '設定繪圖區域座標
        If CuHandle.wBitsPerSample = 16 Then Picture2.CurrentY = 32767
        Picture2.CurrentX = 0
        Picture2.Print "stepT="; Int(N / STEPt) + 1; "StartX"; startX; "endX"; endX; " N="; N; Int(1 / CSng(CuHandle.wSamplesPerSec) * CSng(N) * 1000000) / 1000; "ms"
        
        For i = 0 To N / LabelFFTsize.Caption
             Picture2.Line (i * LabelFFTsize.Caption, 2 ^ CuHandle.wBitsPerSample / 100 * (i Mod 10))-(i * LabelFFTsize.Caption, 0), QBColor(4)
        Next i

End Sub

Private Sub Form_Load()
        'Picture1.Print "DoubleClick here to open file"
        Picture1.Scale (0, 1)-(1, 0)
        ReDim LongData(0)
        Recording = False
        Playing = False
        Viewing = False
        'SoundMeter.BUFFER_SIZE = 800
        TimerstopPlay.Enabled = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    SoundMeter.StopInput

End Sub
Private Sub Commandview_Click()
        
        If Viewing = False Then
                CommandView.Caption = "viewing ..."
                CommandRecorD.Enabled = True
                Viewing = True
                
                Text14.Enabled = False
                Text15.Enabled = False
                Text16.Enabled = False
                Text17.Enabled = False
    
                NUM_BUFFERS = Text21.Text
                ReDim hmem(NUM_BUFFERS - 1) As Long, inHdr(NUM_BUFFERS - 1) As WAVEHDR
                SoundMeter.StartInput Text14.Text, Text15.Text, Text16.Text
                
                Timer1.Interval = Text17.Text
                Timer1.Enabled = True
        Else
                CommandView.Caption = "view"
                CommandRecorD.Enabled = False
                Viewing = False
                
                Text14.Enabled = True
                Text15.Enabled = True
                Text16.Enabled = True
                Text17.Enabled = True
                
                If Recording = True Then Call CommandRecorD_Click
        
                'CommandRecorD.Caption = "RecorDown"
                'Recording = False
                Timer1.Enabled = False
                SoundMeter.StopInput
        End If
End Sub

Private Sub Picture3Draw()
        On Error Resume Next
        
        Dim N As Long, i As Long, STEPt As Long, tts As Single
        N = UBound(CuDataS)
        
        frameN = Val(TextshortTimeN.Caption)               ' 帧长    点数
        frameStepN = Val(TextshortTimeN2.Caption)      ' 帧间步长    点数
        frameMS = Val(TextshortTimeMS.Caption)               ' 帧长    毫秒数
        frameStepMS = Val(TextshortTime2MS.Caption)      ' 帧间步长    毫秒数

        STEPt = Val(TextshortTimeN2.Caption)
        Picture3.Cls
        'If pHandle.wBitsPerSample = 8 Then Picture3.Scale (0, 256)-(N, -256)    '設定繪圖區域座標
        'If pHandle.wBitsPerSample = 16 Then Picture3.Scale (0, 32767)-(N, -32768)
        'If CuHandle.wBitsPerSample = 8 Then Picture3.Line (N, 0)-(0, 0), QBColor(7)
        'If CuHandle.wBitsPerSample = 16 Then Picture3.Line (N, 0)-(0, 0), QBColor(7)
        Picture3.Scale (0, 1)-(N, 0)
        Picture3.Line (N, 0)-(0, 0), QBColor(7)
        
        If CheckShortTimeAveScope.Value = 1 Then
                    Picture3.CurrentX = 0
                    Picture3.CurrentY = 0
                For i = 0 To N Step STEPt     'save time
                    tts = ShortTimeAveScope(CuDataS(), i, Val(TextshortTimeN.Caption))
                    Picture3.Line -(i, tts), QBColor(2)
                    'Picture3.PSet (i, tts)
                    'Debug.Print ShortTimeAveScope(CuDataS(), i, Val(TextshortTimeN.Caption))
                Next i
        End If
        If CheckShortTimeAveCross0.Value = 1 Then
                    Picture3.CurrentX = 0
                    Picture3.CurrentY = 0
                For i = 0 To N Step STEPt     'save time
                    Picture3.Line -(i, ShortTimeAvecross0(CuDataS(), i, Val(TextshortTimeN.Caption))), QBColor(6)
                    'Debug.Print ShortTimeAveScope(CuDataS(), i, Val(TextshortTimeN.Caption))
                Next i
        End If
        If CheckShortTimeAveEnergy.Value = 1 Then
                    Picture3.CurrentX = 0
                    Picture3.CurrentY = 0
                For i = 0 To N Step STEPt   'save time
                    Picture3.Line -(i, ShortTimeAveEnergy(CuDataS(), i, Val(TextshortTimeN.Caption))), QBColor(14)
                    'Debug.Print ShortTimeAveScope(CuDataS(), i, Val(TextshortTimeN.Caption))
                Next i
        End If
        If CheckShortTimeAvecrossPer.Value = 1 Then
                    Picture3.CurrentX = 0
                    Picture3.CurrentY = 0
                For i = 0 To N Step STEPt    'save time
                    Picture3.Line -(i, ShortTimeAvecrossPer(CuDataS(), i, Val(TextshortTimeN.Caption), CheckShortTimeAvecrossPer.Tag)), QBColor(13)
                    'Debug.Print ShortTimeAvecrossPer(CuDataS(), i, Val(TextshortTimeN.Caption), 0#)
                Next i
        End If
        For i = 0 To N / frameStepN
                Picture3.Line (i * frameStepN, 1 - 0.01 * (i Mod 10))-(i * frameStepN, 1), QBColor(4)
        Next i
        
End Sub

Private Sub Picture10_Click()
         Picture10.Tag = Val(Picture10.Tag) * -1
         If Val(Picture10.Tag) = -1 Then
                Picture10.ZOrder 1
         Else
                Picture10.ZOrder 0
         End If

End Sub

Private Sub Text18_Click() ' test
                    Dim A() As Long, B(3) As Long
                    ReDim A(3) As Long
                    A(0) = 5.78
                    A(1) = 6.78
                    A(2) = 787
                    A(3) = 878
                    B(0) = 1.555
                    B(1) = 278.22
                    B(2) = 3.4
                    B(3) = 4.5
                    ReDim Preserve A(7) As Long
                    CopyMemory VarPtr(A(0)), VarPtr(B(0)), 16

End Sub

Private Sub TextPreEmphasis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbLeftButton Then TextPreEmphasis.Caption = Val(TextPreEmphasis.Caption) - 0.005
        If Button = vbRightButton Then TextPreEmphasis.Caption = Val(TextPreEmphasis.Caption) + 0.005
        If Val(TextPreEmphasis.Caption) < 0# Then TextPreEmphasis.Caption = 1#
        If Val(TextPreEmphasis.Caption) > 1 Then TextPreEmphasis.Caption = 0#
        If Checkbefore.Value = 1 Then
            makeCu
            PIC2Option_Click (9)
        End If
End Sub

Private Sub TextshortTimeN_Change()
     On Error Resume Next
     TextshortTimeN.Caption = Str(Int(Val(TextshortTimeN.Caption))) & "p"
     TextshortTimeMS.Caption = Str(Int(Val(TextshortTimeN.Caption) / CuHandle.wSamplesPerSec * 10000 + 0.0001) / 10) & "ms"
     Picture3Draw
End Sub
Private Sub TextshortTimeN2_Change()
     On Error Resume Next
     TextshortTimeN2.Caption = Str(Int(Val(TextshortTimeN2.Caption))) & "p"
     TextshortTime2MS.Caption = Str(Int(Val(TextshortTimeN2.Caption) / CuHandle.wSamplesPerSec * 10000 + 0.0001) / 10) & "ms"
     Picture3Draw
End Sub

Private Sub TextshortTimeN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbRightButton And Shift = vbCtrlMask Then TextshortTimeN.Caption = Val(TextshortTimeN.Caption) + 1
        If Button = vbLeftButton And Shift = vbCtrlMask Then TextshortTimeN.Caption = Val(TextshortTimeN.Caption) - 1
        If Button = vbRightButton And Shift <> vbCtrlMask Then TextshortTimeN.Caption = Val(TextshortTimeN.Caption) * 2
        If Button = vbLeftButton And Shift <> vbCtrlMask Then TextshortTimeN.Caption = Val(TextshortTimeN.Caption) / 2
        If Val(TextshortTimeN.Caption) < 16 Then TextshortTimeN.Caption = 16
        If Val(TextshortTimeN.Caption) > 8192 Then TextshortTimeN.Caption = 8192
End Sub

Private Sub TextshortTimeN2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbRightButton And Shift = vbCtrlMask Then TextshortTimeN2.Caption = Val(TextshortTimeN2.Caption) + 1
        If Button = vbLeftButton And Shift = vbCtrlMask Then TextshortTimeN2.Caption = Val(TextshortTimeN2.Caption) - 1
        If Button = vbRightButton And Shift <> vbCtrlMask Then TextshortTimeN2.Caption = Val(TextshortTimeN2.Caption) * 2
        If Button = vbLeftButton And Shift <> vbCtrlMask Then TextshortTimeN2.Caption = Val(TextshortTimeN2.Caption) / 2
        If Val(TextshortTimeN2.Caption) < 16 Then TextshortTimeN2.Caption = 16
        If Val(TextshortTimeN2.Caption) > 8192 Then TextshortTimeN2.Caption = 8192
End Sub





Private Sub Timer1_Timer()

    'ProgressBar1.Value = SoundMeter.getVolume(buffaddress)
   drawScope Text14.Text, Text15.Text, Text16.Text   ' If Not Recording Then
    If Recording = True Then
            pHandle.wDataSize = (UBound(LongData) + 1) * pHandle.wChannels * (pHandle.wBitsPerSample / 8)
            pHandle.wfdataSize = pHandle.wDataSize + 36
            Text13.Text = pHandle.wDataSize
            Text2.Text = pHandle.wfdataSize
             'PaintWave 'If pHandle.wDataSize > 99 Then
    End If
End Sub

Private Sub drawScope(BpS As Integer, Sps As Long, BUsampleNumber As Long)
   Dim N As Integer
   Dim avg As Integer, tempval As Integer, posval  As Integer
   ' CopyStructFromPtr audbyteArray, buffaddress, SoundMeter.BUFFER_SIZE
    
    scopeBox.Cls
    tempval = 0
    avg = 0
    posval = 0

        If BpS = 8 Then
                scopeBox.Scale (0, 256)-(BUsampleNumber, 0)         '設定繪圖區域座標
                
                For N = 0 To BUsampleNumber - 1
                    scopeBox.Line -(N, audByteArray(N)), QBColor(9)
                    posval = audByteArray(N) - 128
                    If posval < 0 Then posval = (-posval)
                    If posval > tempval Then tempval = posval
                    'Debug.Print audbyteArray(3); audbyteArray(99)
                Next N
            
               Picture8.Height = Picture9.Height * (128 - tempval) / 128
         ElseIf BpS = 16 Then
                scopeBox.Scale (0, -32768)-(BUsampleNumber, 32768)
               
                For N = 0 To BUsampleNumber - 1
                    scopeBox.Line -(N, audIntArray(N)), QBColor(9)
                    posval = audIntArray(N)
                    If posval < 0 Then posval = 0 - (posval + 1)
                    If posval > tempval Then tempval = posval
                    'Debug.Print audbyteArray(3); audbyteArray(99)
                Next N
            
               Picture8.Height = Picture9.Height * (32768 - tempval) / 32768
        End If
        

End Sub

Private Sub CommandplayL_Click()
        On Error Resume Next
        If Playing = False Then
                CommandplayL.Caption = "playing..."
                Playing = True
                TimerstopPlay.Enabled = True
                
                    makeCu
                    If CuHandle.wBitsPerSample = 8 Then
                            Dim byteTarr() As Byte, i As Long
                            ReDim byteTarr(UBound(CuDataS))
                            For i = 0 To UBound(CuDataS)
                                byteTarr(i) = CByte(CuDataS(i))
                            Next i
                        Play2 1, Text11.Text, Text8.Text, VarPtr(byteTarr(0)), UBound(CuDataS)
                    Else
                        Play2 1, Text11.Text, Text8.Text, VarPtr(CuDataS(0)), UBound(CuDataS) * 2
                    End If
                
        Else
                'StopPlay
                Playing = False
                CommandplayL.Caption = "playL(cu)"
                
        End If
End Sub

Private Sub TimerstopPlay_Timer()
         If Playing = False Then CommandplayL.Caption = "playL": TimerstopPlay.Enabled = False: StopPlay
End Sub
