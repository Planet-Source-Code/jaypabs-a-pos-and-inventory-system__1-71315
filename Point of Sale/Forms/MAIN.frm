VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "StyleButtonX.ocx"
Begin VB.MDIForm MAIN 
   BackColor       =   &H8000000C&
   Caption         =   "POS & Inventory System v1.0"
   ClientHeight    =   7995
   ClientLeft      =   5715
   ClientTop       =   3300
   ClientWidth     =   13950
   Icon            =   "MAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList i16x16g 
      Left            =   3525
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3452
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":39EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":4320
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":46BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":4A54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLine 
      Align           =   1  'Align Top
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   10
      Index           =   5
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   13950
      TabIndex        =   14
      Top             =   870
      Width           =   13950
   End
   Begin VB.PictureBox picLine 
      Align           =   1  'Align Top
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   10
      Index           =   4
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   13950
      TabIndex        =   13
      Top             =   885
      Width           =   13950
   End
   Begin VB.PictureBox picLine 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   50
      Index           =   2
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   13950
      TabIndex        =   10
      Top             =   900
      Width           =   13950
   End
   Begin VB.PictureBox picLine 
      Align           =   1  'Align Top
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   10
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   13950
      TabIndex        =   9
      Top             =   855
      Width           =   13950
   End
   Begin VB.PictureBox picLine 
      Align           =   1  'Align Top
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   10
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   13950
      TabIndex        =   8
      Top             =   840
      Width           =   13950
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4800
      Top             =   3750
   End
   Begin VB.PictureBox picSeparator 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6720
      Left            =   11520
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6720
      ScaleWidth      =   120
      TabIndex        =   6
      Top             =   945
      Width           =   125
      Begin StyleButtonX.StyleButton StyleButton2 
         Height          =   1095
         Left            =   0
         TabIndex        =   7
         Top             =   1920
         Width           =   125
         _ExtentX        =   212
         _ExtentY        =   1931
         UpColorTop1     =   -2147483633
         UpColorTop2     =   -2147483633
         UpColorTop3     =   -2147483633
         UpColorTop4     =   -2147483633
         UpColorButtom1  =   -2147483633
         UpColorButtom2  =   -2147483633
         UpColorButtom3  =   -2147483633
         UpColorButtom4  =   -2147483633
         UpColorLeft1    =   -2147483633
         UpColorLeft2    =   -2147483633
         UpColorLeft3    =   -2147483633
         UpColorLeft4    =   -2147483633
         UpColorRight1   =   -2147483633
         UpColorRight2   =   -2147483633
         UpColorRight3   =   -2147483633
         UpColorRight4   =   -2147483633
         DownColorTop1   =   7021576
         DownColorTop2   =   -2147483633
         DownColorTop3   =   -2147483633
         DownColorTop4   =   -2147483633
         DownColorButtom1=   7021576
         DownColorButtom2=   -2147483633
         DownColorButtom3=   -2147483633
         DownColorButtom4=   -2147483633
         DownColorLeft1  =   7021576
         DownColorLeft2  =   -2147483633
         DownColorLeft3  =   -2147483633
         DownColorLeft4  =   -2147483633
         DownColorRight1 =   7021576
         DownColorRight2 =   -2147483633
         DownColorRight3 =   -2147483633
         DownColorRight4 =   -2147483633
         HoverColorTop1  =   7021576
         HoverColorTop2  =   -2147483633
         HoverColorTop3  =   -2147483633
         HoverColorTop4  =   -2147483633
         HoverColorButtom1=   7021576
         HoverColorButtom2=   -2147483633
         HoverColorButtom3=   -2147483633
         HoverColorButtom4=   -2147483633
         HoverColorLeft1 =   7021576
         HoverColorLeft2 =   -2147483633
         HoverColorLeft3 =   -2147483633
         HoverColorLeft4 =   -2147483633
         HoverColorRight1=   7021576
         HoverColorRight2=   -2147483633
         HoverColorRight3=   -2147483633
         HoverColorRight4=   -2147483633
         FocusColorTop1  =   7021576
         FocusColorTop2  =   -2147483633
         FocusColorTop3  =   -2147483633
         FocusColorTop4  =   -2147483633
         FocusColorButtom1=   7021576
         FocusColorButtom2=   -2147483633
         FocusColorButtom3=   -2147483633
         FocusColorButtom4=   -2147483633
         FocusColorLeft1 =   7021576
         FocusColorLeft2 =   -2147483633
         FocusColorLeft3 =   -2147483633
         FocusColorLeft4 =   -2147483633
         FocusColorRight1=   7021576
         FocusColorRight2=   -2147483633
         FocusColorRight3=   -2147483633
         FocusColorRight4=   -2147483633
         DisabledColorTop1=   -2147483633
         DisabledColorTop2=   -2147483633
         DisabledColorTop3=   -2147483633
         DisabledColorTop4=   -2147483633
         DisabledColorButtom1=   -2147483633
         DisabledColorButtom2=   -2147483633
         DisabledColorButtom3=   -2147483633
         DisabledColorButtom4=   -2147483633
         DisabledColorLeft1=   -2147483633
         DisabledColorLeft2=   -2147483633
         DisabledColorLeft3=   -2147483633
         DisabledColorLeft4=   -2147483633
         DisabledColorRight1=   -2147483633
         DisabledColorRight2=   -2147483633
         DisabledColorRight3=   -2147483633
         DisabledColorRight4=   -2147483633
         Caption         =   ""
         MousePointer    =   1
         BackColorUp     =   -2147483633
         BackColorDown   =   11899524
         BackColorHover  =   14073525
         BackColorFocus  =   14604246
         BackColorDisabled=   -2147483633
         DotsInCornerColor=   16777215
         MoveWhenClick   =   0   'False
         ForeColorUp     =   -2147483630
         ForeColorDown   =   -2147483634
         ForeColorHover  =   -2147483630
         ForeColorFocus  =   -2147483630
         ForeColorDisabled=   12632256
         BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowBorderLevel2=   0   'False
         DistanceBetweenPictureAndCaption=   -50
      End
   End
   Begin VB.PictureBox picLeft 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6720
      Left            =   11640
      ScaleHeight     =   6720
      ScaleWidth      =   2310
      TabIndex        =   1
      Top             =   945
      Width           =   2310
      Begin VB.Frame Frame1 
         Height          =   465
         Left            =   0
         TabIndex        =   4
         Top             =   -75
         Width           =   2250
         Begin VB.Image Image1 
            Height          =   240
            Left            =   75
            Picture         =   "MAIN.frx":4DEE
            Top             =   150
            Width           =   240
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Opened Forms"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   375
            TabIndex        =   5
            Top             =   195
            Width           =   1290
         End
      End
      Begin MSComctlLib.ListView lvWin 
         Height          =   4050
         Left            =   0
         TabIndex        =   2
         Top             =   375
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   7144
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "MAIN.frx":57F0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Form Name"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Image Image5 
         Height          =   960
         Left            =   1950
         Picture         =   "MAIN.frx":64CA
         Top             =   6030
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   1950
         Picture         =   "MAIN.frx":7214
         Top             =   4950
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   3
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   13950
      TabIndex        =   0
      Top             =   7665
      Width           =   13950
   End
   Begin VB.Timer tmrMemStatus 
      Interval        =   1000
      Left            =   3600
      Top             =   5025
   End
   Begin MSComctlLib.ImageList ig24x24 
      Left            =   2925
      Top             =   2550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":7F5E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   2925
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":818B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":8B9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":95AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":9949
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":9CE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":A07D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":A417
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":AE29
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":B83B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":C24D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":CC5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":D671
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":E083
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":EA95
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":F031
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   7695
      Width           =   13950
      _ExtentX        =   24606
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   442
            MinWidth        =   442
            Picture         =   "MAIN.frx":F5CD
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "User Name:"
            TextSave        =   "User Name:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10575
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "MAIN.frx":F969
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10/22/2009"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "8:55 PM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   5400
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":FD03
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":11695
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":13027
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":149B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1634B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":17CDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1966F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1B001
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1C993
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1E327
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1F003
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1F8E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":205BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2129B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":21F77
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":22C53
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2392F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picContainer 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   13950
      TabIndex        =   11
      Top             =   0
      Width           =   13950
      Begin VB.PictureBox picFreeMem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   700
         Left            =   6600
         ScaleHeight     =   705
         ScaleWidth      =   2295
         TabIndex        =   15
         Top             =   75
         Width           =   2300
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AVAILABLE FREE MEMORY"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   165
            Left            =   120
            TabIndex        =   20
            Top             =   75
            Width           =   2070
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            BorderStyle     =   3  'Dot
            X1              =   0
            X2              =   2520
            Y1              =   250
            Y2              =   250
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FF00&
            X1              =   825
            X2              =   825
            Y1              =   300
            Y2              =   600
         End
         Begin VB.Label lblPMem 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "                    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   165
            Left            =   960
            TabIndex        =   19
            Top             =   315
            Width           =   900
         End
         Begin VB.Label lblVMem 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "                    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   165
            Left            =   960
            TabIndex        =   18
            Top             =   480
            Width           =   900
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Virtual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   165
            Left            =   75
            TabIndex        =   17
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Physical"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   165
            Left            =   120
            TabIndex        =   16
            Top             =   315
            Width           =   615
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   700
            Left            =   0
            Top             =   0
            Width           =   2295
         End
      End
      Begin MSComctlLib.Toolbar tbMenu 
         Height          =   810
         Left            =   0
         TabIndex        =   12
         Top             =   30
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   1429
         ButtonWidth     =   1508
         ButtonHeight    =   1429
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "itb32x32"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Shortcuts"
               Key             =   "Shortcuts"
               Object.ToolTipText     =   "Ctrl+F1"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               Key             =   "New"
               Object.ToolTipText     =   "Ctrl+F2"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               Key             =   "Edit"
               Object.ToolTipText     =   "Ctrl+F3"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Search"
               Key             =   "Search"
               Object.ToolTipText     =   "Ctrl+F4"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Key             =   "Delete"
               Object.ToolTipText     =   "Ctrl+F5"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "Refresh"
               Object.ToolTipText     =   "Ctrl+F6"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "Print"
               Object.ToolTipText     =   "Ctrl+F7"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Close"
               Key             =   "Close"
               Object.ToolTipText     =   "Ctrl+F8"
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFLO 
         Caption         =   "&Log out    "
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFE 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnu_Transactions 
      Caption         =   "&Transactions"
      Begin VB.Menu mnu_MSales 
         Caption         =   "Sales"
         Begin VB.Menu mnu_CashSales 
            Caption         =   "Cash Sales"
         End
         Begin VB.Menu mnu_CashSalesReturn 
            Caption         =   "Cash Sales Return"
         End
         Begin VB.Menu mnu_SalesOrder 
            Caption         =   "Sales Order"
         End
         Begin VB.Menu mnu_SalesOrderDelivery 
            Caption         =   "Sales Order Delivery"
         End
         Begin VB.Menu mnu_SalesOrderReturn 
            Caption         =   "Sales Order Return"
         End
      End
      Begin VB.Menu mnu_MPurchases 
         Caption         =   "Purchases"
         Begin VB.Menu mnu_LocalPurchase 
            Caption         =   "Local Purchase"
         End
         Begin VB.Menu mnu_LocalPurchaseReturn 
            Caption         =   "Local Purchase Return"
         End
         Begin VB.Menu mnu_PurchaseOrder 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnu_ReceiveItem 
            Caption         =   "Receive Item"
         End
         Begin VB.Menu mnu_PurchaseOrderReturn 
            Caption         =   "Purchase Order Return"
         End
      End
   End
   Begin VB.Menu mnu_Masterfiles 
      Caption         =   "&Masterfiles"
      Begin VB.Menu mnu_MClients 
         Caption         =   "Clients"
         Begin VB.Menu mnu_Clients 
            Caption         =   "Clients"
         End
         Begin VB.Menu mnu_Clients_Category 
            Caption         =   "Clients Category"
         End
      End
      Begin VB.Menu mnu_MVendors 
         Caption         =   "Vendors"
         Begin VB.Menu mnu_Vendors 
            Caption         =   "Vendors"
         End
         Begin VB.Menu mnu_Vendors_Category 
            Caption         =   "Vendors Category"
         End
      End
      Begin VB.Menu mnu_MStocks 
         Caption         =   "Stocks"
         Begin VB.Menu mnu_Stocks 
            Caption         =   "Stocks"
         End
         Begin VB.Menu mnu_Stocks_Category 
            Caption         =   "Stocks Category"
         End
         Begin VB.Menu mnu_Beg_Invty 
            Caption         =   "Beginning Inventory"
         End
         Begin VB.Menu mnu_Stocks_OUM 
            Caption         =   "Stocks UOM"
         End
      End
   End
   Begin VB.Menu mnuR 
      Caption         =   "&Reports"
      Begin VB.Menu mnuRDS 
         Caption         =   "&Daily Sales Report"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWS 
         Caption         =   "&Weekly Sales Report"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRMS 
         Caption         =   "&Monthly Sales Report"
         Visible         =   0   'False
      End
      Begin VB.Menu bar3543 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPCSI 
         Caption         =   "Periodic Cash Sales (Itemized)"
      End
      Begin VB.Menu mnuPCSR 
         Caption         =   "Periodic Cash Sales Return"
      End
      Begin VB.Menu mnuPCSRI 
         Caption         =   "Periodic Cash Sales Return (Itemized)"
      End
      Begin VB.Menu mnuPChS 
         Caption         =   "Periodic Charge Sales"
      End
      Begin VB.Menu mnuPChSI 
         Caption         =   "Periodic Charge Sales (Itemized)"
      End
      Begin VB.Menu mnuPChSR 
         Caption         =   "Periodic Charge Sales Return"
      End
      Begin VB.Menu mnuPChSRI 
         Caption         =   "Periodic Charge Sales Return (Itemized)"
      End
      Begin VB.Menu bar2378465 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockCard 
         Caption         =   "Stock Card"
      End
   End
   Begin VB.Menu mnuU 
      Caption         =   "&Utility"
      Begin VB.Menu mnuPrintBarcode 
         Caption         =   "Print Barcode"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuSMMU 
         Caption         =   "Manage &Users"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuUBI 
         Caption         =   "&Business Information"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuUC 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnuUN 
         Caption         =   "&Notepad"
      End
      Begin VB.Menu mnuUSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUWE 
         Caption         =   "Windows Explorer"
      End
   End
   Begin VB.Menu mnuRecA 
      Caption         =   "&Action"
      Begin VB.Menu mnuRASSM 
         Caption         =   "Show Shortcut &Menu"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuRASep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRACN 
         Caption         =   "Create &New"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuRAES 
         Caption         =   "&Edit Selected"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuRAS 
         Caption         =   "&Search"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuRADS 
         Caption         =   "&Delete Selected"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuRARR 
         Caption         =   "&Refresh"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuRAP 
         Caption         =   "&Print"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu mnuRASep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRAC 
         Caption         =   "&Close"
         Shortcut        =   ^{F8}
      End
   End
   Begin VB.Menu mnuH 
      Caption         =   "&Help"
      Begin VB.Menu mnuHUG 
         Caption         =   "&User's Guide"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuHKS 
         Caption         =   "&Keyboard Shortcut"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuSSS 
         Caption         =   "&Show Splash Screen"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuA 
         Caption         =   "&About"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuSO 
      Caption         =   "ShortcutOption"
      Visible         =   0   'False
      Begin VB.Menu mnuSOAD 
         Caption         =   "(Default)"
      End
      Begin VB.Menu mnuSOSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSOAHL 
         Caption         =   "Horizontal List"
      End
      Begin VB.Menu mnuSOAVL 
         Caption         =   "Vertical List"
      End
   End
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Var for point api function
Dim cursor_pos As POINTAPI

Public CloseMe  As Boolean

Dim resize_down     As Boolean
Dim show_mnu        As Boolean
Dim pos_num         As Integer

Private Sub lvWin_Click()
    If lvWin.ListItems.Count < 1 Then Exit Sub
    
    Select Case lvWin.SelectedItem.Key
        Case "frmShortcuts": frmShortcuts.show: frmShortcuts.WindowState = vbMaximized: frmShortcuts.SetFocus
    
        'For Clients
        Case "frmClients": LoadForm frmClients
        Case "frmClientsCategory": LoadForm frmClientsCategory
        
        'For Vendors
        Case "frmVendors": LoadForm frmVendors
        Case "frmVendorsCategory": LoadForm frmVendorsCategory
        
        'For Stocks
        Case "frmStocks": LoadForm frmStocks
        Case "frmStocksCategory": LoadForm frmStocksCategory
        Case "frmStocksUOM": LoadForm frmStocksUOM
        
        'For Sales
        Case "frmCashSales": LoadForm frmCashSales
        Case "frmfrmCashSalesReturn": LoadForm frmCashSalesReturn
        Case "frmSalesOrder": LoadForm frmSalesOrder
        Case "frmSalesOrderDelivery": LoadForm frmSalesOrderDelivery
        Case "frmSalesOrderReturn": LoadForm frmSalesOrderReturn
                
        'For Purchases
        Case "frmLocalPurchase": LoadForm frmLocalPurchase
        Case "frmLocalPurchaseReturn": LoadForm frmLocalPurchaseReturn
        Case "frmPurchaseOrder": LoadForm frmPurchaseOrder
        Case "frmPurchaseOrderReceive": LoadForm frmPurchaseOrderReceive
        Case "frmPurchaseOrderReturn": LoadForm frmPurchaseOrderReturn
        
        Case "frmUserRec"
            If CurrUser.USER_ISADMIN = False Then
                MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
            Else
                frmUserRec.show vbModal
            End If
        Case "frmBusinessInfo": frmBusinessInfo.show vbModal
    End Select
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If CloseMe = False Then If MsgBox("This will close the application. Do you want to proceed?", vbExclamation + vbYesNo) = vbNo Then Cancel = 1: Exit Sub
    'FRM_MESSAGE.show vbModal
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    picFreeMem.Left = (Me.Width - picFreeMem.ScaleWidth) - 200
End Sub

Private Sub mnu_CashSales_Click()
    frmCashSalesAE.show 1
End Sub

Private Sub mnu_CashSalesReturn_Click()
    LoadForm frmCashSalesReturn
End Sub

Private Sub mnu_Clients_Category_Click()
    LoadForm frmClientsCategory
End Sub

Private Sub mnu_Clients_Click()
    LoadForm frmClients
End Sub

Private Sub mnu_LocalPurchase_Click()
    LoadForm frmLocalPurchase
End Sub

Private Sub mnu_LocalPurchaseReturn_Click()
    LoadForm frmLocalPurchaseReturn
End Sub

Private Sub mnu_PurchaseOrder_Click()
    LoadForm frmPurchaseOrder
End Sub

Private Sub mnu_PurchaseOrderReturn_Click()
    LoadForm frmPurchaseOrderReturn
End Sub

Private Sub mnu_ReceiveItem_Click()
    LoadForm frmPurchaseOrderReceive
End Sub

Private Sub mnu_SalesOrder_Click()
    LoadForm frmSalesOrder
End Sub

Private Sub mnu_SalesOrderDelivery_Click()
    LoadForm frmSalesOrderDelivery
End Sub

Private Sub mnu_SalesOrderReturn_Click()
    LoadForm frmSalesOrderReturn
End Sub

Private Sub mnu_Stocks_Category_Click()
    LoadForm frmStocksCategory
End Sub

Private Sub mnu_Stocks_Click()
    LoadForm frmStocks
End Sub

Private Sub mnu_Stocks_OUM_Click()
    LoadForm frmStocksUOM
End Sub

Private Sub mnu_Vendors_Category_Click()
    LoadForm frmVendorsCategory
End Sub

Private Sub mnu_Vendors_Click()
    LoadForm frmVendors
End Sub

Private Sub mnuFE_Click()
    Unload Me
End Sub

Private Sub mnuFLO_Click()
    If MsgBox("Are you sure you want to log out?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    
    'SendMessage frmShortcuts.hwnd, WM_CLOSE, 0, 0
    UnloadChilds
    SendMessage frmShortcuts.hwnd, WM_ACTIVATE, 0, 0
    
    'ClearInfoMsg
    StatusBar1.Panels(3).Text = ""
    StatusBar1.Panels(4).Text = ""
    
    CurrUser.USER_NAME = ""
    CurrUser.USER_PK = 0
    
    frmLogin.show vbModal: If CloseMe = True Then Unload Me: Exit Sub: Exit Sub
    DisplayUserInfo
    'UpdateInfoMsg
End Sub

Private Sub mnuHUG_Click()
    '
End Sub

Private Sub mnuPChS_Click()
  ReportType = "Periodic Charge Sales"
  frmPreview.show 1
End Sub

Private Sub mnuPChSI_Click()
  ReportType = "Periodic Charge Sales (Itemized)"
  frmPreview.show 1
End Sub

Private Sub mnuPChSR_Click()
  ReportType = "Periodic Charge Sales Returns"
  frmPreview.show 1
End Sub

Private Sub mnuPChSRI_Click()
  ReportType = "Periodic Charge Sales Returns (Itemized)"
  frmPreview.show 1
End Sub

Private Sub mnuPCSI_Click()
    ReportType = "Periodic Cash Sales Itemized"
    frmPreview.show 1
End Sub

Private Sub mnuPCSR_Click()
  ReportType = "Periodic Cash Sales Return"
  frmPreview.show 1
End Sub

Private Sub mnuPCSRI_Click()
  ReportType = "Periodic Cash Sales Return (Itemized)"
  frmPreview.show 1
End Sub

Private Sub mnuPrintBarcode_Click()
  frmPrintBarcode.show 1
End Sub

Private Sub mnuRAC_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Close"
End Sub

Private Sub mnuRACN_Click()
    On Error Resume Next
    ActiveForm.CommandPass "New"
End Sub

Private Sub mnuRADS_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Delete"
End Sub

Private Sub mnuRAES_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Edit"
End Sub

Private Sub mnuRAP_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Print"
End Sub

Private Sub mnuRARR_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Refresh"
End Sub

Private Sub mnuRAS_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Search"
End Sub

Private Sub mnuRASSM_Click()
    frmShortcuts.show
    frmShortcuts.WindowState = vbMaximized
    frmShortcuts.SetFocus
End Sub

Private Sub mnuSMMU_Click()
    If CurrUser.USER_ISADMIN = False Then
        MsgBox "Only admin users can access this record.", vbCritical, "Access Denied"
    Else
        frmUserRec.show vbModal
    End If
End Sub

Private Sub mnuSOAD_Click()
    frmShortcuts.lvMenu.View = lvwIcon
End Sub

Private Sub mnuSOAHL_Click()
    frmShortcuts.lvMenu.View = lvwSmallIcon
End Sub

Private Sub mnuSOAVL_Click()
    frmShortcuts.lvMenu.View = lvwList
End Sub


Private Sub mnuSSS_Click()
    frmSplash.DisableLoader = True
    frmSplash.show vbModal
End Sub

Private Sub mnuStockCard_Click()
  frmStockCard.show 1
End Sub

Private Sub mnuUBI_Click()
    frmBusinessInfo.show vbModal
End Sub

Private Sub mnuUC_Click()
    On Error Resume Next
    Shell "calc.exe", vbNormalFocus
End Sub

Private Sub mnuUN_Click()
    On Error Resume Next
    Shell "notepad.exe", vbNormalFocus
End Sub

Private Sub mnuUWE_Click()
    On Error Resume Next
    Shell "Explorer.exe", vbNormalFocus
End Sub

Private Sub StyleButton2_Click()
    show_mnu = Not show_mnu
    show_menu (show_mnu)
End Sub

Private Sub show_menu(ByVal show As Boolean)
    Dim img As Image
    If show = True Then
        Set img = Image2
    Else
        Set img = Image5
    End If
    'Set the style button graphics
    With StyleButton2
        Set .PictureDown = img.Picture
        Set .PictureFocus = img.Picture
        Set .PictureHover = img.Picture
        Set .PictureUp = img.Picture
    End With
    'Set picture visibility
    picLeft.Visible = show
    
    If show = True Then StyleButton2.ToolTipText = "Hide": picSeparator.MousePointer = vbSizeWE Else picSeparator.MousePointer = vbArrow: StyleButton2.ToolTipText = "Expand"
    
    Set img = Nothing
End Sub

Private Sub picSeparator_Resize()
    Call center_obj_vertical(picSeparator, StyleButton2)
End Sub

Public Sub HideTBButton(ByVal srcPatern As String, Optional srcAllButton As Boolean)
    If srcAllButton = True Then srcPatern = "ttttttt"
    If Mid$(srcPatern, 1, 1) = "t" Then tbMenu.Buttons(3).Visible = False: mnuRACN.Visible = False
    If Mid$(srcPatern, 2, 1) = "t" Then tbMenu.Buttons(4).Visible = False: mnuRAES.Visible = False
    If Mid$(srcPatern, 3, 1) = "t" Then tbMenu.Buttons(5).Visible = False: mnuRAS.Visible = False
    If Mid$(srcPatern, 4, 1) = "t" Then tbMenu.Buttons(6).Visible = False: mnuRADS.Visible = False
    If Mid$(srcPatern, 5, 1) = "t" Then tbMenu.Buttons(7).Visible = False: mnuRARR.Visible = False
    If Mid$(srcPatern, 6, 1) = "t" Then tbMenu.Buttons(8).Visible = False: mnuRAP.Visible = False
    If Mid$(srcPatern, 7, 1) = "t" Then tbMenu.Buttons(9).Visible = False: mnuRAC.Visible = False
    If mnuRAC.Visible = False Then mnuRASep2.Visible = False
End Sub

Public Sub ShowTBButton(ByVal srcPatern As String, Optional srcAllButton As Boolean)
    'Highligh active form in opened form list
    If srcAllButton = True Then srcPatern = "ttttttt"
    If Mid$(srcPatern, 1, 1) = "t" Then tbMenu.Buttons(3).Visible = True: mnuRACN.Visible = True
    If Mid$(srcPatern, 2, 1) = "t" Then tbMenu.Buttons(4).Visible = True: mnuRAES.Visible = True
    If Mid$(srcPatern, 3, 1) = "t" Then tbMenu.Buttons(5).Visible = True: mnuRAS.Visible = True
    If Mid$(srcPatern, 4, 1) = "t" Then tbMenu.Buttons(6).Visible = True: mnuRADS.Visible = True
    If Mid$(srcPatern, 5, 1) = "t" Then tbMenu.Buttons(7).Visible = True: mnuRARR.Visible = True
    If Mid$(srcPatern, 6, 1) = "t" Then tbMenu.Buttons(8).Visible = True: mnuRAP.Visible = True
    If Mid$(srcPatern, 7, 1) = "t" Then tbMenu.Buttons(9).Visible = True: mnuRAC.Visible = True
    If mnuRAC.Visible = True Then mnuRASep2.Visible = True
End Sub

Public Sub ShowMe()
    Me.Visible = True
End Sub

Private Sub MDIForm_Load()
    show
    Me.BackColor = &H80000005
    HideTBButton "", True
    frmShortcuts.show
    
    'DBPath = App.Path & "\Data\DB.mdb"
    
    DBPath = GetINI("Inventory Settings", "Path")      'get path from file
    If Trim(DBPath) = "" Or IsNull(DBPath) Then
JumpHere:
        frmLocate.show 1                            'browse database
    End If
     
    'If OpenDB = False Then CloseMe = True: Unload Me: Exit Sub
    If OpenDB = vbRetry Then GoTo JumpHere
    
    createDSN
    
    frmLogin.show vbModal: If CloseMe = True Then Unload Me: Exit Sub: Exit Sub
    
    'Set the control properties
    Set lvWin.SmallIcons = i16x16
    Set lvWin.Icons = i16x16
    
    DisplayUserInfo
    
    lvWin.ListItems.Add(, "frmShortcuts", "@Shortcuts", 1, 1).Bold = True
    
    show_mnu = True
    show_menu (show_mnu)
End Sub

Private Sub DisplayUserInfo()
    'Display the current user info
    If CurrUser.USER_ISADMIN = True Then
        StatusBar1.Panels(4).Text = "Admin"
    Else
        StatusBar1.Panels(4).Text = "Operator"
    End If
    StatusBar1.Panels(3).Text = CurrUser.USER_NAME
    
    Dim rs As New Recordset
    
    rs.Open "SELECT * FROM [Business Info]", CN, adOpenStatic, adLockReadOnly
    
    CurrBiz.BUSINESS_NAME = rs.Fields(0)
    CurrBiz.BUSINESS_ADDRESS = rs.Fields(1)
    CurrBiz.BUSINESS_CONTACT_INFO = rs.Fields(2)
    
    Set rs = Nothing
End Sub

Public Sub AddToWin(ByVal srcDName As String, ByVal srcFormName As String)
    On Error Resume Next
    Dim xItem As ListItem
    
    Set xItem = lvWin.ListItems.Add(, srcFormName, srcDName, 1, 1)
    xItem.ToolTipText = srcDName
    xItem.SubItems(1) = "***" & srcDName & "***"
    xItem.Selected = True
    
    Set xItem = Nothing
End Sub

Public Sub RemToWin(ByVal srcDName As String)
    On Error Resume Next
    search_in_listview lvWin, "***" & srcDName & "***"
    lvWin.ListItems.Remove (lvWin.SelectedItem.Index)
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    Set MAIN = Nothing
End Sub

Private Sub mnuA_Click()
    frmAbout.show vbModal
End Sub

Private Sub mnuHKS_Click()
    'AddTest
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    Frame1.Width = picLeft.ScaleWidth
    lvWin.Width = picLeft.ScaleWidth
    lvWin.Height = picLeft.ScaleHeight - lvWin.Top - 20
End Sub

Private Sub picSeparator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If show_mnu = False Then Exit Sub
    If Button = vbLeftButton Then
        tmrResize.Enabled = True
        resize_down = True
    End If
End Sub

Private Sub picSeparator_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If show_mnu = False Then Exit Sub
    If Button = vbLeftButton Then
        tmrResize.Enabled = False
        resize_down = False
    End If
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Shortcuts" Then
        frmShortcuts.show
        frmShortcuts.WindowState = vbMaximized
        frmShortcuts.SetFocus
    Else
        On Error Resume Next
        ActiveForm.CommandPass Button.Key
    End If
End Sub

Private Sub tmrResize_Timer()
    On Error Resume Next
    GetCursorPos cursor_pos
    picLeft.Width = (Me.Width - ((cursor_pos.X * Screen.TwipsPerPixelX) - Me.Left)) - 90
End Sub

Private Sub tmrMemStatus_Timer()
    Call GlobalMemoryStatus(MEM_STAT)
    lblPMem.Caption = Format((MEM_STAT.dwAvailPhys / 1024) / 1024, "#,##0.0") & " MB"
    lblVMem.Caption = Format((MEM_STAT.dwAvailVirtual / 1024) / 1024, "#,##0.0") & " MB"
End Sub

Public Sub UnloadChilds()
''Unload all active forms
Dim Form As Form
   For Each Form In Forms
      ''Unload all active childs
      If Form.Name <> Me.Name And Form.Name <> "frmShortcuts" Then Unload Form
   Next Form
   
Set Form = Nothing
End Sub
