VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "aDe's TTD Saver - Configure"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReadMe 
      Cancel          =   -1  'True
      Caption         =   "ReadMe!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   18
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   17
      Top             =   540
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   3735
      Begin VB.ComboBox cmbTileset 
         Height          =   315
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   270
         Width           =   3375
      End
      Begin VB.CommandButton CmdExportTilesets 
         Caption         =   "Export default tileset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkEnableCustomTilesets 
         Caption         =   "Custom tileset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   1455
      End
      Begin VB.CheckBox chkRandomTileSet 
         Caption         =   "Use random tileset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Random"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   1305
         TabIndex        =   26
         Top             =   1800
         Width           =   2175
         Begin VB.OptionButton optTrackDrawDirect 
            Caption         =   "Progressive"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   975
            TabIndex        =   28
            Top             =   0
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optTrackDrawDirect 
            Caption         =   "Direct"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   870
         End
      End
      Begin VB.OptionButton optTrainStyle 
         Caption         =   "Semi-uniform"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2265
         TabIndex        =   22
         Top             =   1185
         Width           =   1260
      End
      Begin VB.OptionButton optTrainStyle 
         Caption         =   "Random"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1305
         TabIndex        =   20
         Top             =   1185
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtSmStruct 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   13
         Text            =   "100"
         Top             =   870
         Width           =   495
      End
      Begin VB.TextBox txtSmStruct 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   11
         Text            =   "30"
         Top             =   870
         Width           =   495
      End
      Begin VB.TextBox txtBigStruct 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Text            =   "150"
         Top             =   555
         Width           =   495
      End
      Begin VB.TextBox txtBigStruct 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   7
         Text            =   "50"
         Top             =   555
         Width           =   495
      End
      Begin VB.TextBox txtTrainLen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   5
         Text            =   "32"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtTrainLen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optTrainStyle 
         Caption         =   "Uniform"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1305
         TabIndex        =   21
         Top             =   1425
         Width           =   855
      End
      Begin VB.Label lblGui 
         Caption         =   "Track draw:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblGui 
         Caption         =   "Train style:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   870
         Width           =   255
      End
      Begin VB.Label lblGui 
         Caption         =   "Small structures:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   555
         Width           =   255
      End
      Begin VB.Label lblGui 
         Caption         =   "Large structures:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   555
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblGui 
         Caption         =   "Train length:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   120
      Picture         =   "frmSettings.frx":0E42
      Top             =   120
      Width           =   3750
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
End
End Sub

Public Sub CmdExportTilesets_Click()
On Error GoTo nd:
Dim newCfg As New adeSettings, sOutDir As String
sOutDir = appInstallDir & "tilesets\default\"

newCfg.AutoSave = False
newCfg.Save "version", CURTILESETVERSION
newCfg.Save "buildings_big.bmp", "8", "imagecounts"
newCfg.Save "buildings_med.bmp", "16", "imagecounts"
newCfg.Save "double", "0", "engine0"

For i = 0 To frmMain.ImgTS.UBound
  SavePicture frmMain.ImgTS(i).Picture, sOutDir & frmMain.ImgTS(i).Tag
Next

newCfg.SaveFile sOutDir & "tileset.txt"

'/e = installation!
If Command <> "/e" Then MsgBox "Tileset extraction successfull", vbInformation
Exit Sub
nd:
MsgBox "Error extracting tilesets. Make sure directory exists...(" & sOutDir & ")", vbCritical
End Sub

Private Sub cmdOK_Click()
Settings.iBldngBigMin = txtBigStruct(0).Text
Settings.iBldngBigMax = txtBigStruct(1).Text
Settings.iBldngSmMin = txtSmStruct(0).Text
Settings.iBldngSmMax = txtSmStruct(1).Text
Settings.iEnableCustomTilesets = chkEnableCustomTilesets.Value
Settings.iTrainLenMin = txtTrainLen(0).Text
Settings.iTrainLenMax = txtTrainLen(1).Text
Settings.iTrainStyleUniform = CInt(optTrainStyle(1).Value)
Settings.iTrainStyleSemiUniform = CInt(optTrainStyle(2).Value)
Settings.iRandomTileSet = chkRandomTileSet.Value
Settings.iDirectTrackDraw = CInt(optTrackDrawDirect(0).Value)

Settings.sTileSet = cmbTileset.Text

SaveSettings

End
End Sub


Private Sub cmdReadMe_Click()
frmReadMe.Show

End Sub

Private Sub Form_Load()
Dim ListIndx As Integer, i As Integer
txtBigStruct(0).Text = Settings.iBldngBigMin
txtBigStruct(1).Text = Settings.iBldngBigMax
txtSmStruct(0).Text = Settings.iBldngSmMin
txtSmStruct(1).Text = Settings.iBldngSmMax
chkEnableCustomTilesets.Value = Settings.iEnableCustomTilesets
txtTrainLen(0).Text = Settings.iTrainLenMin
txtTrainLen(1).Text = Settings.iTrainLenMax
chkRandomTileSet.Value = Settings.iRandomTileSet
optTrainStyle(1).Value = CBool(Settings.iTrainStyleUniform)
optTrainStyle(2).Value = CBool(Settings.iTrainStyleSemiUniform)
optTrackDrawDirect(0).Value = CBool(Settings.iDirectTrackDraw)

frmMain.Dir1.Path = appInstallDir & "tilesets"
If frmMain.Dir1.ListCount = 0 Then
  chkEnableCustomTilesets.Value = 0
  chkEnableCustomTilesets.Enabled = False
Else
  For i = 0 To frmMain.Dir1.ListCount - 1
    cmbTileset.AddItem AfterLast(frmMain.Dir1.List(i), "\")
    If Settings.sTileSet = AfterLast(frmMain.Dir1.List(i), "\") Then ListIndx = i
  Next
  If ListIndx Then
    cmbTileset.ListIndex = ListIndx
  Else
    cmbTileset.ListIndex = 0
  End If
End If
End Sub

