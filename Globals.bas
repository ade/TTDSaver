Attribute VB_Name = "Globals"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Const HWND_TOPMOST As Long = -1
Public Const SWP_SHOWWINDOW As Long = &H40
Public Const SWP_NOACTIVATE As Long = &H10
Public Const GWL_STYLE As Long = (-16)
Public Const WS_CHILD As Long = &H40000000
Public Const GWL_HWNDPARENT As Long = (-8)
Public Const HWND_TOP As Long = 0
Public Const SWP_NOZORDER = &H4

Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Global PreviewHwnd As Long
Global PreviewRect As RECT

Global appInstallDir As String

Global Const GRND_EARTH = 0
Global Const GRND_GRASS = 1
Global Const GRND_ROAD_E = 2
Global Const GRND_ROAD_ES = 3
Global Const GRND_ROAD_WE = 4
Global Const GRND_ROAD_N = 5
Global Const GRND_ROAD_NE = 6
Global Const GRND_ROAD_NS = 7
Global Const GRND_ROAD_NW = 8
Global Const GRND_ROAD_S = 9
Global Const GRND_ROAD_W = 10
Global Const GRND_ROAD_WS = 11
Global Const GRND_ROAD_RRX_NS = 12
Global Const GRND_ROAD_RRX_EW = 13
Global Const GRND_ROAD_X = 14
Global Const GRND_TRACK_ES = 16
Global Const GRND_TRACK_WE = 17
Global Const GRND_TRACK_NE = 18
Global Const GRND_TRACK_NS = 19
Global Const GRND_TRACK_NW = 20
Global Const GRND_TRACK_WS = 21
Global Const GRND_TRACK_X = 22

Global Const TS_GRND = 0
Global Const TS_BUILDINGS_BIG = 1
Global Const TS_BUILDINGS_MED = 2

Global Const CURTILESETVERSION = 1

Type trainInfo
  iTileSet As Integer
  iTileNum As Integer
  iDir As Integer
  iSpeed As Integer
  iNode As Integer
  iNodePos As Integer
  bReversed As Boolean
End Type

Type tsINF
  w As Long
  h As Long
End Type

Type cellinf
  X As Integer
  Y As Integer
  Free As Boolean
  TileSet As Integer
  Overlay As Integer
End Type

Type tSettings
  iBldngBigMax As Integer
  iBldngBigMin As Integer
  iBldngSmMax As Integer
  iBldngSmMin As Integer
  iTrainLenMin As Integer
  iTrainLenMax As Integer
  iEnableCustomTilesets As Integer
  sTileSet As String
  iTrainStyleUniform As Integer
  iTrainStyleSemiUniform As Integer
  iRandomTileSet As Integer
  iDirectTrackDraw As Integer
End Type

Global Train(0 To 32) As trainInfo
Global cellInfo(0 To 808) As cellinf
Global Settings As tSettings
Global tsInfo(0 To 4) As tsINF


Sub SaveSettings()
SaveSetting "aDe", "TTDSaver", "iBldngBigMax", Settings.iBldngBigMax
SaveSetting "aDe", "TTDSaver", "iBldngBigMin", Settings.iBldngBigMin
SaveSetting "aDe", "TTDSaver", "iBldngSmMax", Settings.iBldngSmMax
SaveSetting "aDe", "TTDSaver", "iBldngSmMin", Settings.iBldngSmMin
SaveSetting "aDe", "TTDSaver", "iEnableCustomTilesets", Settings.iEnableCustomTilesets
SaveSetting "aDe", "TTDSaver", "iTrainLenMax", Settings.iTrainLenMax
SaveSetting "aDe", "TTDSaver", "iTrainLenMin", Settings.iTrainLenMin
SaveSetting "aDe", "TTDSaver", "sTileSet", Settings.sTileSet
SaveSetting "aDe", "TTDSaver", "iRandomTileSet", Settings.iRandomTileSet
SaveSetting "aDe", "TTDSaver", "iTrainStyleUniform", Settings.iTrainStyleUniform
SaveSetting "aDe", "TTDSaver", "iTrainStyleSemiUniform", Settings.iTrainStyleSemiUniform
SaveSetting "aDe", "TTDSaver", "iDirectTrackDraw", Settings.iDirectTrackDraw
End Sub

Sub LoadSettings()
appInstallDir = Reg_GetString(HKEY_LOCAL_MACHINE, "Software\ade.se\ttdsaver", "InstallDir")
If Len(appInstallDir) Then
  If Right(appInstallDir, 1) <> "\" Then appInstallDir = appInstallDir & "\"
Else
  MsgBox "Installation directory could not be located, please reinstall me!", vbCritical
  End
End If

Settings.iBldngBigMax = GetSetting("aDe", "TTDSaver", "iBldngBigMax", 150)
Settings.iBldngBigMin = GetSetting("aDe", "TTDSaver", "iBldngBigMin", 25)
Settings.iBldngSmMax = GetSetting("aDe", "TTDSaver", "iBldngSmMax", 70)
Settings.iBldngSmMin = GetSetting("aDe", "TTDSaver", "iBldngSmMin", 350)
Settings.iEnableCustomTilesets = GetSetting("aDe", "TTDSaver", "iEnableCustomTilesets", 0)
Settings.iTrainLenMax = GetSetting("aDe", "TTDSaver", "iTrainLenMax", 16)
Settings.iTrainLenMin = GetSetting("aDe", "TTDSaver", "iTrainLenMin", 2)
Settings.sTileSet = GetSetting("aDe", "TTDSaver", "sTileSet", "default")
Settings.iTrainStyleUniform = GetSetting("aDe", "TTDSaver", "iTrainStyleUniform", 0)
Settings.iTrainStyleSemiUniform = GetSetting("aDe", "TTDSaver", "iTrainStyleSemiUniform", 1)
Settings.iRandomTileSet = GetSetting("aDe", "TTDSaver", "iRandomTileSet", 0)
Settings.iDirectTrackDraw = GetSetting("aDe", "TTDSaver", "iDirectTrackDraw", 0)
If Settings.iEnableCustomTilesets Then frmMain.tileSetCFG.ReadFile appInstallDir & "\tilesets\" & Settings.sTileSet & "\tilesets.txt"
End Sub

Public Function BeforeFirst(sIn, sFirst)
    If InStr(1, sIn, sFirst) Then
        BeforeFirst = Left(sIn, InStr(1, sIn, sFirst) - 1)
    Else
        BeforeFirst = ""
    End If
End Function

Public Function AfterFirst(sIn, sFirst)
    If InStr(1, sIn, sFirst) Then
        AfterFirst = Right(sIn, Len(sIn) - InStr(1, sIn, sFirst) - (Len(sFirst) - 1))
    Else
        AfterFirst = ""
    End If
End Function

Public Function AfterLast(sFrom, sAfterLast)
    If InStr(1, sFrom, sAfterLast) Then
        AfterLast = Right(sFrom, Len(sFrom) - InStrRev(sFrom, sAfterLast) - (Len(sAfterLast) - 1))
    Else
        AfterLast = ""
    End If
End Function

Public Function BeforeLast(sFrom, sBeforeLast)
    If InStr(1, sFrom, sBeforeLast) Then
        BeforeLast = Left(sFrom, InStrRev(sFrom, sBeforeLast) - 1)
    Else
        BeforeLast = ""
    End If
End Function

