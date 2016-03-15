VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10485
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0E42
   ScaleHeight     =   7830
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.PictureBox picSize 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Timer tmrDrawTrack 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   1800
      Top             =   1800
   End
   Begin VB.Timer tmrDraw 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   1320
      Top             =   1800
   End
   Begin VB.Image ImgTS 
      Height          =   3840
      Index           =   2
      Left            =   240
      Picture         =   "frmMain.frx":D604
      Tag             =   "buildings_med.bmp"
      Top             =   3720
      Width           =   3840
   End
   Begin VB.Image ImgTS 
      Height          =   1920
      Index           =   1
      Left            =   2520
      Picture         =   "frmMain.frx":3D646
      Tag             =   "buildings_big.bmp"
      Top             =   1680
      Width           =   7680
   End
   Begin VB.Image ImgTS 
      Height          =   1080
      Index           =   4
      Left            =   4200
      Picture         =   "frmMain.frx":6D688
      Tag             =   "wagons.bmp"
      Top             =   4320
      Width           =   3360
   End
   Begin VB.Image ImgTS 
      Height          =   540
      Index           =   3
      Left            =   4200
      Picture         =   "frmMain.frx":793CA
      Tag             =   "engines.bmp"
      Top             =   3720
      Width           =   3360
   End
   Begin VB.Image ImgTS 
      Height          =   1440
      Index           =   0
      Left            =   2520
      Picture         =   "frmMain.frx":7F28C
      Tag             =   "ground.bmp"
      Top             =   120
      Width           =   7680
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Welcome to the source code of my TTD saver
' I commented here & there... hope you get something out of this :)
' If you make a recompile, please include my original credit and web address.
' And try to make it a little different so people don't mess your version up with
' the official one (make your own section in the readme form)...
'
' ade
' http://ade.se

Public WithEvents bDisp As BMDXDisplay
Attribute bDisp.VB_VarHelpID = -1
Dim objTileSet(0 To 2) As IBMDXTileSet
Dim tsTrain(0 To 1) As IBMDXTileSet
Dim bDraw As Boolean
Dim cellImg(0 To 809) As Integer
Dim AIPath() As Integer, AIOrigo() As Integer
Dim EdgeCells(0 To 83) As Integer
Dim sMsg As String, AINodes As Integer
Dim numWagons As Integer
Public tileSetCFG As adeSettings
Dim curDrawTrack As Integer 'used by tmrDrawTrack

Private Sub bDisp_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)

If Button = 2 Then End
StartSaver

End Sub

Sub StartSaver()
Dim upperBound As Integer, lowerBound As Integer, iCar As Integer
If Settings.iRandomTileSet Then LoadImages
tmrDraw.Enabled = True
upperBound = Settings.iTrainLenMax
lowerBound = Settings.iTrainLenMin
numWagons = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
Train(0).iTileSet = 0
Train(0).iTileNum = Int(Rnd * tileSetCFG.Read(ImgTS(3).Tag, "1", "trains"))

For i = 1 To UBound(Train) 'set the rest of the cars to wagons tileset
  Train(i).iTileSet = 1
Next

'now generate the different types of cars
If Settings.iTrainStyleUniform + Settings.iTrainStyleSemiUniform = 0 Then 'random train style
  For i = 1 To UBound(Train)
    Train(i).iTileNum = Int(Rnd * tileSetCFG.Read(ImgTS(4).Tag, "4", "trains"))
  Next
ElseIf Settings.iTrainStyleUniform Then 'uniform train style
  iCar = Int(Rnd * tileSetCFG.Read(ImgTS(4).Tag, "4", "trains"))
  For i = 1 To UBound(Train)
    Train(i).iTileNum = iCar
  Next
Else ' semi-uniform train style
  iCar = Int(Rnd * tileSetCFG.Read(ImgTS(4).Tag, "4", "trains"))
  For i = 1 To UBound(Train)
    If Int(Rnd * 5) = 0 Then iCar = Int(Rnd * tileSetCFG.Read(ImgTS(4).Tag, "4", "trains")) 'new type
    Train(i).iTileNum = iCar
  Next
End If

'double headed train?
If Val(tileSetCFG.Read("double", "0", "engine" & Train(0).iTileNum)) = 1 Then
  Train(numWagons).bReversed = True
  Train(numWagons).iTileSet = 0
  Train(numWagons).iTileNum = Train(0).iTileNum
End If

For i = 0 To UBound(Train)
  Train(i).iNode = -2
Next

ResetCells
AI_MakePath
RandomBuildings

For i = 0 To numWagons
  Train(i).iNode = -1
Next

curDrawTrack = 0
tmrDrawTrack.Enabled = True
tmrDrawTrack_Timer ' call the timer to avoid making the track flash in when using direct mode

End Sub

Sub RandomBuildings()
Dim itmCount As Integer
Dim itmCell As Integer, upperBound As Integer, lowerBound As Integer
Randomize
'Large structures
lowerBound = Settings.iBldngBigMin
upperBound = Settings.iBldngBigMax
itmCount = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
'itmCount = Int(Rnd * 50) + 1
For i = 0 To itmCount
newcell:
  itmCell = Int(Rnd * 808)
  If cellInfo(itmCell).Free Then
    cellInfo(itmCell).Overlay = Int(Rnd * tileSetCFG.Read(ImgTS(TS_BUILDINGS_BIG).Tag, "8", "imagecounts"))
    cellInfo(itmCell).TileSet = TS_BUILDINGS_BIG
  Else
    GoTo newcell
  End If
Next

'Medium/small structures
lowerBound = Settings.iBldngSmMin
upperBound = Settings.iBldngSmMax
itmCount = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
'itmCount = Int(Rnd * 350) + 40
For i = 0 To itmCount
newcell2:
  itmCell = Int(Rnd * 808)
  If cellInfo(itmCell).Free Then
    cellInfo(itmCell).Overlay = Int(Rnd * tileSetCFG.Read(ImgTS(TS_BUILDINGS_MED).Tag, "16", "imagecounts"))
    cellInfo(itmCell).TileSet = TS_BUILDINGS_MED
  Else
    GoTo newcell2
  End If
Next

End Sub

Private Sub Form_Load()
Dim Style As Long
Height = 1680
Width = 2280
Set bDisp = New BMDXDisplay
Set tileSetCFG = New adeSettings

LoadSettings
RunLicense

Select Case LCase(Left(Command, 2))
  Case "/s" 'start
    OpenWin
    SetEdgeCells
    LoadImages
    ResetCells
    StartSaver
    bDraw = True
  Case "/p" 'preview
    ' Do some funky stuff to show little preview window in Desktop properties page.
    PreviewHwnd = Val(Mid(Command, 4))
    GetClientRect PreviewHwnd, PreviewRect
    Style = GetWindowLong(Me.hwnd, GWL_STYLE)
    Style = Style Or WS_CHILD
    SetWindowLong Me.hwnd, GWL_STYLE, Style
    SetParent Me.hwnd, PreviewHwnd
    SetWindowLong Me.hwnd, GWL_HWNDPARENT, PreviewHwnd
    SetWindowPos Me.hwnd, HWND_TOP, 0&, 0&, PreviewRect.Right, PreviewRect.Bottom, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW ' SWP_NOACTIVATE Or SWP_SHOWWINDOW
    DoEvents
    Me.Visible = True
  Case "/c" 'configure
    frmSettings.Show
  Case "/d" 'debug
    OpenWin
    SetEdgeCells
    LoadImages
    ResetCells
    bDraw = True
  Case "/e" 'installation.. extract tilesets & stuff
    frmSettings.CmdExportTilesets_Click
    End
  Case Else
    MsgBox "You have started me the wrong way. I'm a screensaver. Try with /s or /c !", vbCritical
    End
End Select
End Sub
Sub RunLicense()
On Error Resume Next
bDisp.ValidateLicense "TeeIJLoMywzxCwoIIIGqoMtx"
bDisp.ValidateLicense "TeeIJLyvUxPgXxBPIrOGKyRy"
End Sub

Sub SetEdgeCells()
'this sub pre-calculates which cells are on the edge of the map, to save CPU later.
For i = 0 To 16 ' top
EdgeCells(i) = i
Next

For i = 0 To 16 ' 17 to 32  - bottom
EdgeCells(i + 17) = 792 + i
Next

'0,33,66.. left side.
For i = 0 To 24 ' 33 to 58
EdgeCells(i + 33) = i * 33
Next

'16, 49, 82, 115... right side.
For i = 0 To 24 '59 to 83
EdgeCells(i + 59) = 16 + (i * 33)
Next

End Sub

Sub LoadImages()
Dim tilePic() As StdPicture, tsDir As String, i As Integer
ReDim tilePic(0 To 5)
Randomize

'detect where to get the images from and load them into memory
If Settings.iEnableCustomTilesets Then
  If Settings.iRandomTileSet Then
    Dir1.Path = appInstallDir & "tilesets\"
    If Dir1.ListCount Then
      tsDir = Dir1.List(Int(Rnd * Dir1.ListCount)) & "\"
    Else
      MsgBox "Error: No custom tilesets found", vbCritical
      End
    End If
  Else
    tsDir = appInstallDir & "tilesets\" & Settings.sTileSet & "\"
  End If
  For i = 0 To 4
    Set tilePic(i) = LoadPicture(tsDir & ImgTS(i).Tag)
  Next
  tileSetCFG.ReadFile tsDir & "tileset.txt", True
Else
  For i = 0 To 4
    Set tilePic(i) = ImgTS(i).Picture
  Next
End If

'now create compatible tilesets for the directx object
If tsInfo(0).h = 0 Then
  Set objTileSet(0) = bDisp.CreateTileSet(tilePic(0), 64, 32) 'ground
  Set objTileSet(1) = bDisp.CreateTileSet(tilePic(1), 64, 128) 'lg buildings
  Set objTileSet(2) = bDisp.CreateTileSet(tilePic(2), 64, 64) 'med/sm buildings
  Set tsTrain(0) = bDisp.CreateTileSet(tilePic(3), 28, 18) ' engines
  Set tsTrain(1) = bDisp.CreateTileSet(tilePic(4), 28, 18) ' wagons
  For i = 0 To 4
    tsInfo(i).h = tilePic(i).Height
    tsInfo(i).w = tilePic(i).Width
  Next
Else
  'when reloading images, this is the way to do it...
  objTileSet(0).PaintPicture tilePic(0), 0, 0, bDisp.HimetricToPixelX(tilePic(0).Width), bDisp.HimetricToPixelY(tilePic(0).Height), 0, 0, tilePic(0).Width, tilePic(0).Height
  objTileSet(1).PaintPicture tilePic(1), 0, 0, bDisp.HimetricToPixelX(tilePic(1).Width), bDisp.HimetricToPixelY(tilePic(1).Height), 0, 0, tilePic(1).Width, tilePic(1).Height
  objTileSet(2).PaintPicture tilePic(2), 0, 0, bDisp.HimetricToPixelX(tilePic(2).Width), bDisp.HimetricToPixelY(tilePic(2).Height), 0, 0, tilePic(2).Width, tilePic(2).Height
  tsTrain(0).PaintPicture tilePic(3), 0, 0, bDisp.HimetricToPixelX(tilePic(3).Width), bDisp.HimetricToPixelY(tilePic(3).Height), 0, 0, tilePic(3).Width, tilePic(3).Height
  tsTrain(1).PaintPicture tilePic(4), 0, 0, bDisp.HimetricToPixelX(tilePic(4).Width), bDisp.HimetricToPixelY(tilePic(4).Height), 0, 0, tilePic(4).Width, tilePic(4).Height
End If

'detect sizes and wagon/engine counts and remake tilesets if they are bigger then previous
picSize.Picture = tilePic(3)
tileSetCFG.Save ImgTS(3).Tag, CStr(picSize.ScaleHeight / 18), "trains" 'detect how many engine types there are...
If tsInfo(3).h < picSize.Height Then
  For i = 0 To bDisp.TileSets.Count - 1
    If bDisp.TileSets(i) Is tsTrain(0) Then
      bDisp.TileSets.Remove i
      Exit For
     End If
  Next
  
  Set tsTrain(0) = bDisp.CreateTileSet(tilePic(3), 28, 18)
  tsInfo(3).h = picSize.Height
  tsInfo(3).w = picSize.Width
End If

picSize.Picture = tilePic(4)
tileSetCFG.Save ImgTS(4).Tag, CStr(picSize.ScaleHeight / 18), "trains" 'detect how many wagon types there are...

End Sub
Private Sub bDisp_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = vbKeyEscape Then End
If KeyCode = vbKeyP Then tmrDraw.Enabled = Not tmrDraw.Enabled
End Sub
Sub OpenWin()

bDisp.OpenEx 1024, 768


End Sub
Sub ResetCells()

For i = 0 To 808
  cellImg(i) = 1
  cellInfo(i).Free = True
  cellInfo(i).TileSet = 0
Next

End Sub

Private Sub tmrDraw_Timer()
UpdateTrains

If Not bDraw Then Exit Sub
Dim iRow As Integer, iCol As Integer, bShift As Boolean, dX As Integer, dY As Integer
Dim cellNum As Integer
Dim tOffSetY As Integer, tOffSetX As Integer, curTrain As Integer
Dim fStart As Integer, fEnd As Integer, fStep As Integer

bDisp.ForeColor = vbWhite 'for writing text.

For iRow = 0 To 48
  bShift = Not bShift
  For iCol = 0 To 16
    If Not bShift And (iCol = 16) Then
    Else
    dX = (64 * iCol)
    If bShift Then dX = dX - 32
    dY = (16 * iRow) - 16
    If cellInfo(808).X = 0 Then
      cellInfo(cellNum).X = dX
      cellInfo(cellNum).Y = dY
    End If
    bDisp.DrawTile objTileSet(0), cellImg(cellNum), dX, dY, 0, 0, 1024, 768, True
    'bDisp.DrawText cellNum, dX + 16, dY + 8  ' makes a number in each cell for debugging
    cellNum = cellNum + 1
    End If
  Next
Next

If AINodes Then ' <-- if there is a track...
' paint train(s). this involves making alot of math on where to put the sprite on the
' track, depending on how far it came in the cell
Select Case Train(0).iDir
  Case 0, 1, 2, 3
    fStart = 0
    fEnd = UBound(Train)
    fStep = 1
  Case 4, 5, 6, 7
    fStart = UBound(Train)
    fEnd = 0
    fStep = -1
End Select

For curTrain = fStart To fEnd Step fStep
If Train(curTrain).iNode > -1 Then
Select Case Train(curTrain).iDir
  Case 0 'north
    tOffSetX = 38
    tOffSetY = 14
    If Train(curTrain).iTileSet <> 0 Then tOffSetY = tOffSetY - 1
    tOffSetX = tOffSetX - (Train(curTrain).iNodePos * 2)
    tOffSetY = tOffSetY - Train(curTrain).iNodePos
  Case 1 ' northeast
    If AIOrigo(Train(curTrain).iNode) = 3 Then 'coming from south
      tOffSetX = 44
      tOffSetY = 13
    Else 'coming from west (4)
      tOffSetX = 12
      tOffSetY = 13
    End If
    tOffSetY = tOffSetY - Train(curTrain).iNodePos * 2
  Case 2
    tOffSetX = 7
    tOffSetY = 13
    If Train(curTrain).iTileSet <> 0 Then tOffSetY = tOffSetY - 1
    tOffSetX = tOffSetX + (Train(curTrain).iNodePos * 2)
    tOffSetY = tOffSetY - Train(curTrain).iNodePos
  Case 3 ' -> southeast
    If AIOrigo(Train(curTrain).iNode) = 1 Then 'coming from north
      tOffSetX = 0
      tOffSetY = -3
    Else 'coming from west (4)
      tOffSetX = 0
      tOffSetY = 13
    End If
    tOffSetX = tOffSetX + (Train(curTrain).iNodePos * 4)
  Case 4
    tOffSetX = 7
    tOffSetY = -4
    If Train(curTrain).iTileSet <> 0 Then tOffSetY = tOffSetY - 1
    tOffSetX = tOffSetX + (Train(curTrain).iNodePos * 2)
    tOffSetY = tOffSetY + Train(curTrain).iNodePos
  Case 5 'southwest
    If AIOrigo(Train(curTrain).iNode) = 1 Then 'coming from north
      tOffSetX = 12
      tOffSetY = -3
    Else 'coming from east (2)
      tOffSetX = 44
      tOffSetY = -3
    End If
    tOffSetY = tOffSetY + Train(curTrain).iNodePos * 2
  Case 6
    tOffSetX = 36
    tOffSetY = -3
    If Train(curTrain).iTileSet <> 0 Then tOffSetY = tOffSetY - 1
    tOffSetX = tOffSetX - (Train(curTrain).iNodePos * 2)
    tOffSetY = tOffSetY + Train(curTrain).iNodePos
  Case 7 'northwest
    If AIOrigo(Train(curTrain).iNode) = 3 Then 'coming from south
      tOffSetX = 35
      tOffSetY = 13
    Else 'coming from east (2)
      tOffSetX = 35
      tOffSetY = -3
    End If
    tOffSetX = tOffSetX - (Train(curTrain).iNodePos * 4)
End Select

'for double headed trains, the last 'car' will be reversed
If Train(curTrain).bReversed Then
  Select Case Train(curTrain).iDir
    Case 6, 4
      tOffSetY = tOffSetY + 1
    Case 2, 0
      tOffSetY = tOffSetY - 1
  End Select
  bDisp.DrawTile tsTrain(Train(curTrain).iTileSet), InvDir2(Train(curTrain).iDir) + (8 * Train(curTrain).iTileNum), cellInfo(AIPath(Train(curTrain).iNode)).X + tOffSetX, cellInfo(AIPath(Train(curTrain).iNode)).Y + tOffSetY, 0, 0, 1024, 768, True
Else
  bDisp.DrawTile tsTrain(Train(curTrain).iTileSet), Train(curTrain).iDir + (8 * Train(curTrain).iTileNum), cellInfo(AIPath(Train(curTrain).iNode)).X + tOffSetX, cellInfo(AIPath(Train(curTrain).iNode)).Y + tOffSetY, 0, 0, 1024, 768, True
End If

End If
Next
End If

'draw overlay (structures etc)
For i = 0 To 808
  If cellInfo(i).TileSet <> 0 Then
    Select Case cellInfo(i).TileSet
      Case TS_BUILDINGS_BIG
        tOffSetX = 0
        tOffSetY = -97
      Case TS_BUILDINGS_MED
        tOffSetX = 0
        tOffSetY = -33
    End Select
    bDisp.DrawTile objTileSet(cellInfo(i).TileSet), cellInfo(i).Overlay, cellInfo(i).X + tOffSetX, cellInfo(i).Y + tOffSetY, 0, 0, 1024, 768, True
  End If
Next

'bDisp.DrawText sMsg & " - ttd_saver " & cellNum - 1 & " cells, " & Time & "    " & Timer, 5, 750
bDisp.Flip

End Sub

Sub UpdateTrains()
'this sub starts and moves the traincars.
If AINodes = 0 Then Exit Sub
Dim cTrain As Integer, bMove As Boolean

For cTrain = 0 To UBound(Train)
If Train(cTrain).iNode <> -2 Then ' -2 = no draw
  If Train(cTrain).iNode = -1 Then ' -1 = waiting to start
    If cTrain > 0 Then 'wagons
      If Train(cTrain - 1).iNode > 0 Then
          Train(cTrain).iNode = 0
          Train(cTrain).iNodePos = 0
          Train(cTrain).iDir = FindDir(AIOrigo(0), InvDir(AIOrigo(1)))
      Else
        If Train(cTrain - 1).iNodePos > 7 Then
          Train(cTrain).iNode = 0
          Train(cTrain).iNodePos = 0
          Train(cTrain).iDir = FindDir(AIOrigo(0), InvDir(AIOrigo(1)))
        End If
      End If
    Else 'engine
      Train(cTrain).iNode = 0
      Train(cTrain).iNodePos = 0
      Train(cTrain).iDir = FindDir(AIOrigo(0), InvDir(AIOrigo(1)))
    End If
    bDraw = True
    Exit Sub
  Else
    Train(cTrain).iNodePos = Train(cTrain).iNodePos + 2
  End If
    
  'detect wheter to advance to the next cell
  'it only moves half the distance on the halfsize cells
  Select Case Train(cTrain).iDir
    Case 1, 3, 5, 7
      If Train(cTrain).iNodePos > 7 Then bMove = True
    Case 0, 2, 4, 6
      If Train(cTrain).iNodePos > 15 Then bMove = True
  End Select
  
  If bMove Then
    Train(cTrain).iNode = Train(cTrain).iNode + 1
    If Train(cTrain).iNode >= AINodes Then
      Train(cTrain).iNode = -2
      If cTrain = numWagons Then StartSaver
      Exit Sub
    End If
    Train(cTrain).iDir = FindDir(AIOrigo(Train(cTrain).iNode), InvDir(AIOrigo(Train(cTrain).iNode + 1)))
    Train(cTrain).iNodePos = 0
  End If
End If
bMove = False

Next
End Sub

Function AI_MakePath(Optional bPaint As Boolean) As Boolean
'this makes the track
Dim iStart As Integer, iEnd As Integer, iTestCell As Integer, iCurNode As Integer
ReDim AIPath(800), AIOrigo(800)
Dim skipDir() As Boolean, rDir As Integer, badPath As Boolean, badCount As Integer
Dim hintDir As Integer
ReDim skipDir(4)

Randomize


'find an edge cell to start pathfinding with.
iStart = Int(Rnd * 83)
AIPath(0) = EdgeCells(iStart)
If bPaint Then cellImg(AIPath(0)) = 0
AIOrigo(0) = WhichEdge(iStart)
cellInfo(AIPath(0)).Free = False

'now find which corner it came out on and select an end-directon.
Select Case WhichEdge(iStart)
  Case 1 ' north
    iEnd = Int((32 - 17 + 1) * Rnd + 17) ' find a south cell
    hintDir = 3
    skipDir(1) = True
    skipDir(2) = True
  Case 2 ' east
    iEnd = Int((58 - 33 + 1) * Rnd + 33) ' find a west cell
    hintDir = 4
    skipDir(2) = True
    skipDir(3) = True
  Case 3 ' south
    iEnd = Int(Rnd * 17) 'find a north corner
    hintDir = 1
    skipDir(3) = True
    skipDir(4) = True
  Case 4 ' west
    iEnd = Int((83 - 59 + 1) * Rnd + 59) ' find a east cell
    hintDir = 2
    skipDir(4) = True
    skipDir(1) = True
End Select


If AIPath(0) = 16 Then
  skipDir(1) = True
  skipDir(2) = True
  skipDir(3) = True
End If

If AIPath(0) = 0 Then
  skipDir(1) = True
  skipDir(2) = True
  skipDir(4) = True
End If

If AIPath(0) = 792 Then
  skipDir(1) = True
  skipDir(3) = True
  skipDir(4) = True
End If

If AIPath(0) = 808 Then
  skipDir(1) = True
  skipDir(2) = True
  skipDir(3) = True
End If

' ok, now find a good cell to move the focus to...
Do
retry:
  sMsg = "path find running: node " & iCurNode & " (" & skipDir(1) & "/" & skipDir(2) & "/" & skipDir(3) & "/" & skipDir(4) & ")"
  DoEvents
  
  rDir = Int((Rnd * 4) + 1)
  
  If Not skipDir(hintDir) And (Int(Rnd * 2) = 1) Then rDir = hintDir
  If Not skipDir(1) Then: If AIPath(iCurNode) - 17 = EdgeCells(iEnd) Then rDir = 1
  If Not skipDir(2) Then: If AIPath(iCurNode) - 16 = EdgeCells(iEnd) Then rDir = 2
  If Not skipDir(4) Then: If AIPath(iCurNode) + 16 = EdgeCells(iEnd) Then rDir = 4
  If Not skipDir(3) Then: If AIPath(iCurNode) + 17 = EdgeCells(iEnd) Then rDir = 3
  
  If skipDir(1) And skipDir(2) And skipDir(3) And skipDir(4) Then
    badPath = True
    Exit Do
  End If
  If skipDir(rDir) Then GoTo retry
  Select Case rDir
    Case 1 ' go north
      iTestCell = AIPath(iCurNode) - 17
      If Not IsValidMove(AIPath(iCurNode), iTestCell, 1, EdgeCells(iEnd)) Then
        skipDir(1) = True
        GoTo retry
      End If
    Case 2 ' go east
      iTestCell = AIPath(iCurNode) - 16
      If Not IsValidMove(AIPath(iCurNode), iTestCell, 1, EdgeCells(iEnd)) Then
        skipDir(2) = True
        GoTo retry
      End If
    Case 3 ' go south
      iTestCell = AIPath(iCurNode) + 17
      If Not IsValidMove(AIPath(iCurNode), iTestCell, 1, EdgeCells(iEnd)) Then
        skipDir(3) = True
        GoTo retry
      End If
    Case 4 ' go west
      iTestCell = AIPath(iCurNode) + 16
      If Not IsValidMove(AIPath(iCurNode), iTestCell, 1, EdgeCells(iEnd)) Then
        skipDir(4) = True
        GoTo retry
      End If
  End Select

  iCurNode = iCurNode + 1
  AIPath(iCurNode) = iTestCell
  cellInfo(iTestCell).Free = False
  
  skipDir(1) = False
  skipDir(2) = False
  skipDir(3) = False
  skipDir(4) = False
  AINodes = iCurNode
  If bPaint Then TrackCell AIPath(iCurNode - 1), AIOrigo(iCurNode - 1), rDir
  If bPaint Then cellImg(iTestCell) = GRND_EARTH
  
  AIOrigo(iCurNode) = InvDir(rDir)
  
  If IsOnEdge(iTestCell) = hintDir Then
    If bPaint Then TrackCell iTestCell, AIOrigo(iCurNode), hintDir
    badPath = False
    Exit Do
  End If
Loop

If badPath Then
  badCount = badCount + 1
  If badCount < 100 Then
    If iCurNode > 5 Then
      cellImg(AIPath(iCurNode)) = GRND_GRASS
      cellImg(AIPath(iCurNode - 1)) = GRND_GRASS
      cellImg(AIPath(iCurNode - 2)) = GRND_GRASS
      cellImg(AIPath(iCurNode - 3)) = GRND_GRASS
      cellImg(AIPath(iCurNode - 4)) = GRND_GRASS
      cellImg(AIPath(iCurNode - 5)) = GRND_GRASS
      iCurNode = iCurNode - 5
      AINodes = iCurNode
      badPath = False
      skipDir(1) = False
      skipDir(2) = False
      skipDir(3) = False
      skipDir(4) = False
      GoTo retry
    End If
  End If
End If

If badPath Then
  ResetCells
  AI_MakePath
  Exit Function
End If

AI_MakePath = Not badPath
sMsg = "path find completed: " & IIf(badPath, "bad path", "good path") & " (" & badCount & ")"
End Function

Sub TrackCell(iCell As Integer, iDir1 As Integer, iDir2 As Integer)
  If (iDir1 = 1) And iDir2 = 2 Then cellImg(iCell) = GRND_TRACK_NE
  If (iDir1 = 1) And iDir2 = 3 Then cellImg(iCell) = GRND_TRACK_NS
  If (iDir1 = 1) And iDir2 = 4 Then cellImg(iCell) = GRND_TRACK_NW
  If (iDir1 = 2) And iDir2 = 1 Then cellImg(iCell) = GRND_TRACK_NE
  If (iDir1 = 2) And iDir2 = 3 Then cellImg(iCell) = GRND_TRACK_ES
  If (iDir1 = 2) And iDir2 = 4 Then cellImg(iCell) = GRND_TRACK_WE
  If (iDir1 = 3) And iDir2 = 1 Then cellImg(iCell) = GRND_TRACK_NS
  If (iDir1 = 3) And iDir2 = 2 Then cellImg(iCell) = GRND_TRACK_ES
  If (iDir1 = 3) And iDir2 = 4 Then cellImg(iCell) = GRND_TRACK_WS
  If (iDir1 = 4) And iDir2 = 1 Then cellImg(iCell) = GRND_TRACK_NW
  If (iDir1 = 4) And iDir2 = 2 Then cellImg(iCell) = GRND_TRACK_WE
  If (iDir1 = 4) And iDir2 = 3 Then cellImg(iCell) = GRND_TRACK_WS
End Sub

Function FindDir(iOrigo1 As Integer, iOrigo2 As Integer) As Integer
  If (iOrigo1 = 1) And iOrigo2 = 2 Then FindDir = 3
  If (iOrigo1 = 1) And iOrigo2 = 3 Then FindDir = 4
  If (iOrigo1 = 1) And iOrigo2 = 4 Then FindDir = 5
  If (iOrigo1 = 2) And iOrigo2 = 1 Then FindDir = 7
  If (iOrigo1 = 2) And iOrigo2 = 3 Then FindDir = 5
  If (iOrigo1 = 2) And iOrigo2 = 4 Then FindDir = 6
  If (iOrigo1 = 3) And iOrigo2 = 1 Then FindDir = 0
  If (iOrigo1 = 3) And iOrigo2 = 2 Then FindDir = 1
  If (iOrigo1 = 3) And iOrigo2 = 4 Then FindDir = 7
  If (iOrigo1 = 4) And iOrigo2 = 1 Then FindDir = 1
  If (iOrigo1 = 4) And iOrigo2 = 2 Then FindDir = 2
  If (iOrigo1 = 4) And iOrigo2 = 3 Then FindDir = 3
End Function

Function InvDir(iDir As Integer)
If iDir = 1 Then InvDir = 3
If iDir = 2 Then InvDir = 4
If iDir = 3 Then InvDir = 1
If iDir = 4 Then InvDir = 2
End Function

Function InvDir2(iDir As Integer)
'train directions
If iDir = 0 Then InvDir2 = 4
If iDir = 1 Then InvDir2 = 5
If iDir = 2 Then InvDir2 = 6
If iDir = 3 Then InvDir2 = 7
If iDir = 4 Then InvDir2 = 0
If iDir = 5 Then InvDir2 = 1
If iDir = 6 Then InvDir2 = 2
If iDir = 7 Then InvDir2 = 3
End Function
Function WhichEdge(iCell As Integer) As Integer 'send an index of EdgeCells array here, not a cellnumber!
If iCell < 17 Then WhichEdge = 1 ' north
If (iCell > 16) And (iCell < 33) Then WhichEdge = 3 ' south
If (iCell > 32) And (iCell < 59) Then WhichEdge = 4 ' west
If iCell > 58 Then WhichEdge = 2 ' east
End Function

Function IsOnEdge(iCell As Integer) As Integer
Dim iFind As Integer
iFind = -1

For i = 0 To 83
  If iCell = EdgeCells(i) Then iFind = i
Next

If iFind >= 0 Then IsOnEdge = WhichEdge(iFind)
End Function

Function IsValidMove(iCell1 As Integer, iCell2 As Integer, iDir As Integer, Optional iSkip As Integer = -1) As Boolean
Dim bValid As Boolean
bValid = True

'detects a collision with an edge cell/occupied cell...

If iCell2 = iSkip Then
  bValid = True
  GoTo nd
End If

If (iCell2 > 808) Or (iCell2 < 0) Then
  bValid = False
  Exit Function
End If

If cellImg(iCell2) <> GRND_GRASS Then
  bValid = False
  Exit Function
End If

For i = 0 To AINodes
  If AIPath(i) = iCell2 Then
    bValid = False
    Exit Function
  End If
Next

For i = 0 To 83
  If iCell2 = EdgeCells(i) Then bValid = False
Next


nd:
IsValidMove = bValid


End Function

Private Sub tmrDrawTrack_Timer()
'draws out the track
start:
If curDrawTrack > AINodes Then
  tmrDrawTrack.Enabled = False
  Exit Sub
End If

If curDrawTrack = AINodes Then
  TrackCell AIPath(curDrawTrack), AIOrigo(curDrawTrack), IsOnEdge(AIPath(curDrawTrack))
Else
  TrackCell AIPath(curDrawTrack), AIOrigo(curDrawTrack), InvDir(AIOrigo(curDrawTrack + 1))
End If

bDraw = True

curDrawTrack = curDrawTrack + 1
If Settings.iDirectTrackDraw Then GoTo start
End Sub
