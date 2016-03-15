VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "adeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Last change 2004-01-08
'------------------------------------------------------------------------------------
' Welcome to the source code of...
' aDeSettings Class version 1.0
' © Adrian Nilsson (aDe) 2004
' www.ade.se
' Latest version of this class can be obtained at the above web adress.
'------------------------------------------------------------------------------------
' Thanks for voting: http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=49556&lngWId=1
'------------------------------------------------------------------------------------
' You may modify and/or use this class in your own programs.
' You may redistribute this unmodified sourcecode freely.
' You may not redistribute this sourcecode modified in any way - unless as a part of
' another open-source program, without my personal admittance.
' You may not remove the information above.
'------------------------------------------------------------------------------------
' Usage information about the class:
' This class will manage your program user's settings, or other data, and store it
' in a textfile in your application's directory.
' You can use sections, each section hold settings.
' File format:
'   [section1]
'   <setting1></setting1>
'   <setting2></setting2>
'   [/section1]
'   [section2]
'   <setting1></setting1>
'   <setting2></setting2>
'   [/section2]
'
' If you set the Encrypted property to True, a very simple encryption will be added
' to mask the settings for user manipulation.
' It is a very unsafe encryption though, and any skilled person could decrypt it
' with ease.
'
' To use the class, simply add it to a form or other object's declarations:
' Public WithEvents Settings As adeSettings
'
' Then, create the class when the object is started:
' Set Settings = New adeSettings
'
' After that, set the properties and use the functions.
' If the class detects errors, the Error event will be raised.
' These are usually only for debugging, you can use the error-numbers for detecting
' different things though.
' Good luck
' ade

'----- Property value holders ---------
Private mvarSettingsFile As String
Private mvarAutoReadWrite As Boolean
Private mvarEncryptionKey As Integer
Private mvarDefaultSection As String
Private mvarEncrypt As Boolean
'--------------------------------------

Private bLoadedOnce As Boolean ' To remember if the settings-file has been auto-loaded once or not

Public Event Error(strDescription As String, iNumber As Integer) ' Define the Error event.

Public fileData As String  ' This variable always contains the current settings-file data

'------------------ Error messages / codes ------------------------------------
Private Const errNoFile = "Couldn't read file, file not found"
Private Const numNoFile = 0

Private Const errNoSettingFile = "No SettingsFile has been specified"
Private Const numNoSettingFile = 1

Private Const errWriteError = "Couldn't save - write error. Check your permissions and settingfile's attributes (read-only?)" & vbCrLf & "Error: "
Private Const numWriteError = 2

Private Const errRemoveParameters = "Couldn't remove - no parameters!"
Private Const numRemoveParameters = 3

Private Const errNotFound = "Setting/section not found"
Private Const numNotFound = 4

Private Const errReadError = "Unknown read error: "
Private Const numReadError = 5
'------------------------------------------------------------------------------

Private Sub Class_Initialize()
  ' Default values that will be used if these settings aren't made.
  Me.DefaultSection = "general"
  Me.EncryptionKey = 48
End Sub

Function Remove(Optional Section As String, Optional Setting As String) As Boolean
Dim sectData As String, sVal As String, oldSectData As String

' Check if both Section and Setting was omitted
If Len(Section) + Len(Setting) = 0 Then
  RaiseEvent Error(errRemoveParameters, numRemoveParameters)
  Exit Function
End If

' Check wheter to load the settingsfile if it hasn't been done yet
If Not bLoadedOnce Then
  If Len(fileData) = 0 Then
    If Len(Me.SettingsFile) Then
      Me.ReadFile
    Else
      RaiseEvent Error(errNoSettingFile, numNoSettingFile)
    End If
    bLoadedOnce = True
  End If
End If

' Use the default section if none specified
If Len(Section) = 0 Then Section = Me.DefaultSection

' Find and grab the whole section to sectData
sectData = BeforeFirst(AfterFirst(fileData, "[" & Section & "]"), "[/" & Section & "]")
If sectData = "" Then
  ' The section was not found
  RaiseEvent Error(errNotFound, numNotFound)
  Exit Function
End If
sectData = "[" & Section & "]" & sectData & "[/" & Section & "]" & vbCrLf
oldSectData = sectData

If Len(Setting) Then
  ' Setting was specified, find and grab it to sVal
  sVal = BeforeFirst(AfterFirst(sectData, "<" & Setting & ">"), "</" & Setting & ">")
  If sVal = "" Then
    ' The setting was not found
    RaiseEvent Error(errNotFound, numNotFound)
    Exit Function
  End If
  sVal = "<" & Setting & ">" & sVal & "</" & Setting & ">" & vbCrLf
  ' Remove the setting
  sectData = Replace(sectData, sVal, "")
  fileData = Replace(fileData, oldSectData, sectData)
Else
  ' Setting not specified. Remove the whole section
  fileData = Replace(fileData, sectData, "")
End If

Remove = True
End Function

Function Read(Setting, Optional Default As String, Optional Section As String) As String
Dim sVal As String, sectData As String

' Use default section if none specified
If Section = "" Then Section = Me.DefaultSection

' Check wheter to load the settingsfile if it hasn't been done yet
If Not bLoadedOnce Then
  If Len(fileData) = 0 Then
    If Len(Me.SettingsFile) Then
      Me.ReadFile
    Else
      RaiseEvent Error(errNoSettingFile, numNoSettingFile)
    End If
    bLoadedOnce = True
  End If
End If

' Find the setting...
If InStr(1, fileData, "[" & Section & "]") Then
  ' Find and grab the section to sectData
  sectData = BeforeFirst(AfterFirst(fileData, "[" & Section & "]"), "[/" & Section & "]")
  If InStr(1, sectData, "<" & Setting & ">") Then
    sVal = BeforeFirst(AfterFirst(sectData, "<" & Setting & ">"), "</" & Setting & ">")
  Else
    ' If setting not found - set the specfied Default
    sVal = Default
  End If
Else
  ' Section not found - set the specified Default
  sVal = Default
End If
Read = UnEscape(sVal) ' Remove character escaping and set the function return value.
End Function

Sub Save(Setting As String, Value As String, Optional Section As String)
Dim sOldSect As String, sectData As String, sOldVal As String

' Use default section if none specified
If Section = "" Then Section = Me.DefaultSection

' Check wheter to load the settingsfile if it hasn't been done yet
If Len(fileData) = 0 Then
  If Not bLoadedOnce Then
    If Len(Me.SettingsFile) Then
      Me.ReadFile , True ' True = Tell ReadFile function to not send errors.
    End If
    bLoadedOnce = True
  End If
End If

If InStr(1, fileData, "[" & Section & "]") Then
  ' Find and grab the section to sectData
  sectData = BeforeFirst(AfterFirst(fileData, "[" & Section & "]"), "[/" & Section & "]")
  sOldSect = sectData
  If InStr(1, sectData, "<" & Setting & ">") Then
    ' Setting exists in section - replace the current
    sOldVal = BeforeFirst(AfterFirst(sectData, "<" & Setting & ">"), "</" & Setting & ">")
    sectData = Replace(sectData, "<" & Setting & ">" & sOldVal, "<" & Setting & ">" & Escape(Value))
  Else
    ' Setting doesnt exist in section - add new
    sectData = sectData & "<" & Setting & ">" & Value & "</" & Setting & ">" & vbCrLf
  End If
  ' Replace the old section with the updated one.
  fileData = Replace(fileData, "[" & Section & "]" & sOldSect, "[" & Section & "]" & sectData)
Else
  ' Section wasn't found - add the section and value
  fileData = fileData & "[" & Section & "]" & vbCrLf & "<" & Setting & ">" & Escape(Value) & "</" & Setting & ">" & vbCrLf & "[/" & Section & "]" & vbCrLf
End If

If Me.AutoSave Then Me.SaveFile
End Sub

Sub ClearAll()
  ' Clear all setting data
  fileData = ""
End Sub

Function Escape(strEsc As String) As String
Dim sOut As String
' Add character escaping in order to avoid confusing the file parser
sOut = Replace(strEsc, "%", "%25")
sOut = Replace(sOut, "<", "%3C")
sOut = Replace(sOut, ">", "%3E")
sOut = Replace(sOut, "[", "%5B")
sOut = Replace(sOut, "]", "%5D")
Escape = sOut
End Function

Function UnEscape(strEsc As String) As String
Dim sOut As String
' Remove escaping
sOut = Replace(strEsc, "%5D", "]")
sOut = Replace(sOut, "%5B", "[")
sOut = Replace(sOut, "%3C", "<")
sOut = Replace(sOut, "%3E", ">")
sOut = Replace(sOut, "%25", "%")
UnEscape = sOut
End Function

Function ReadFile(Optional File As String, Optional IgnoreErrors As Boolean) As Boolean
On Error GoTo readErr

If Len(File) = 0 Then File = Me.SettingsFile

' Check if for filename
If Len(File) = 0 Then
  If Not IgnoreErrors Then RaiseEvent Error(errNoSettingFile, numNoSettingFile)
  Exit Function
End If

' Check if file exists
If Dir(File) = "" Then
  If Not IgnoreErrors Then RaiseEvent Error(errNoFile, numNoFile)
  Exit Function
End If

' Read the file.
Open File For Input As 1
Do While Not EOF(1)
   fileData = Input(LOF(1), 1)
Loop
Close 1

' Check if file is encrypted, and decrypt if it is
If Left(fileData, 1) = "E" Then
  fileData = DecryptStr(Right(fileData, Len(fileData) - 1))
End If
ReadFile = True
Exit Function

readErr:
RaiseEvent Error(errReadError & Error$, numReadError) 'Send VB's error description
End Function

Function SaveFile(Optional ByVal File As String) As Boolean
Dim saveData As String
On Error GoTo errRaise
If Len(File) = 0 Then File = Me.SettingsFile
' Check if the settingsfile property has been set
If Len(File) = 0 Then
  RaiseEvent Error(errNoSettingFile, numNoSettingFile)
  Exit Function
End If

' Make a copy of fileData to encrypt if Encrypted property is True
saveData = fileData
If Me.Encrypted Then
  saveData = "E" & EncryptStr(saveData)
End If

' Save file
Open File For Binary As 1
Put #1, 1, saveData
Close 1

Exit Function
errRaise:
RaiseEvent Error(errWriteError & Error$, numWriteError) ' Send VB's error description
End Function

Private Function EncryptStr(strEnc As String) As String
Dim i As Integer, sOut As String, char As String

' Make a simple character shifting encryption (not very safe)
For i = 1 To Len(strEnc)
  char = Left(Right(strEnc, i), 1)
  If Asc(char) - Me.EncryptionKey < 1 Then
    sOut = sOut & Chr(255 + (Asc(char) - Me.EncryptionKey - 1))
  Else
    sOut = sOut & Chr(Asc(char) - Me.EncryptionKey)
  End If
Next
EncryptStr = sOut
End Function

Private Function DecryptStr(strDec As String) As String
Dim i As Integer, sOut As String, char As String

' Remove simple encryption
For i = 1 To Len(strDec)
  char = Left(Right(strDec, i), 1)
  If (Asc(char) + Me.EncryptionKey) > 255 Then
    sOut = sOut & Chr((Asc(char) + Me.EncryptionKey + 1) - 255)
  Else
    sOut = sOut & Chr(Asc(char) + Me.EncryptionKey)
  End If
Next
DecryptStr = sOut
End Function


' The rest is some string-handling functions and property functions...

Public Property Let DefaultSection(ByVal vData As String)
    mvarDefaultSection = vData
End Property

Public Property Get DefaultSection() As String
    DefaultSection = mvarDefaultSection
End Property

Public Property Let EncryptionKey(ByVal vData As Integer)
    mvarEncryptionKey = vData
End Property

Public Property Get EncryptionKey() As Integer
    EncryptionKey = mvarEncryptionKey
End Property

Public Property Let Encrypted(ByVal vData As Boolean)
    mvarEncrypt = vData
End Property

Public Property Get Encrypted() As Boolean
    Encrypted = mvarEncrypt
End Property

Public Property Let AutoSave(ByVal vData As Boolean)
    mvarAutoReadWrite = vData
End Property

Public Property Get AutoSave() As Boolean
    AutoSave = mvarAutoReadWrite
End Property

Public Property Let SettingsFile(ByVal vData As String)
    mvarSettingsFile = vData
End Property

Public Property Get SettingsFile() As String
    SettingsFile = mvarSettingsFile
End Property

