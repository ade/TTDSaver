VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCellPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarpCol As Collection 'local copy
Public Property Set pCol(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.pCol = Form1
    Set mvarpCol = vData
End Property


Public Property Get pCol() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.pCol
    Set pCol = mvarpCol
End Property


Function Item(sIndex As String) As clsCell
On Error Resume Next
Set Item = pCol.Item("$" & sIndex & "$")
End Function

Function Add(Image As Integer) As clsCell
Dim sNew As clsCell
Set sNew = New clsCell

sNew.Image = Image
sNew.Index = pCol.Count + 1
pCol.Add sNew

Set Add = sNew
End Function

Private Sub Class_Initialize()
Set pCol = New Collection

End Sub

Private Sub Class_Terminate()
Set pCol = Nothing
End Sub
