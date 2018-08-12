VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegexMatchCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private matchCollection_ As Collection

Private Sub Class_Initialize()
    Set matchCollection_ = New Collection
End Sub

Public Property Get Item(ByVal index As Long) As RegexMatch
    Set Item = matchCollection_.Item(index)
End Property

Friend Property Get MatchCollection() As Collection
    Set MatchCollection = matchCollection_
End Property

Private Sub Class_Terminate()
    Set matchCollection_ = Nothing
End Sub

Public Function NewEnum() As IUnknown
    Set NewEnum = matchCollection_.[_NewEnum]
End Function
