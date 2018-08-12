VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private index_ As Long
Private value_ As String
Private length_ As Long
Private submatchCollection_ As Collection

Private Sub Class_Initialize()
    Set submatchCollection_ = New Collection
End Sub

Public Property Get SubMatches(ByVal Index As Long) As String
    SubMatches = submatchCollection_.Item(Index)
End Property

Friend Property Get SubMatchCollection() As Collection
    Set SubMatchCollection = submatchCollection_
End Property

Public Property Get Index() As Long
    Index = index_
End Property

Friend Property Let Index(ByVal rhs As Long)
    index_ = rhs
End Property

Public Property Get Value() As String
    Value = value_
End Property

Friend Property Let Value(ByVal rhs As String)
    value_ = rhs
End Property

Public Property Get Length() As Long
    Length = length_
End Property

Friend Property Let Length(ByVal rhs As Long)
    length_ = rhs
End Property

Private Sub Class_Terminate()
    Set submatchCollection_ = Nothing
End Sub

Public Function NewEnum() As IUnknown
    Set NewEnum = submatchCollection_.[_NewEnum]
End Function

