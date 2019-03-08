VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ARDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''Not Used
Private m_ARDocuments As Collection
Public Property Get Count() As Integer
    Count = m_ARDocuments.Count
End Property
Public Sub Add(ARDoc As ARDocument)
    Stop
    m_ARDocuments.Add Item:=ARDoc
End Sub
Public Sub Remove(ARDoc As ARDocument)
    m_ARDocuments.Remove ARDoc
End Sub
Public Function Item(index As Integer) As ARDocument
    Item = m_ARDocuments.Item(index)
End Function
