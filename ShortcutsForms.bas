Attribute VB_Name = "ShortcutsForms"
''Not Used
Public Function AdjustWidth(thisForm As UserForm, length As Integer) As UserForm
    UserForm.Width = length
    Set Extend = UserForm
End Function
