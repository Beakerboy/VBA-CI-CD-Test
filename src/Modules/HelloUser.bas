Attribute VB_Name = "HelloUser"
Public Function TestClass()
    Dim emp As New clsName
    'emp = Create_Name()
    emp.Last = "Animal"
    emp.Middle = "Filthy"
    emp.First = "U."
    emp.Change
End Function

Public Function Create_Name() As clsName
    Set Create_Name = New clsName
End Function
