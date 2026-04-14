VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public First As String
Public Middle As String
Public Last As String

Public Sub Change()
    MsgBox "Keep the change " & Join(Array(Me.First, Me.Middle, Me.Last))
End Sub
