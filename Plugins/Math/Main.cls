VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Math"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IPluginInterface

Dim Program As Form

Private Sub IPluginInterface_BuildSettings(strS As String)
End Sub

Private Function IPluginInterface_Startup(Parent As Object) As String
Set Program = Parent
IPluginInterface_Startup = "math"
End Function

Private Function IPluginInterface_Info(Inf() As String) As Object
Inf(0) = "1.0"
Inf(1) = "MikiSoft"
Inf(2) = "Parses strings as math expressions." & vbNewLine & "Example: math('3*2+sqr(4)-2^2')" & vbNewLine & "Special thanks to: Leonardo Volpi, Michael Ruder, Thomas Zeutschler, Lieven Dossche and Arnaud d.Grammont."
End Function

Private Function IPluginInterface_Execute(Params() As String) As String
Dim Funct As New clsMathParser
Funct.StoreExpression Params(1)
IPluginInterface_Execute = Funct.Eval
Set Funct = Nothing
End Function
