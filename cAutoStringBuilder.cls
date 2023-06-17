VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAutoStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum StringBuilderTypes
    DotNET
    PureVB
End Enum

Private StringBuilder As Object
Public StringBuilderType As StringBuilderTypes

Public Sub InitStringBuilder(Optional PreferredType As StringBuilderTypes = DotNET)
    On Error GoTo UsePureVBImplementation
    If PreferredType = DotNET Then
        Set StringBuilder = CreateObject("System.Text.StringBuilder")
        StringBuilderType = DotNET
        Exit Sub
    Else
        GoTo UsePureVBImplementation
    End If
UsePureVBImplementation:
    Set StringBuilder = New cStringBuilder
    StringBuilderType = PureVB
End Sub

Private Sub Class_Initialize()
    InitStringBuilder
End Sub

Public Sub Append(ByRef sThis As String)
    If StringBuilderType = DotNET Then
        StringBuilder.Append_3 sThis
    Else
        StringBuilder.Append sThis
    End If
End Sub

Public Sub AppendByVal(ByVal sThis As String)
    Append sThis
End Sub

Public Sub AppendNL(ByRef sThis As String)
    Append sThis
    Append vbCrLf
End Sub

Public Property Get ToString() As String
    ToString = StringBuilder.ToString
End Property
