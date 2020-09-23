VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ColumnCol As Collection

Public SQLTableName As String
Public SQLTableFieldsCount As Integer
Public SQLTableProcessed As Boolean
Public SQLTableID As String
Public SQLTableOwner As String

Private Sub Class_Initialize()
    Set ColumnCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set ColumnCol = Nothing
End Sub
