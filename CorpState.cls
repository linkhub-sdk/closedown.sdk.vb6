VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorpSTate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public corpNum      As String
Public ctype        As String
Public state        As String
Public stateDate    As String
Public checkDate    As String


Public Sub fromDictionary(dic As Variant)
    If IsEmpty(dic.Item("corpNum")) = False Then
        corpNum = dic.Item("corpNum")
    End If
    
    If IsEmpty(dic.Item("type")) = False Then
        ctype = dic.Item("type")
    End If
    
    If IsEmpty(dic.Item("state")) = False Then
        state = dic.Item("state")
    End If
    
    If IsEmpty(dic.Item("stateDate")) = False Then
        stateDate = dic.Item("stateDate")
    End If
    
    If IsEmpty(dic.Item("checkDate")) = False Then
        checkDate = dic.Item("checkDate")
    End If
End Sub
