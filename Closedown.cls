VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Closedown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=================================================================================
' Class Module for base module for Closedown API SDK. It include base functionality for
' RESTful web service request and parse json result. It uses Linkhub class module
' to accomplish authentication APIs.
'
' This module uses advapi32.dll and crypt32.dllm. most of windows system has it
' on system32 by default.
'
' http://www.linkhub.com
' Author : Jeong yohan (yhjeong@linkhub.co.kr)
' Written : 2015-06-21
'
' Thanks for your interest.
'=================================================================================
Private Const serviceID = "CLOSEDOWN"
Private Const ServiceURL = "https://closedown.linkhub.co.kr"
Private Const APIVersion = "1.0"

Private m_Linkhub As Linkhub
Private m_scope As New Collection
Private m_LastErrCode As Long
Private m_LastErrMessage As String
Private m_token As LinkhubToken

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'�ֱ� �����ڵ� Ȯ��
Public Property Get LastErrCode() As Long
    LastErrCode = m_LastErrCode
End Property
'�ֱٿ����޽��� Ȯ��
Public Property Get LastErrMessage() As String
    LastErrMessage = m_LastErrMessage
End Property
'�ʱ�ȭ
Public Sub Initialize(linkID As String, SecretKey As String)
    Set m_Linkhub = New Linkhub
    
    m_Linkhub.linkID = linkID
    m_Linkhub.SercetKey = SecretKey
    
    m_scope.Add "170"
End Sub
Private Function getSession_token() As String
    Dim refresh As Boolean
    
    refresh = False
       
    If m_token Is Nothing Then
        refresh = True
    Else
        Dim utcnow As String
        utcnow = Format(m_Linkhub.UTCTime, "yyyy-MM-ddTHH:mm:ss.sssZ")
        refresh = m_token.expiration < utcnow
    End If
    
    If refresh Then
    
        Set m_token = m_Linkhub.getToken(serviceID, "", m_scope)
        
        If m_token Is Nothing Then
            Err.Raise m_Linkhub.LastErrCode, "CLOSEDOWN", m_Linkhub.LastErrMessage
            Exit Function
        End If
        
    End If
    
    getSession_token = m_token.session_token

End Function

Private Function httpGET(url As String, BearerToken As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("GET", ServiceURL + url)
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-api-version", APIVersion)
        
    winhttp1.send
    winhttp1.WaitForResponse
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "CLOSEDOWN", parsedDic.Item("message")
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpGET = m_Linkhub.parse(Response)

End Function

Private Function httpPOST(url As String, BearerToken As String, postdata As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", ServiceURL + url)
    Call winhttp1.setRequestHeader("x-api-version", APIVersion)
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    
    winhttp1.send (postdata)
    winhttp1.WaitForResponse
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "CLOSEDOWN", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST = m_Linkhub.parse(Response)

End Function
'��Ʈ�� �ܾ���ȸ
Public Function GetBalance() As Double
    On Error GoTo ErrHandler
    
    GetBalance = m_Linkhub.GetPartnerBalance(getSession_token(), serviceID)
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetBalance = m_LastErrCode
End Function
'��ȸ�ܰ� Ȯ��
Public Function GetUnitCost() As Double
    On Error GoTo ErrHandler
    
    Dim result As Variant
    
    Set result = httpGET("/UnitCost", getSession_token())
    
    GetUnitCost = result.Item("unitCost")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetUnitCost = m_LastErrCode
End Function
'����ڹ�ȣ �ܰ���ȸ
Public Function CheckCorpNum(corpNum As String) As Corpstate
On Error GoTo ErrHandler
    If corpNum = "" Then
        Err.Raise -99999999, "CLOSEDOWN", "����ڹ�ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Check?CN=" + corpNum, getSession_token())
        
    Set CheckCorpNum = New Corpstate
        
    CheckCorpNum.fromDictionary result
            
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckCorpNum = Nothing
End Function
'����ڹ�ȣ �ٷ���ȸ(�ִ� 1000��)
Public Function CheckCorpNums(CorpNumList As Collection) As Collection
    On Error GoTo ErrHandler

    If CorpNumList Is Nothing Or CorpNumList.Count = 0 Then
        Err.Raise -99999999, "CLOSEDOWN", "����ڹ�ȣ �迭�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
        
    Dim postdata As String
    
    postdata = m_Linkhub.toString(CorpNumList)
    
    Set result = httpPOST("/Check", getSession_token(), postdata)
    
    Set CheckCorpNums = New Collection
    
    Dim T As Variant
    Dim CorpInfo As Corpstate
    
    For Each T In result
        Set CorpInfo = New Corpstate
        
        CorpInfo.fromDictionary T
                
        CheckCorpNums.Add CorpInfo
    Next
        
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckCorpNums = Nothing

End Function
