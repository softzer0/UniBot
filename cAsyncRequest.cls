VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAsyncRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const HTTPREQUEST_PROXYSETTING_PROXY& = 2

Public WithEvents http As WinHttp.WinHttpRequest
Attribute http.VB_VarHelpID = -1
Private mParent As cAsyncRequests, mUrl As String, mKey As String, mTimeout As Long ', mTag As String
 
Friend Sub Init(ParentHandler As cAsyncRequests, Key As String, Timeout As Long) ', Tag As String)
  Set mParent = ParentHandler
  mKey = Key
  mTimeout = Timeout
  'mTag = Tag
  Set http = New WinHttp.WinHttpRequest
End Sub

'Public Property Let Tag(Tag As String)
'mTag = Tag
'End Property

Public Property Get Key() As String
  Key = mKey
End Property

Public Property Get URL() As String
  URL = mUrl
End Property

Public Sub SendRequest(URL As String, Optional PostBody As Variant, Optional AdditionalHeaders As Collection, Optional Proxy As String)
  mUrl = URL
  http.Option(4) = 13056
  http.Option(WinHttpRequestOption_EnableRedirects) = False
  http.Option(12) = True
  If Proxy <> vbNullString Then http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, Proxy
  If mTimeout > 0 Then http.SetTimeouts mTimeout, mTimeout, mTimeout, mTimeout
  If VarType(PostBody) <> vbArray + vbByte Then
    If PostBody <> Empty Then
      If PostBody <> "-" And PostBody <> "@" Then
        If PostBody = vbNullChar Then PostBody = vbNullString
        http.Open "POST", URL, True
        'If VarType(PostBody) = vbString Then If InStr(PostBody, vbNewLine) = 0 Then http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
      Else
        If PostBody = "-" Then http.Open "DELETE", URL, True Else: http.Open "HEAD", URL, True
        PostBody = vbNullString
      End If
    Else: http.Open "GET", URL, True
    End If
  Else: If Left$(URL, 1) = "*" Then http.Open "PUT", Mid$(URL, 2), True Else: http.Open "POST", URL, True
  End If
  If AdditionalHeaders.count > 0 Then
    Dim i As Integer
    For i = 1 To AdditionalHeaders.count
      http.SetRequestHeader Split(AdditionalHeaders.Item(i), vbLf)(0), Split(AdditionalHeaders.Item(i), vbLf)(1)
    Next
  End If
  http.Send PostBody 'send the http-request
End Sub

Public Sub AbortRequest()
  If Not http Is Nothing Then http.Abort
End Sub

'Event-Delegation by direct Calls into the Parent (the aggregating cAsyncRequests-Instance)
Private Sub http_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
  mParent.OnResponseStart Me, Status, ContentType
End Sub
Private Sub http_OnResponseDataAvailable(Data() As Byte)
  mParent.OnResponseDataAvailable Me, Data
End Sub
Private Sub http_OnResponseFinished()
  mParent.OnResponseFinished Me
End Sub
Private Sub http_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
  mParent.OnError Me, ErrorNumber, ErrorDescription
End Sub

Private Sub Class_Terminate()
  If Not http Is Nothing Then
    http.Abort
    Set http = Nothing
  End If
End Sub
