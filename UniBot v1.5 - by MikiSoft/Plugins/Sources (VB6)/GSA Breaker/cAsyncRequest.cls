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
Private mParent As cAsyncRequests, mUrl As String, mKey As String, mTimeout As Long, mTag As String
 
Friend Sub Init(ParentHandler As cAsyncRequests, Key As String, Timeout As Long, Optional Tag As String)
  Set mParent = ParentHandler
  mKey = Key
  mTimeout = Timeout
  mTag = Tag
  Set http = New WinHttp.WinHttpRequest
End Sub

Public Property Let Tag(Tag As String)
  mTag = Tag
End Property

Public Property Get Tag() As String
  Tag = mTag
End Property

Public Property Get Key() As String
  Key = mKey
End Property

Public Property Get URL() As String
  URL = mUrl
End Property

Public Sub SendRequest(URL As String, Optional PostBody As Variant, Optional AdditionalHeaders As String, Optional Proxy As String)
  mUrl = URL
  http.Option(4) = 13056
  'http.Option(WinHttpRequestOption_EnableRedirects) = False
  http.Option(12) = True
  If Proxy <> vbNullString Then http.SetProxy HTTPREQUEST_PROXYSETTING_PROXY, Proxy
  If mTimeout > 0 Then http.SetTimeouts mTimeout, mTimeout, mTimeout, mTimeout
  Dim t As String: t = "GET"
  If Not IsMissing(PostBody) Then
    If VarType(PostBody) = vbString Then
      If StrPtr(PostBody) <> 0 Then t = "POST"
    Else: If UBound(PostBody) >= LBound(PostBody) Then t = "POST"
    End If
  End If
  http.Open t, URL, True
  If t = "POST" And VarType(PostBody) <> vbByte + vbArray Then
    If InStr(PostBody, vbNewLine) = 0 Then
      Const PostH As String = "Content-Type" & vbLf & "application/x-www-form-urlencoded"
      If InStr(AdditionalHeaders, PostH) = 0 Then If AdditionalHeaders = vbNullString Then AdditionalHeaders = PostH Else: AdditionalHeaders = AdditionalHeaders & vbLf & PostH
    End If
  End If
  Dim i As Integer, s() As String
  s() = Split(AdditionalHeaders, vbLf)
  For i = 0 To UBound(s())
    If i Mod 2 <> 0 Then http.SetRequestHeader t, s(i) Else: t = s(i)
  Next
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
  If Not http Is Nothing Then http.Abort
End Sub
