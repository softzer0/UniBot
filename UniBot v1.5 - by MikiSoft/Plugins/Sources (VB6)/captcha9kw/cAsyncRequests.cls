VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAsyncRequests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event ResponseStart(Req As cAsyncRequest, ByVal Status As Long, ByVal ContentType As String)
Event ResponseDataAvailable(Req As cAsyncRequest, Data() As Byte)
Event ResponseFinished(Req As cAsyncRequest)
Event Error(Req As cAsyncRequest, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)

Private mReqs As New Collection, mTimeout As Long

'centralized Event-Delegation
Friend Sub OnResponseStart(Req As cAsyncRequest, ByVal Status As Long, ByVal ContentType As String)
  RaiseEvent ResponseStart(Req, Status, ContentType)
End Sub
Friend Sub OnResponseDataAvailable(Req As cAsyncRequest, Data() As Byte)
  RaiseEvent ResponseDataAvailable(Req, Data)
End Sub
Friend Sub OnResponseFinished(Req As cAsyncRequest)
  RaiseEvent ResponseFinished(Req)
End Sub
Friend Sub OnError(Req As cAsyncRequest, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
  RaiseEvent Error(Req, ErrorNumber, ErrorDescription)
End Sub

Public Property Let Timeout(ByVal Milliseconds As Long)
  mTimeout = Milliseconds
End Property

Public Function AddRequest(Key As String, Optional Timeout As Long = -1, Optional Tag As String) As cAsyncRequest
  If Timeout > -1 Then mTimeout = Timeout
  mReqs.Add New cAsyncRequest, Key
  Set AddRequest = mReqs(Key)
      AddRequest.Init Me, Key, mTimeout, Tag
End Function

Public Function RemoveRequest(ReqObjOrKey) As cAsyncRequest
  If IsObject(ReqObjOrKey) Then mReqs.Remove ReqObjOrKey.Key Else mReqs.Remove ReqObjOrKey
End Function
 
Public Property Get RequestCount() As Long
  RequestCount = mReqs.Count
End Property

Public Property Get RequestItem(KeyOrOneBasedIndex) As cAsyncRequest
  Set RequestItem = mReqs(KeyOrOneBasedIndex)
End Property

Public Sub Cleanup()
Dim Req As cAsyncRequest
  For Each Req In mReqs
    Req.http.Abort
  Next
  Set mReqs = Nothing
End Sub

Private Sub Class_Terminate()
  Cleanup
End Sub
