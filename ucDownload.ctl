VERSION 5.00
Begin VB.UserControl ucDownload 
   BackStyle       =   0  'Í¸Ã÷
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   825
   ScaleHeight     =   540
   ScaleWidth      =   825
End
Attribute VB_Name = "ucDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event DownloadProgress(ByVal BytesRead As Long, ByVal BytesTotal As Long)
Event DownloadFailed(ByVal Status As String, ByVal StatusCode As AsyncStatusCodeConstants)
Event DownloadComplete()

Public Url As String, LocalFileName As String
Attribute LocalFileName.VB_VarUserMemId = 1073938432

Public Sub DownloadFile(Url As String, LocalFileName As String, Optional ByVal Mode As AsyncReadConstants = vbAsyncReadForceUpdate)
    CancelDownload
    Me.LocalFileName = LocalFileName
    Me.Url = Url
    On Error Resume Next
    AsyncRead Url, vbAsyncTypeFile, LocalFileName, Mode
    If Err Then
        RaiseEvent DownloadFailed(Err.Description, Err.Number)
    End If
End Sub

Public Sub CancelDownload()
    On Error Resume Next
    CancelAsyncRead LocalFileName    'cancel a possibly still running Download with the same Destination-Filename
    On Error GoTo 0
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    With AsyncProp
        If .BytesRead Then RaiseEvent DownloadProgress(.BytesRead, IIf(.BytesMax <= .BytesRead, .BytesRead, .BytesMax))
    End With
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    If AsyncProp.StatusCode <> vbAsyncStatusCodeEndDownloadData Or AsyncProp.BytesRead = 0 Then
        RaiseEvent DownloadFailed(AsyncProp.Status, AsyncProp.StatusCode)
        CancelDownload
    Else
        Name AsyncProp.Value As LocalFileName
        RaiseEvent DownloadComplete
    End If
End Sub
