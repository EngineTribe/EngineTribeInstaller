Attribute VB_Name = "Module1"
Public Locale As String
Public Stage As Integer
Public FreshInstall As Boolean

Public Const LatestVersionLink As String = "https://storage.enginetribe.gq/enginetribe-releases/latest.txt?raw=true&proxied=true"
Public Const VanillaLink As String = "https://storage.enginetribe.gq/enginetribe-releases/SMM_WE_3.2.3.exe?raw=true"
Public Const PatchRoot  As String = "https://storage.enginetribe.gq/enginetribe-releases/"


Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Function ChooseDir(ByVal frmTitle As String, onForm As Object) As String
'oleexp 选择目录
    On Error Resume Next
    Dim pChooseDir As New FileOpenDialog
    Dim psiResult As IShellItem
    Dim lpPath As Long, sPath As String
    With pChooseDir
        .SetOptions FOS_PICKFOLDERS
        .SetTitle frmTitle
        .Show onForm.hWnd
        .GetResult psiResult
        If (psiResult Is Nothing) = False Then
            psiResult.GetDisplayName SIGDN_FILESYSPATH, lpPath
            If lpPath Then
                SysReAllocString VarPtr(sPath), lpPath
                CoTaskMemFree lpPath
            End If
        End If
    End With
    If BStrFromLPWStr(lpPath) <> "" Then ChooseDir = BStrFromLPWStr(lpPath)
End Function

Public Function ChooseFile(ByVal frmTitle As String, ByVal fileDescription As String, ByVal fileFilter As String, ByVal onForm As Object) As String
'oleexp 选择文件
    On Error Resume Next
    Dim pChoose As New FileOpenDialog
    Dim psiResult As IShellItem
    Dim lpPath As Long, sPath As String
    Dim tFilt() As COMDLG_FILTERSPEC
    ReDim tFilt(0)
    tFilt(0).pszName = fileDescription
    tFilt(0).pszSpec = fileFilter
    With pChoose
        .SetFileTypes UBound(tFilt) + 1, VarPtr(tFilt(0))
        .SetTitle frmTitle
        .SetOptions FOS_FILEMUSTEXIST + FOS_DONTADDTORECENT
        .Show onForm.hWnd
        .GetResult psiResult
        If (psiResult Is Nothing) = False Then
            psiResult.GetDisplayName SIGDN_FILESYSPATH, lpPath
            If lpPath Then
                SysReAllocString VarPtr(sPath), lpPath
                CoTaskMemFree lpPath
            End If
        End If
    End With
    If BStrFromLPWStr(lpPath) <> "" Then ChooseFile = BStrFromLPWStr(lpPath)
End Function

Public Function BStrFromLPWStr(lpWStr As Long) As String
    SysReAllocString VarPtr(BStrFromLPWStr), lpWStr
End Function

Public Function GETString(Url As String) As String
    With New MSXML2.ServerXMLHTTP30
        .Open "GET", Url, False
        .setRequestHeader "User-Agent", "EngineTribeInstaller/1"
        .send
        GETString = .responseText
    End With
End Function

Public Sub ShellAndWait(pathFile As String)
    With CreateObject("WScript.Shell")
        .Run pathFile, 0, True
    End With
End Sub

Public Sub Unzip(ZipPath As String, UnzipTo As String)
    ShellAndWait Chr(34) & App.Path & "\7z-cmd\7z.exe" & Chr(34) & " x " & Chr(34) & ZipPath & Chr(34) & " -o" & Chr(34) & UnzipTo & Chr(34) & " -aoa"
End Sub
