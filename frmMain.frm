VERSION 5.00
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "vbccr17.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Engine Tribe Installer"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9615
   StartUpPosition =   3  '窗口缺省
   Begin VBCCR17.ProgressBar DownloadProgress 
      Height          =   375
      Left            =   240
      Top             =   2040
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      Step            =   10
      Scrolling       =   1
   End
   Begin EngineTribeInstaller.ucDownload Downloader 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   4200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
   End
   Begin VB.TextBox FileSelect 
      Height          =   405
      Left            =   240
      TabIndex        =   8
      Text            =   "Path to SMM_WE.exe (Double click to select)"
      Top             =   3000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.OptionButton OptionInstallMethod 
      BackColor       =   &H80000005&
      Caption         =   "Patch existing game"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.OptionButton OptionInstallMethod 
      BackColor       =   &H80000005&
      Caption         =   "Fresh install"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "Next"
      Height          =   615
      Left            =   7800
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
   Begin VB.OptionButton OptionLocales 
      BackColor       =   &H80000005&
      Caption         =   "Espanol"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.OptionButton OptionLocales 
      BackColor       =   &H80000005&
      Caption         =   "English"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OptionButton OptionLocales 
      BackColor       =   &H80000005&
      Caption         =   "中文"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label L13 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Label L12 
      BackStyle       =   0  'Transparent
      Caption         =   "First, select your language."
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   8280
      Picture         =   "frmMain.frx":54AA
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label L11 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Engine Tribe, an open source unofficial online service for SMM:WE!"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DownloadCompleted As Boolean

Private Sub btnNext_Click()
    Select Case Stage
    Case 1
        LoadStage2Text
        L12.Visible = False
        For i = 0 To 2
            OptionLocales(i).Visible = False
        Next
        For i = 0 To 1
            OptionInstallMethod(i).Visible = True
        Next
        FileSelect.Visible = True
        OptionInstallMethod(0).Value = True
        Stage = 2
    Case 2
        If InStr(FileSelect.Text, "\") = 0 Then Exit Sub
        For i = 0 To 1
            OptionInstallMethod(i).Visible = False
        Next
        FileSelect.Visible = False
        Stage = 3
        LoadStage3Text
        InstallTheGame
    End Select
End Sub

Private Sub FileSelect_DblClick()
    Dim FileName As String
    If FreshInstall Then
        FileName = ChooseDir("Installation folder", frmMain)
    Else
        FileName = ChooseFile("SMM_WE.exe", "SMM_WE", "SMM_WE.exe", frmMain)
    End If
    If FileName = "" Then
        Exit Sub
    Else
        FileSelect.Text = FileName
    End If
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub Form_Load()
    OptionLocales(0).Value = True
    OptionLocales(1).Value = False
    OptionLocales(2).Value = False
    Locale = "CN"
    Stage = 1
End Sub

Private Sub LoadStage1Text()
    Select Case Locale
    Case "CN"
        frmMain.Caption = "引擎部落安装器"
        L11.Caption = "欢迎来到引擎部落，SMM:WE 的开源在线服务！"
        L12.Caption = "首先，选择你的语言。"
        btnNext.Caption = "下一步"
    Case "EN"
        frmMain.Caption = "Engine Tribe Installer"
        L11.Caption = "Welcome to Engine Tribe, an open source unofficial online service for SMM:WE!"
        L12.Caption = "First, select your language."
        btnNext.Caption = "Next"
    Case "ES"
        frmMain.Caption = "Instalador de Engine Tribe"
        L11.Caption = "Bienvenido a Engine Tribe, un servicio en linea no oficial de codigo abierto para SMM:WE!"
        L12.Caption = "Primero, seleccione su idioma."
        btnNext.Caption = "Proximo"
    End Select
End Sub

Private Sub LoadStage2Text()
    Select Case Locale
    Case "CN"
        L11.Caption = "你想要安装新游戏，还是在已经安装好的游戏上打补丁？"
        OptionInstallMethod(0).Caption = "安装新游戏"
        OptionInstallMethod(1).Caption = "在已安装好的游戏上打补丁"
    Case "EN"
        L11.Caption = "Do you want to do a fresh install, or patch an already installed game?"
        OptionInstallMethod(0).Caption = "Fresh install"
        OptionInstallMethod(1).Caption = "Patch existing game"
    Case "ES"
        L11.Caption = "Quieres hacer una instalacion nueva o parchear un juego ya instalado?"
        OptionInstallMethod(0).Caption = "Instalacion nueva"
        OptionInstallMethod(1).Caption = "Parchear el juego existente"
    End Select
End Sub


Private Sub LoadStage3Text()
    Select Case Locale
    Case "CN"
        L11.Caption = "正在安装 ..."
    Case "EN"
        L11.Caption = "Installing ..."
    Case "ES"
        L11.Caption = "Instalando ..."
    End Select
End Sub
Private Sub OptionInstallMethod_Click(Index As Integer)
    Select Case Index
    Case 0
        FreshInstall = True
        Select Case Locale
        Case "CN"
            FileSelect.Text = "安装文件夹 (双击选择)"
        Case "EN"
            FileSelect.Text = "Install folder (Double click to select)"
        Case "ES"
            FileSelect.Text = "Carpeta de instalacion (Doble clic para seleccionar)"
        End Select
    Case 1
        FreshInstall = False
        Select Case Locale
        Case "CN"
            FileSelect.Text = "SMM_WE.exe 路径 (双击选择)"
        Case "EN"
            FileSelect.Text = "Path to SMM_WE.exe (Double click to select)"
        Case "ES"
            FileSelect.Text = "Ruta a SMM_WE.exe (Doble clic para seleccionar)"
        End Select
    End Select
End Sub

Private Sub OptionLocales_Click(Index As Integer)
    Select Case Index
    Case 0: Locale = "CN"
    Case 1: Locale = "EN"
    Case 2: Locale = "ES"
    End Select
    LoadStage1Text
End Sub


Private Sub InstallTheGame()
    Dim InstallPath As String, LatestVersion As String, FSO As New FileSystemObject, PatchURL As String
    If FreshInstall Then
        InstallPath = FileSelect.Text
    Else
        InstallPath = Left(FileSelect.Text, Len(FileSelect.Text) - 10)
    End If
    L12.Visible = True
    Select Case Locale
    Case "CN"
        L12.Caption = "正在获取版本号 ..."
    Case "EN"
        L12.Caption = "Getting latest version ..."
    Case "ES"
        L12.Caption = "Obtener la ultima version ..."
    End Select
    DoEvents

    LatestVersion = Replace(GETString(LatestVersionLink), vbCrLf, "")
    L12.Caption = LatestVersion
    DoEvents

    If Not FSO.FolderExists(InstallPath & "\Downloads") Then FSO.CreateFolder (InstallPath & "\Downloads")

    If FreshInstall Then
        If Not FSO.FileExists(InstallPath & "\Downloads\SMM_WE.exe") Then
            DownloadCompleted = False
            Select Case Locale
            Case "CN"
                L12.Caption = "正在下载原版 SMM:WE 3.2.3 ..."
            Case "EN"
                L12.Caption = "Downloading vanilla SMM:WE 3.2.3 ..."
            Case "ES"
                L12.Caption = "Descargando SMM:WE original 3.2.3 ..."
            End Select
            DownloadProgress.Visible = True
            L13.Visible = True
            DownloadProgress.Value = 0
            Downloader.DownloadFile VanillaLink, InstallPath & "\Downloads\SMM_WE.exe"
            Do While DownloadCompleted = False
                Sleep 500
                DoEvents
            Loop
            DownloadProgress.Visible = False
            L13.Visible = False
        End If
    End If

    Select Case Locale
    Case "CN"
        PatchURL = PatchRoot & "SMM_WE%20%E5%BC%95%E6%93%8E%E9%83%A8%E8%90%BD%E8%A1%A5%E4%B8%81%20" + LatestVersion + "%20(PC%20CN).7z?raw=true"
    Case "EN"
        PatchURL = PatchRoot & "SMM_WE%20Engine%20Tribe%20patch%20" + LatestVersion + "%20(PC%20EN).7z?raw=true"
    Case "ES"
        PatchURL = PatchRoot & "SMM_WE%20Engine%20Tribe%20parche%20" + LatestVersion + "%20(PC%20ES).7z?raw=true"
    End Select
    
    Debug.Print PatchURL
    
    If Not FSO.FileExists(InstallPath & "\Downloads\EnginePatch.7z") Then
        DownloadCompleted = False
        Select Case Locale
        Case "CN"
            L12.Caption = "正在下载引擎部落补丁 ..."
        Case "EN"
            L12.Caption = "Downloading patch of Engine Tribe ..."
        Case "ES"
            L12.Caption = "Descargando parche para Engine Tribe ..."
        End Select
        DownloadProgress.Visible = True
        L13.Visible = True
        DownloadProgress.Value = 0
        Downloader.DownloadFile PatchURL, InstallPath & "\Downloads\EnginePatch.7z"
        Do While DownloadCompleted = False
            Sleep 500
            DoEvents
        Loop
        DownloadProgress.Visible = False
        L13.Visible = False
    End If

    Select Case Locale
    Case "CN"
        L12.Caption = "正在解压缩 ..."
    Case "EN"
        L12.Caption = "Extracting ..."
    Case "ES"
        L12.Caption = "Descomprimiendo ..."
    End Select
    DoEvents

    If FreshInstall Then
        Unzip InstallPath & "\Downloads\SMM_WE.exe", InstallPath
        FSO.DeleteFile InstallPath & "\Downloads\SMM_WE.exe"
    End If
    Unzip InstallPath & "\Downloads\EnginePatch.7z", InstallPath
    FSO.DeleteFile InstallPath & "\Downloads\EnginePatch.7z"

    FSO.DeleteFolder InstallPath & "\$PLUGINSDIR", True
    FSO.DeleteFolder InstallPath & "\Downloads", True
    
    FSO.DeleteFile InstallPath & "\Usage.txt"
    FSO.DeleteFile InstallPath & "\Uso.txt"
    FSO.DeleteFile InstallPath & "\使用方法.txt"
    FSO.DeleteFile InstallPath & "\uninstall.exe"

    Select Case Locale
    Case "CN"
        MsgBox "安装完成!", vbInformation
    Case "EN"
        MsgBox "Installation completed!", vbInformation
    Case "ES"
        MsgBox "Instalacion completa!", vbInformation
    End Select
    End

End Sub

Private Sub Downloader_DownloadComplete()
    DownloadCompleted = True
End Sub


Private Sub Downloader_DownloadProgress(ByVal BytesRead As Long, ByVal BytesTotal As Long)
'刷新进度条
    L13.Caption = Format(BytesRead / BytesTotal, "Percent") & " " & CStr(FormatNumber(BytesRead / 1048576, 2, vbTrue)) & "MB / " & CStr(FormatNumber(BytesTotal / 1048576, 2, vbTrue)) & "MB"
    DownloadProgress.Value = CInt(BytesRead / BytesTotal * 100)
End Sub

