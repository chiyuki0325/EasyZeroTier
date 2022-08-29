VERSION 5.00
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "vbccr17.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EasyZeroTier | 0.0.1 By 是一刀斩哒"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10080
   BeginProperty Font 
      Name            =   "Microsoft YaHei UI"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   10080
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton btnStartMoon 
      Caption         =   "使用本机自建服务器"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton btnLeaveMoon 
      Caption         =   "退出自建服务器"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton btnJoinMoon 
      Caption         =   "加入自建服务器"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton btnExitNetwork 
      Caption         =   "退出网络"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton btnRefreshNetworks 
      Caption         =   "刷新"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VBCCR17.ListView lstNetworks 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2990
      VisualTheme     =   1
      View            =   3
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
   End
   Begin VB.CommandButton btnJoinNetwork 
      Caption         =   "加入网络"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VBCCR17.ListView lstPeers 
      Height          =   4215
      Left            =   4560
      TabIndex        =   6
      Top             =   960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7435
      VisualTheme     =   1
      View            =   3
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HideColumnHeaders=   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "成员列表"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblServerList 
      BackStyle       =   0  'Transparent
      Caption         =   "网络列表"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblZeroTierVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblZeroTierVersion"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblMyAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "myaddress"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AddressCaption As String

Private Sub btnExitNetwork_Click()
'离开网络
If MsgBox("请问是否要退出网络 " & lstNetworks.SelectedItem.Text & " ?", vbQuestion + vbYesNo) = vbYes Then
ZeroTierCLI.Query ("leave " & lstNetworks.SelectedItem.key)
UpdateNetworks
End If
End Sub

Private Sub btnJoinMoon_Click()
'加入moon
If Dir(ZeroTierPath & "\moons.d", vbDirectory) = "" Then MkDir ZeroTierPath & "\moons.d"
Dim MoonFile As String, MoonFileName As String, MoonID As String, RetVal As String
MoonFile = ""
Do While MoonFile = ""
    MoonFile = ChooseFile("选择自建服务器签名文件", "自建服务器签名文件", "*.moon", Me.hwnd)
Loop
MoonFileName = GetFileNameFromPath(MoonFile)
MoonID = Right(Left(MoonFileName, 16), 10)
FileCopy MoonFile, ZeroTierPath & "\moons.d\" & MoonFileName
Sleep 100
RetVal = ZeroTierCLI.QueryRaw("orbit " & MoonID & " " & MoonID)
If Left(RetVal, 3) = "200" Then
MsgBox "加入完毕。", vbInformation
Else
MsgBox RetVal, vbCritical
End If
    UpdateNetworks
    UpdatePeers
End Sub

Private Sub btnLeaveMoon_Click()
'退出moon
If Dir(ZeroTierPath & "\moons.d", vbDirectory) = "" Then MkDir ZeroTierPath & "\moons.d"
Dim MoonFile As String, MoonFileName As String, MoonID As String, RetVal As String
MoonFile = ""
Do While MoonFile = ""
    MoonFile = ChooseFile("选择自建服务器签名文件", "自建服务器签名文件", "*.moon", Me.hwnd)
Loop
MoonFileName = GetFileNameFromPath(MoonFile)
MoonID = Right(Left(MoonFileName, 16), 10)
FileCopy MoonFile, ZeroTierPath & "\moons.d\" & MoonFileName
Sleep 100
RetVal = ZeroTierCLI.QueryRaw("deorbit " & MoonID)
If InStr(RetVal, "true") Then
MsgBox "退出完毕。", vbInformation
Else
MsgBox RetVal, vbCritical
End If
    UpdateNetworks
    UpdatePeers
End Sub

Private Sub btnRefreshNetworks_Click()
    UpdateNetworks
    UpdatePeers
End Sub

Private Sub btnStartMoon_Click()
If MsgBox("是否用本机自建服务器?" & vbCrLf & "请确保打开光猫的 UPnP 功能！" & vbCrLf & vbCrLf & "EasyZeroTier 对创建服务器的支持是实验性的，可能提示创建成功，但仍然失败。", vbQuestion + vbYesNo) = vbNo Then Exit Sub
btnStartMoon.Caption = "正在测试 IP"

Dim Tmp As Variant
Dim TestedIP As String

With New MSXML2.ServerXMLHTTP30
.Open "GET", "http://api.ip.sb", False
.setRequestHeader "User-Agent", "curl"
.send
TestedIP = Replace(Replace(.responseText, vbLf, ""), vbCr, "")
End With

If Left(TestedIP, 2) = 10 Then
MsgBox "本机没有公网 IP，无法自建 ZeroTier 服务器。", vbCritical
GoTo ExitCreateMoon
End If

btnStartMoon.Caption = "正在尝试 UPnP 打洞"

Dim LocalIP As String
ShellAndWait "cmd /c ipconfig /all > " & QUOT & Environ("LocalAppData") & "\Temp\EZT.txt" & QUOT
Tmp = Split(Filter(Split(ReadTextFile(Environ("LocalAppData") & "\Temp\EZT.txt"), vbCrLf), "IPv4")(0), " ")
LocalIP = Split(Tmp(UBound(Tmp) - 1), "(")(0)

ShellAndWait "cmd /c " & QUOT & QUOT & App.Path & "\upnpc-shared.exe" & QUOT & " -a " & LocalIP & " 9993 9993 udp -i > " & QUOT & Environ("LocalAppData") & "\Temp\EZT.txt" & QUOT & " 2> " & QUOT & Environ("LocalAppData") & "\Temp\EZT2.txt" & QUOT & QUOT
Tmp = ReadTextFile(Environ("LocalAppData") & "\Temp\EZT.txt") & vbCrLf & ReadTextFile(Environ("LocalAppData") & "\Temp\EZT2.txt")

MsgBox "尝试 UPnP 打洞，日志如下" & vbCrLf & "======================================" & vbCrLf & Tmp

Dim MoonJSON As String, MoonID As String
If InStr(Tmp, "is redirected to internal") Then
'开始走moon流程
    MoonJSON = ZeroTierCLI.IDToolRaw("initmoon " & QUOT & ZeroTierPath & "\identity.public" & QUOT)
    MoonJSON = Replace(MoonJSON, QUOT & "stableEndpoints" & QUOT & ": []", QUOT & "stableEndpoints" & QUOT & ": [" & QUOT & TestedIP & "/9993" & QUOT & "]")
    Open App.Path & "\moon.json" For Output As #2
        Print #2, MoonJSON;
    Close #2
    MoonID = JSON.parse(MoonJSON)("id")
    MsgBox "创建服务器签名文件，日志如下" & vbCrLf & ZeroTierCLI.IDToolRaw("genmoon " & QUOT & App.Path & "\moon.json" & QUOT)
    'FileCopy ZeroTierPath & "\000000" & MoonID & ".moon", App.Path & "\000000" & MoonID & ".moon"
    MsgBox "自建服务器成功！" & vbCrLf & "本机自建服务器的签名文件已经生成在程序文件夹，发送给别人即可使用，自己也可以加入。" & vbCrLf & vbCrLf & "000000" & MoonID & ".moon", vbInformation
Else
    MsgBox "自建服务器失败！" & vbCrLf & "当前网络环境不可以使用 UPnP 打洞。", vbCritical
End If

ExitCreateMoon:
btnStartMoon.Caption = "使用本机自建服务器"
Exit Sub
End Sub

Private Sub Form_Load()
    ZeroTierCLI.Init
    InitUI
    UpdateNetworks
    UpdatePeers
End Sub

Private Sub InitUI()
    Dim QueryInfo As Object
On Error GoTo ReInit
ReInit:
    Set QueryInfo = ZeroTierCLI.Query("info")
    AddressCaption = "设备地址: " & QueryInfo("address")
    lblMyAddress.Caption = AddressCaption
    lblZeroTierVersion.Caption = "ZeroTier " & QueryInfo("version")
    With lstPeers
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, "name", "设备名称", 2500
        .ColumnHeaders.Add 2, "role", "角色", 1700
        .ColumnHeaders.Add 3, "latency", "延迟", 1150
    End With
    
    With lstNetworks
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, "name", "网络名称", 3000
        .ColumnHeaders.Add 2, "status", "状态", 1150
    End With
End Sub

Private Sub UpdateNetworks()
On Error GoTo ExitSub
    Dim NetworksQuery As Object, Network As Variant
    Dim i As Integer
    Set NetworksQuery = ZeroTierCLI.Query("listnetworks")
    With lstNetworks
        .ListItems.Clear
        If JSON.toString(NetworksQuery) = "[]" Then
            .ListItems.Add 1, , "当前尚未加入任何网络"
        Else
            i = 1
            For Each Network In NetworksQuery
                If Network("name") = "" Then
                    .ListItems.Add i, Network("nwid"), "(等待管理员同意)"
                Else
                    With New cUTF8
                        lstNetworks.ListItems.Add i, Network("nwid"), .DecodeUTF8(Network("name"))
                    End With
                End If
                .ListItems(i).ToolTipText = .ListItems(i).Text
                .ListItems(i).SubItems(1) = ZeroTierCLI.StatusToReadable(Network("status"))
                i = i + 1
            Next
        End If
    End With
ExitSub:
    Exit Sub
End Sub

Private Sub UpdatePeers()
    Dim MoonExists As Boolean, MoonAddress As String
    MoonExists = False
    Dim PeersQuery As Object, Peer As Variant
    Dim i As Integer
    Set PeersQuery = ZeroTierCLI.Query("listpeers")
    With lstPeers
        .ListItems.Clear
        If JSON.toString(PeersQuery) = "[]" Then
            .ListItems.Add 1, , "当前网络没有成员"
        Else
            i = 1
            For Each Peer In PeersQuery
                .ListItems.Add i, Peer("address"), Peer("address")
                .ListItems(i).ToolTipText = .ListItems(i).Text

                Select Case Peer("role")
                Case "LEAF": .ListItems(i).SubItems(1) = "成员"
                Case "PLANET": .ListItems(i).SubItems(1) = "中心服务器"
                Case "MOON"
                    .ListItems(i).SubItems(1) = "自建服务器"
                    MoonExists = True
                    MoonAddress = Peer("address")
                Case Else: .ListItems(i).SubItems(1) = Peer("role")
                End Select
                
                If Peer("latency") = 0 Then
                    .ListItems(i).SubItems(2) = "(本机)"
                ElseIf Peer("latency") < 0 Then
                    .ListItems(i).SubItems(2) = "(离线)"
                Else
                    .ListItems(i).SubItems(2) = str(Peer("latency")) & "ms"
                End If
                i = i + 1
            Next
        End If
    End With
    If MoonExists Then
        lblMyAddress.Caption = AddressCaption & " | 正在使用自建服务器 " & MoonAddress
    Else
        lblMyAddress.Caption = AddressCaption & " | 正在使用中心服务器"
    End If
End Sub

Private Sub btnJoinNetwork_Click()
'加入网络
    Dim JoinQuery As Object, NetworkID As String
    NetworkID = InputBox("请输入你要加入的网络 ID", "EasyZeroTier", "16 位网络 ID")
    MsgBox ("如果弹出加入新网络的通知，询问“是否允许你的设备被其他设备发现，请点击【允许】。”")
    Set JoinQuery = ZeroTierCLI.Query("join " & NetworkID)
    If JSON.toString(JoinQuery) = "" Then
        MsgBox "加入网络失败，可能是网络 ID 不正确。", vbCritical
    Else
        If JoinQuery("type") = "PRIVATE" Then
            MsgBox "加入成功，请等待网络管理员确认。", vbInformation
        Else
            MsgBox "加入成功。", vbInformation
        End If
    End If
UpdateNetworks
End Sub

Private Sub lstPeers_ContextMenu(ByVal x As Single, ByVal Y As Single)
    Dim PeersQuery As Object, Peer As Variant, Path As Variant, WndName As String, pid As Long, tmphWnd As Long
    Set PeersQuery = ZeroTierCLI.Query("listpeers")
    For Each Peer In PeersQuery
        If Peer("address") = lstPeers.SelectedItem.Text Then
            If Peer("role") = "LEAF" Then
                For Each Path In Peer("paths")
                    If Path("active") And Not Path("expired") Then
                        WndName = "测试到 " & Peer("address") & " 的连接延迟"
                        pid = Shell("cmd /c mode con cols=75 lines=20 && color f0 && title " & WndName & " && echo.正在测试连接延迟... && echo.如果延迟过大说明是中心服务器中转，&& echo.请考虑自建服务器 & ping " & Split(Path("address"), "/")(0) & " && pause >nul", vbNormalFocus)
                        Exit Sub
                    End If
                Next
                MsgBox "该成员不在线，未能与其建立可靠连接，无法测试延迟", vbCritical
            Else
                For Each Path In Peer("paths")
                    If Path("active") And Not Path("expired") Then
                        WndName = "测试到服务器 " & Peer("address") & " 的连接延迟"
                        pid = Shell("cmd /c mode con cols=75 lines=20 && color f0 && title " & WndName & " && echo.正在测试连接延迟... && echo.到中心服务器的延迟过高是正常现象，&& echo.如果无法游玩请考虑自建服务器 & ping " & Split(Path("address"), "/")(0) & " && pause >nul", vbNormalFocus)
                        Exit Sub
                    End If
                Next
                MsgBox "这台服务器不在线，未能与其建立可靠连接，无法测试延迟", vbCritical
            End If
        End If
    Next
End Sub
