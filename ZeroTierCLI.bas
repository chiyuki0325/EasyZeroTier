Attribute VB_Name = "ZeroTierCLI"
Public ZeroTierPath As String
Public Const ZeroTierExe = "zerotier-one_x64.exe"

Sub Init()
    ZeroTierPath = Environ("SystemDrive") & "\ProgramData\ZeroTier\One"
End Sub

Function Query(Command As String) As Object
    Set Query = JSON.parse(ZeroTierShell("-q " & Command & " -j"))
End Function

Function IDToolRaw(Command As String) As String
    IDToolRaw = ZeroTierShell("-i " & Command)
End Function

Private Function ZeroTierShell(SubCommand As String) As String
    ShellAndWait ("cmd /c " & QUOT & QUOT & ZeroTierPath & "\" & ZeroTierExe & QUOT & " " & SubCommand & " >" & QUOT & Environ("LocalAppData") & "\ZTCLI.txt" & QUOT & QUOT)
    ZeroTierShell = ReadTextFile(Environ("LocalAppData") & "\ZTCLI.txt")
End Function

Function StatusToReadable(ByVal Status As String) As String
    Select Case Status
        Case "ACCESS_DENIED": StatusToReadable = "等待确认"
        Case "REQUESTING_CONFIGURATION": StatusToReadable = "正在加入"
        Case "OK": StatusToReadable = "正常"
        Case Else: StatusToReadable = Status
    End Select
End Function


Function QueryRaw(Command As String) As String
    QueryRaw = ZeroTierShell("-q " & Command & " -j")
End Function


