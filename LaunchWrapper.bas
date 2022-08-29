Attribute VB_Name = "LaunchWrapper"
Private Declare Function IsUserAnAdmin Lib "shell32" () As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()

Sub main()
    IsUserAnAdmin
    InitCommonControls
    frmMain.Show
End Sub
