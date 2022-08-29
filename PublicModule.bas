Attribute VB_Name = "PublicModule"
Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
'Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Const QUOT As String = """"
Public Sub ShellAndWait(pathFile As String)
    With CreateObject("WScript.Shell")
        .Run pathFile, 0, True
    End With
End Sub


Public Function ChooseFile(ByVal frmTitle As String, ByVal fileDescription As String, ByVal fileFilter As String, ByVal onForm As Variant) As String
'oleexp Ñ¡ÔñÎÄ¼þ
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
        .Show onForm
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

Public Function ReadTextFile(sFilePath As String) As String
    On Error Resume Next
    Dim handle As Integer
    If LenB(Dir$(sFilePath)) > 0 Then
        handle = FreeFile
        Open sFilePath For Binary As #handle
        ReadTextFile = Space$(LOF(handle))
        Get #handle, , ReadTextFile
        Close #handle
    End If
End Function

Public Function GetFileNameFromPath(ByVal strFullPath As String, Optional ByVal strSplitor As String = "\") As String
GetFileNameFromPath = Right$(strFullPath, Len(strFullPath) - InStrRev(strFullPath, strSplitor, , vbTextCompare))
End Function
