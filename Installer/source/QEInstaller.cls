VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QEInstaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const MACRONAME = "QuickExport"
Private Const CLASSNAME = "QuickExport.Ui"


Private Sub GlobalMacroStorage_Start()
  If GetSetting("cdrpro.ru", "macrosX7", MACRONAME & "Installed", "0") = "1" Then
    Call DeleteGMS
  Else
    Dim msg&: msg = MsgBox("Install " & MACRONAME & "?", vbOKCancel, MACRONAME & " installer")
    If msg = "1" Then
      Call Install
    Else
      Dim msgd&: msgd = MsgBox("Delete " & MACRONAME & " installer?", vbOKCancel, MACRONAME & " installer")
      If msgd = "1" Then Call DeleteGMS
    End If
  End If
End Sub

Private Sub DeleteGMS()
  Dim path$: path = GMSManager.UserGMSPath & MACRONAME & "Installer.gms"
  Dim file$: file = FileSystem.Dir(path)
  If file <> "" Then FileSystem.Kill path
End Sub

Private Sub Install()
  Dim dllPath$

  If InStr(1, AppWindow.Caption, "64", vbTextCompare) > 0 Then
    dllPath = SetupPath & "Programs64\Addons\" & MACRONAME & "\" & MACRONAME & ".dll"
  Else
    dllPath = SetupPath & "Programs\Addons\" & MACRONAME & "\" & MACRONAME & ".dll"
  End If
  
  If FileSystem.Dir(dllPath) = "" Then
    dllPath = CorelScriptTools.GetFileBox("Assembly File (*.dll)|*.dll", "Open Assembly File", 0, , , "C:\", "Open")
  End If

  If Trim(dllPath) = "" Then Exit Sub

  If FrameWork.CommandBars("Standard").Visible = False Then
    FrameWork.CommandBars("Standard").Visible = True
  End If

  Call FrameWork.CommandBars("Standard").Controls.AddCustomControl(CLASSNAME, dllPath, 0, False)
  
  SaveSetting "cdrpro.ru", "macrosX7", MACRONAME & "Installed", "1"
  Call DeleteGMS
  
  MsgBox MACRONAME & " installed!", vbInformation, MACRONAME & " installer"
End Sub
