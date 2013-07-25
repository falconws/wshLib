Option Explicit

'Windows OSのバージョンを返す
'存在しない場合自動的に作成する
'@return
'	OSのバージョンを示す文字列
Function getOS()
	Dim OSInfoCollection, OSInfo, retOS

	Set OSInfoCollection = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")

	For Each OSInfo In OSInfoCollection
    	retOS = OSInfo.Caption
	Next
	
	Set OSInfoCollection = Nothing
	getOS = retOS
End Function

'AllUsersのスタートアップへプログラムを登録する
'@param workDir 実行するプログラムが格納されているフォルダへの絶対パス
'@param execFileName 実行するプログラムのファイル名
'@param windowStyle ショートカット実行時のウィンドウサイズ 1:普通 3:最大化 7:最小化
Sub registAllUsersStartup(workDir, execFileName, windowStyle)
	Dim objWsh
	Dim startupDir, oShellLink
	
	Set objWsh = CreateObject("WScript.Shell")
	
	rem スタートアップディレクトリの取得
	startupDir = objWsh.SpecialFolders("AllUsersStartup")

	rem スタートアップへ登録
	Set oShellLink = objWsh.CreateShortcut(startupDir & "\" & execFileName & ".lnk")
	oShellLink.TargetPath = workDir & "\" & execFileName
	oShellLink.WorkingDirectory = workDir

	oshellLink.WindowStyle = windowStyle

	oshellLink.Save
	
	Set objWsh = Nothing
	Set oShellLink = Nothing
End Sub

'指定されたプロセス名のプロセスを強制終了する
'@param processName プロセス名
Function killProcessByName(processName)
	Dim objLocator, objService, colProcSet, objProc
	
	'オブジェクトの生成
	Set objLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
	Set objService = objLocator.ConnectServer
	
	Set colProcSet = objService.ExecQuery("Select * From Win32_Process Where Caption='" & processName & "'")
	For Each objProc In colProcSet
		objProc.Terminate
	Next
End Function

'OSをシャットダウンする
Sub shutdown()
	Dim objShutdown, objOS, objSystem
	
	' シャットダウンオブジェクト取得
	Set objShutdown = GetObject("winmgmts:{impersonationLevel = impersonate, (Shutdown)}")
	' OSのオブジェクト取得
	Set objOS = objShutdown.InstancesOf("Win32_OperatingSystem")
 
	' シャットダウン
	For Each objSystem In objOS
    	objSystem.Win32Shutdown 8
	Next
End Sub

'OSを再起動する
Sub reboot()
	Dim objShutdown, objOS, objSystem
	
	' シャットダウンオブジェクト取得
	Set objShutdown = GetObject("winmgmts:{impersonationLevel = impersonate, (Shutdown)}")
	' OSのオブジェクト取得
	Set objOS = objShutdown.InstancesOf("Win32_OperatingSystem")
 
	' 再起動
	For Each objSystem In objOS
    	objSystem.Win32Shutdown 2
	Next
End Sub