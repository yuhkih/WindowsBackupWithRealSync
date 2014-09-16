'---------------------------------------
' Title: 自動バックアップ用スクリプト
' History: 2009/02/15 Yuhki 
'          2013/08/03 アップデート
'---------------------------------------

'WOL.exe を使用して Landiskを起動する。
Set objShell = WScript.CreateObject("WScript.Shell")
Set objExec = objShell.Exec("C:\Program Files\RealSync\wol.exe 00A0B08260AC")


'LANDISK が起動するまで、120秒間スリープする（単位は、msec)
'Wscript.Echo "sleep します。"
WScript.Sleep(120000)


'LAN Diskに対して ping を実行
' Dim strTargetAddress

strTargetIPAddress = "192.168.11.10"

If GetPingResult(strTargetIPAddress) = False Then
   'Wscript.Echo "pingが通らないので終了します。"
   Wscript.Quit  ' ping が失敗している場合はそのまま終了する。
End If


'ping が通る場合は、LANDISK のある環境にあると判断して、RealSyncを実行する。
'Wscript.Echo "RealSyncを実行します。"
'RealSync を実行する。
'RealSync の更新を開始して、更新の終了を待たずにコマンドは返ってくる。
Set objShell = objShell.Exec("C:\Program Files\RealSync\RealSyncUtl.exe -s")


'終了
WScript.Quit

'-----------------------------------------
' ping 実行関数
' 戻り値: ping 成功  True
'         ping 失敗  False
'-----------------------------------------
Function GetPingResult(strTarget)

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")  
Set objPingResult = objWMIService.ExecQuery _
                    ("Select * from Win32_PingStatus " & _
                         "Where Address = '" & strTarget & "'")  

For Each Ping in objPingResult
    If Ping.StatusCode = 0 Then  
             GetPingResult = True  
    Else  
             GetPingResult = False 
    End If  
Next 

End Function
