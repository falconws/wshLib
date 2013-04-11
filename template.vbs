Option Explicit

' ライブラリの読み込み
Const FOR_READ = 1
Const LIB_DIR = "lib"

Dim objFso
Set objFso = CreateObject("Scripting.FileSystemObject")

' OSライブラリ
Execute objFso.OpenTextFile(objFso.BuildPath(LIB_DIR, "OS.vbs"), FOR_READ, False).ReadAll()

' Utilライブラリ
Execute objFso.OpenTextFile(objFso.BuildPath(LIB_DIR, "Util.vbs"), FOR_READ, False).ReadAll()