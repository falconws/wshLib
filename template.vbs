Option Explicit

' ���C�u�����̓ǂݍ���
Const FOR_READ = 1
Const LIB_DIR = "lib"

Dim objFso
Set objFso = CreateObject("Scripting.FileSystemObject")

' OS���C�u����
Execute objFso.OpenTextFile(objFso.BuildPath(LIB_DIR, "OS.vbs"), FOR_READ, False).ReadAll()

' Util���C�u����
Execute objFso.OpenTextFile(objFso.BuildPath(LIB_DIR, "Util.vbs"), FOR_READ, False).ReadAll()