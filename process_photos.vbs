Option Explicit

Dim fso, folderOld, folderNew, file, srcPath, dstPath, cmd
Dim WshShell, jpText

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' ★ 日本語文字列は変数に入れる（これが重要）
jpText = "(C) Tsurukawa @ Showa"

Dim basePath
basePath = "C:\Users\h_nis\Pictures\photo-map-site"

folderOld = fso.BuildPath(basePath, "oldphotos")
folderNew = fso.BuildPath(basePath, "newphotos")

If Not fso.FolderExists(folderOld) Then
    WScript.Echo "Error: old folder not found: " & folderOld
    WScript.Quit
End If

If Not fso.FolderExists(folderNew) Then
    fso.CreateFolder(folderNew)
End If

For Each file In fso.GetFolder(folderOld).Files
    If LCase(fso.GetExtensionName(file.Name)) = "jpg" Then

        srcPath = fso.BuildPath(folderOld, file.Name)
        dstPath = fso.BuildPath(folderNew, file.Name)

        ' ★ 日本語を直接書かず、変数 jpText を使う
' ImageMagick command (safe concat version)
cmd = ""
cmd = cmd & "magick """
cmd = cmd & srcPath
cmd = cmd & """ -strip"
cmd = cmd & " -gravity southeast"
cmd = cmd & " -pointsize 64"
cmd = cmd & " -font ""Meiryo"""
cmd = cmd & " -fill black"
cmd = cmd & " -undercolor white"
cmd = cmd & " -annotate +40+40 """
cmd = cmd & jpText
cmd = cmd & """ """
cmd = cmd & dstPath
cmd = cmd & """"


        WshShell.Run cmd, 0, True
    End If
Next

WScript.Echo "Done."
