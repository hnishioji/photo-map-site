// make_photos_json.js
var fso = new ActiveXObject("Scripting.FileSystemObject");

var baseFolder = fso.GetParentFolderName(WScript.ScriptFullName);
var photosFolder = fso.BuildPath(baseFolder, "photos");
var outputFile = fso.BuildPath(baseFolder, "photos.json");

if (!fso.FolderExists(photosFolder)) {
    WScript.Echo("photos フォルダが見つかりません: " + photosFolder);
    WScript.Quit(1);
}

var folder = fso.GetFolder(photosFolder);
var e = new Enumerator(folder.Files);
var list = [];

for (; !e.atEnd(); e.moveNext()) {
    var file = e.item();
    var name = fso.GetFileName(file);
    var lower = name.toLowerCase();
    if (lower.match(/\.jpe?g$/)) {
        list.push('"' + name + '"');
    }
}

// JSON を文字列として自前で作成
var json = "[\r\n  " + list.join(",\r\n  ") + "\r\n]";

// 出力
var ts = fso.CreateTextFile(outputFile, true, true);
ts.Write(json);
ts.Close();

WScript.Echo("photos.json を生成しました: " + outputFile + " (" + list.length + " files)");
