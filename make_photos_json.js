// 簡易 JSON.stringify 実装（配列とオブジェクト、文字列のみ対応）
if (typeof JSON === "undefined") {
  var JSON = {};
}
if (typeof JSON.stringify !== "function") {
  JSON.stringify = function (obj, space) {
    var indent = (typeof space === "number") ? Array(space+1).join(" ") : "";
    if (obj instanceof Array) {
      var arr = [];
      for (var i=0; i<obj.length; i++) {
        arr.push(JSON.stringify(obj[i], space));
      }
      return "[\n" + indent + arr.join(",\n" + indent) + "\n]";
    } else if (typeof obj === "object") {
      var props = [];
      for (var k in obj) {
        if (obj.hasOwnProperty(k)) {
          props.push('"' + k + '": ' + JSON.stringify(obj[k], space));
        }
      }
      return "{\n" + indent + props.join(",\n" + indent) + "\n}";
    } else if (typeof obj === "string") {
      return '"' + obj.replace(/\\/g,"\\\\").replace(/"/g,'\\"') + '"';
    } else {
      return String(obj);
    }
  };
}
// make_photos_json.js
// photos/ フォルダ内のファイル名を取得して
// thumbs/ と photos/ の両方を参照する JSON を生成する。
// 出力は UTF-8 (BOMなし)

var fso = new ActiveXObject("Scripting.FileSystemObject");
var folder = fso.GetFolder("photos");
var e = new Enumerator(folder.Files);

var arr = [];
for (; !e.atEnd(); e.moveNext()) {
    var file = fso.GetFile(e.item());
    var name = file.Name;

    // JPEG/JPGのみ対象
    var ext = name.slice(-4).toLowerCase();
    if (ext === ".jpg" || ext === "jpeg") {
        arr.push({
            "file": name,
            "thumb": "thumbs/" + name,
            "full": "photos/" + name
        });
    }
}

// JSON 文字列化
var jsonText = JSON.stringify(arr, null, 2);

// UTF-8 (BOMなし) で出力
var adTypeText = 2;
var adSaveCreateOverWrite = 2;
var stream = new ActiveXObject("ADODB.Stream");
stream.Type = adTypeText;
stream.Charset = "utf-8";
stream.Open();
stream.WriteText(jsonText);
stream.Position = 0;
// Skip BOM
stream.Position = 3;
var out = new ActiveXObject("ADODB.Stream");
out.Type = adTypeText;
out.Charset = "utf-8";
out.Open();
stream.CopyTo(out);
out.SaveToFile("photos.json", adSaveCreateOverWrite);
out.Close();
stream.Close();

WScript.Echo("photos.json made");
