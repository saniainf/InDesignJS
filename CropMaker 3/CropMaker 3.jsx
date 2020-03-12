/* =========================================================================================

NAME: CropMaker
VERSION: 3.3.6

AUTHOR: Николаев Алексей
DATE  : 22.05.2015

DESCRIPTION:
Скрипт создает на отдельном слое метки резки
E-MAIL: cropmaker@mail.ru

=========================================================================================*/

#target indesign
#targetengine "session"

var myScript = loadScript("CropMaker 3.3.jsxbin");
if(myScript != undefined) {  eval(myScript); }

 function getScriptsFolderPath() {
    try { var script = app.activeScript;
    } catch(e) {  var script = File(e.fileName);   }
    return script.path + "/";
}
 
function loadScript (filename) {
    var file = File(getScriptsFolderPath() + "/Resources/" + filename);
    var script = undefined;
    if (file.exists) { file.open();
        script = file.read();
        file.close();
    } else {
        alert("Failed to locate the file: " + filename);
        exit();
    }
    return script;
}