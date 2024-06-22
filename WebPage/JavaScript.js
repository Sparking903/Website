ButtonElement = document.getElementById("Test")

function W() {
    ButtonElement.innerText = "Wow"
    var objShell = new ActiveXObject("WScript.Shell");
    objShell.Run("path_to_your_vbscript.vbs", 0, false);
}