function getTxt(s)
    {
	switch(s)
		{
		case "Cannot delete Asset Base Folder.":return "Der Hauptordner kann nicht gel\u00F6scht werden.";
        case "Delete this file ?":return "Soll diese Datei gel\u00F6scht werden ?";
        case "Uploading...":return "Es wird hochgeladen...";
        case "File already exists. Do you want to replace it?":return "Datei existiert bereits. Soll sie ersetzt werden?"
				
		case "Files": return "Dateien";
		case "del": return "l&ouml;schen";
		case "Empty...": return "Empty...";
        }
    }
function loadTxt()
    {
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "Neuer&nbsp;Ordner";
    txtLang[1].innerHTML = "Ordner&nbsp;l&ouml;schen";
    txtLang[2].innerHTML = "Datei hochladen";

	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "alle Dateien";
    optLang[1].text = "Media";
    optLang[2].text = "Bilder";
    optLang[3].text = "Flash";
    
    document.getElementById("btnOk").value = " OK ";
    document.getElementById("btnUpload").value = "hochladen";
    }
function writeTitle()
    {
    document.write("<title>Asset manager</title>")
    }