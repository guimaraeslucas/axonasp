function getTxt(s)
    {
    switch(s)
        {
        case "Cannot delete Asset Base Folder.":return "Kan ikke slette hovedmappen.";
        case "Delete this file ?":return "Vil du slette filen ?";
        case "Uploading...":return "Laster opp til server..."; // or "Sender..."
        case "File already exists. Do you want to replace it?":return "File already exists. Do you want to replace it?"
				
		case "Files": return "Filer";
		case "del": return "Slett";
		case "Empty...": return "Tomt...";
        }
    }
function loadTxt()
    {
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "Opprett&nbsp;mappe";
    txtLang[1].innerHTML = "Slett&nbsp;mappe";
    txtLang[2].innerHTML = "Send fil";

	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "Alle Filer";
    optLang[1].text = "Medier";
    optLang[2].text = "Bilder";
    optLang[3].text = "Flash";
    
    document.getElementById("btnOk").value = " ok ";
    document.getElementById("btnUpload").value = "Send";
    }
function writeTitle()
    {
    document.write("<title>Filbehandler</title>")
    }