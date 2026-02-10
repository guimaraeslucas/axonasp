function getTxt(s)
	{
    switch(s)
        {
        case "Cannot delete Asset Base Folder.":return "Du kan ikke slette fil rodmappen.";
        case "Delete this file ?":return "Vil du slette filen ?";
        case "Uploading...":return "Oploader...";//or "Sender..."
        case "File already exists. Do you want to replace it?":return "Fil eksisterer allerede. \u00D8nsker du at overskrive den?";
		
		case "Files": return "Filer"
		case "del": return "Slet"
		case "Empty...": return "Tom..."
        }
    }
function loadTxt()
    {
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "Opret&nbsp;mappe";
    txtLang[1].innerHTML = "Slet&nbsp;mappe";
    txtLang[2].innerHTML = "Send fil";
    
	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "Alle filer";
    optLang[1].text = "Medier";
    optLang[2].text = "Billeder";
    optLang[3].text = "Flash";

    document.getElementById("btnOk").value = " ok ";
    document.getElementById("btnUpload").value = "Upload";
    }
function writeTitle()
    {
    document.write("<title>Filmanager</title>")
    }
