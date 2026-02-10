function getTxt(s)
    {
    switch(s)
        {
        case "Cannot delete Asset Base Folder.":return "Kan inte ta bort rotmappen.";
        case "Delete this file ?":return "Ta bort den h\u00E4r filen?";
        case "Uploading...":return "Laddar upp...";
        case "File already exists. Do you want to replace it?":return "File already exists. Do you want to replace it?"
		
		case "Files": return "Filer";
		case "del": return "Ta bort";
		case "Empty...": return "Tom...";
        }
    }
function loadTxt()
    {
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "Ny mapp";
    txtLang[1].innerHTML = "Ta bort mapp";
    txtLang[2].innerHTML = "Ladda upp fil";

	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "Alla filer";
    optLang[1].text = "Media";
    optLang[2].text = "Bilder";
    optLang[3].text = "Flash";

    document.getElementById("btnOk").value = " OK ";
    document.getElementById("btnUpload").value = "Ladda upp";
    }
function writeTitle()
    {
    document.write("<title>Filhanteraren</title>")
    }