function getTxt(s)
    {
    switch(s)
        {
        case "Cannot delete Asset Base Folder.":return "Impossible d'effacer le r\u00E9pertoire racine.";
        case "Delete this file ?":return "Voulez vous effacer ce fichier ?";
        case "Uploading...":return "publication en cours...";
        case "File already exists. Do you want to replace it?":return "File already exists. Do you want to replace it?"
        
		case "Files": return "Fichiers";
		case "del": return "Effacer";
		case "Empty...": return "Empty...";
        }
    }
function loadTxt()
    {
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "Nouveau r\u00E9pertoire";
    txtLang[1].innerHTML = "Effacer le r\u00E9pertoire";
    txtLang[2].innerHTML = "Publier un fichier";

	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "Tous les fichiers";
    optLang[1].text = "Media";
    optLang[2].text = "Images";
    optLang[3].text = "Flash";
    
    document.getElementById("btnOk").value = " ok ";
    document.getElementById("btnUpload").value = "publier";
    }
function writeTitle()
    {
    document.write("<title>Gestionnaire de fichiers</title>")
    }