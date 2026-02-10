function getTxt(s)
    {
    switch(s)
        {
        case "Cannot delete Asset Base Folder.":return "Basismap kan niet worden verwijderd.";
        case "Delete this file ?":return "Dit bestand verwijderen ?";
        case "Uploading...":return "Uploading...";
        case "File already exists. Do you want to replace it?":return "Dit bestand bestaat reeds. Wilt U het vervangen ?"

		case "Files": return "Bestanden";
		case "del": return "Verwijderen";
		case "Empty...": return "Leeg...";
        }
    }
function loadTxt()
    {
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "Nieuwe&nbsp;Map";
    txtLang[1].innerHTML = "Verwijder&nbsp;Map";
    txtLang[2].innerHTML = "Bestand plaatsen";

	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "*";
    optLang[1].text = "Media";
    optLang[2].text = "Afbeelding";
    optLang[3].text = "Flash";

    document.getElementById("btnOk").value = " ok ";
    document.getElementById("btnUpload").value = "upload";
    }
function writeTitle()
    {
    document.write("<title>Media Beheer</title>")
    }