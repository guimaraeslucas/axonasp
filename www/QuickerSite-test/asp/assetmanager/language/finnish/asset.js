function getTxt(s)
	{
    switch(s)
        {
        case "Cannot delete Asset Base Folder.":return "P\u00E4\u00E4hakemistoa ei voi poistaa";
        case "Delete this file ?":return "Haluatko poistaa tiedoston?";
        case "Uploading...":return "Lataa...";
        case "File already exists. Do you want to replace it?":return "File already exists. Do you want to replace it?"
				
		case "Files": return "Tiedostot";
		case "del": return "Poista";
		case "Empty...": return "Empty...";
        }
    }
function loadTxt()
    {
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "Uusi&nbsp;kansio";
    txtLang[1].innerHTML = "Poista&nbsp;kansio";
    txtLang[2].innerHTML = "Lataa tiedosto";

	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "Kaikki tiedostot";
    optLang[1].text = "Media";
    optLang[2].text = "Kuvat";
    optLang[3].text = "Flash";

    document.getElementById("btnOk").value = " OK ";
    document.getElementById("btnUpload").value = "Lataa";
    }
function writeTitle()
    {
    document.write("<title>Tiedostomanageri</title>")
    }