function getTxt(s)
	{
	switch(s)
		{
		case "Cannot delete Asset Base Folder.":return "Non puoi eliminare questa cartella Risorse.";
		case "Delete this file ?":return "Elimina queto file?";
		case "Uploading...":return "Caricando...";
		case "File already exists. Do you want to replace it?":return "Il File esiste già. Lo vuoi eliminare?";
				
		case "Files": return "Files";
		case "del": return "canc";
		case "Empty...": return "Vuoto..";
		}
	}
function loadTxt()
	{
	var txtLang = document.getElementsByName("txtLang");
	txtLang[0].innerHTML = "Nuova&nbsp;Cartella";
	txtLang[1].innerHTML = "Canc&nbsp;Cartella";
	txtLang[2].innerHTML = "Carica File";
	
	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "Tutti i Files";
    optLang[1].text = "Audio-Video";
    optLang[2].text = "Immagini";
    optLang[3].text = "Flash";
	
    document.getElementById("btnOk").value = " ok ";
    document.getElementById("btnUpload").value = "Carica";
	}
function writeTitle()
    {
    document.write("<title>Gestisci Risorse</title>")
    }