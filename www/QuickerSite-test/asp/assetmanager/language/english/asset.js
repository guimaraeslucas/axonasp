function getTxt(s)
	{
	switch(s)
		{
		case "Cannot delete Asset Base Folder.":return "Cannot delete Asset Base Folder.";
		case "Delete this file ?":return "Delete this file ?";
		case "Uploading...":return "Uploading...";
		case "File already exists. Do you want to replace it?":return "File already exists. Do you want to replace it?";
				
		case "Files": return "Files";
		case "del": return "del";
		case "Empty...": return "Empty...";
		}
	}
function loadTxt()
	{
	var txtLang = document.getElementsByName("txtLang");
	txtLang[0].innerHTML = "New&nbsp;Folder";
	txtLang[1].innerHTML = "Del&nbsp;Folder";
	txtLang[2].innerHTML = "Upload File";
	
	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "All Files";
    optLang[1].text = "Media";
    optLang[2].text = "Images";
    optLang[3].text = "Flash";
	
    document.getElementById("btnOk").value = " ok ";
    document.getElementById("btnUpload").value = "upload";
	}
function writeTitle()
    {
    document.write("<title>Asset manager</title>")
    }