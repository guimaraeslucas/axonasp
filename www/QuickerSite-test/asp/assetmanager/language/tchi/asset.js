function getTxt(s)
	{
	switch(s)
		{
		case "Cannot delete Asset Base Folder.":return "\u4e0d\u80fd\u522a\u9664\u6b64\u8cc7\u6599\u593e .";
		case "Delete this file ?":return "\u522a\u9664\u6b64\u6a94\u6848  ?";
		case "Uploading...":return "\u4e0a\u8f09\u4e2d ...";
		case "File already exists. Do you want to replace it?":return "\u6a94\u6848\u5df2\u5b58\u5728\uff0c\u8981\u5132\u5b58\u4e26\u66ff\u4ee3\u820a\u6709\u6a94\u6848\u55ce ?";
				
		case "Files": return "\u6a94\u6848 ";
		case "del": return "\u522a\u9664 ";
		case "Empty...": return "\u7a7a\u767d ...";
		}
	}
function loadTxt()
	{
	var txtLang = document.getElementsByName("txtLang");
	txtLang[0].innerHTML = "\u65b0\u589e\u8cc7\u6599\u593e ";
	txtLang[1].innerHTML = "\u522a\u9664\u8cc7\u6599\u593e ";
	txtLang[2].innerHTML = "\u4e0a\u8f09\u6a94\u6848 ";
	
	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "\u5168\u90e8\u6a94\u6848 ";
    optLang[1].text = "\u591a\u5a92\u9ad4\u6a94\u6848 ";
    optLang[2].text = "\u5f71\u50cf ";
    optLang[3].text = "Flash";
	
    document.getElementById("btnOk").value = " \u78ba\u8a8d  ";
    document.getElementById("btnUpload").value = "\u4e0a\u8f09 ";
	}
function writeTitle()
    {
    document.write("<title>Asset manager</title>")
    }