function getTxt(s)
	{
	switch(s)
		{
		case "Cannot delete Asset Base Folder.":return "\u4e0d\u80fd\u5220\u9664\u6b64\u8d44\u6599\u5939 .";
		case "Delete this file ?":return "\u5220\u9664\u6b64\u6863\u6848  ?";
		case "Uploading...":return "\u4e0a\u8f7d\u4e2d ...";
		case "File already exists. Do you want to replace it?":return "\u6863\u6848\u5df2\u5b58\u5728\uff0c\u8981\u50a8\u5b58\u5e76\u66ff\u4ee3\u65e7\u6709\u6863\u6848\u5417 ?";
				
		case "Files": return "\u6863\u6848 ";
		case "del": return "\u5220\u9664 ";
		case "Empty...": return "\u7a7a\u767d ...";
		}
	}
function loadTxt()
	{
	var txtLang = document.getElementsByName("txtLang");
	txtLang[0].innerHTML = "\u65b0\u589e\u8d44\u6599\u5939 ";
	txtLang[1].innerHTML = "\u5220\u9664\u8d44\u6599\u5939 ";
	txtLang[2].innerHTML = "\u4e0a\u8f7d\u6863\u6848 ";
	
	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "\u5168\u90e8\u6863\u6848 ";
    optLang[1].text = "\u591a\u5a92\u4f53\u6863\u6848 ";
    optLang[2].text = "\u5f71\u50cf ";
    optLang[3].text = "Flash";
	
    document.getElementById("btnOk").value = " \u786e\u8ba4  ";
    document.getElementById("btnUpload").value = "\u4e0a\u8f7d ";
	}
function writeTitle()
    {
    document.write("<title>Asset manager</title>")
    }
