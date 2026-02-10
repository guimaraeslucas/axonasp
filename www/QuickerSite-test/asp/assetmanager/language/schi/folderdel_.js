function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "\u8d44\u6599\u5939\u5df1\u5220\u9664 .";
		case "Folder does not exist.": return "\u8d44\u6599\u5939\u4e0d\u5b58\u5728 .";
		case "Cannot delete Asset Base Folder.": return "\u4e0d\u80fd\u5220\u9664\u6b64\u8d44\u6599\u5939 .";
        }
    }
function loadTxt()
	{
	document.getElementById("btnCloseAndRefresh").value = "\u5173\u95ed\u53ca\u66f4\u65b0 ";
	}
function writeTitle()
	{
	document.write("<title>\u5220\u9664\u8d44\u6599\u5939 </title>")
	}
