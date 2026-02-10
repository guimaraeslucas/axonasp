function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "\u8cc7\u6599\u593e\u5df1\u522a\u9664 .";
		case "Folder does not exist.": return "\u8cc7\u6599\u593e\u4e0d\u5b58\u5728 .";
		case "Cannot delete Asset Base Folder.": return "\u4e0d\u80fd\u522a\u9664\u6b64\u8cc7\u6599\u593e .";
        }
    }
function loadTxt()
	{
	document.getElementById("btnCloseAndRefresh").value = "\u95dc\u9589\u53ca\u66f4\u65b0 ";
	}
function writeTitle()
	{
	document.write("<title>\u522a\u9664\u8cc7\u6599\u593e </title>")
	}