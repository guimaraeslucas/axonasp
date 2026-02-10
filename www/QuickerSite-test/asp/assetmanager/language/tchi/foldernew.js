function getTxt(s)
	{
	switch(s)
		{
		case "Folder already exists.": return "\u8cc7\u6599\u593e\u5df2\u5b58\u5728 .";
		case "Folder created.": return "\u8cc7\u6599\u593e\u5df1\u5efa\u7acb .";
		case "Invalid input.": return "\u8f38\u5165\u4e0d\u6b63\u78ba .";
		}
	}	
function loadTxt()
	{
    document.getElementById("txtLang").innerHTML = "\u65b0\u589e\u8cc7\u6599\u593e\u540d\u7a31 ";
    document.getElementById("btnCloseAndRefresh").value = "\u95dc\u9589\u53ca\u66f4\u65b0 ";
    document.getElementById("btnCreate").value = "\u5efa\u7acb ";
	}
function writeTitle()
	{
	document.write("<title>\u5efa\u7acb\u8cc7\u6599\u593e </title>")
	}
