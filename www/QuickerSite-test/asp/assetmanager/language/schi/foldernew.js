function getTxt(s)
	{
	switch(s)
		{
		case "Folder already exists.": return "\u8d44\u6599\u5939\u5df2\u5b58\u5728 .";
		case "Folder created.": return "\u8d44\u6599\u5939\u5df1\u5efa\u7acb .";
		case "Invalid input.": return "\u8f93\u5165\u4e0d\u6b63\u786e .";
		}
	}	
function loadTxt()
	{
    document.getElementById("txtLang").innerHTML = "\u65b0\u589e\u8d44\u6599\u5939\u540d\u79f0 ";
    document.getElementById("btnCloseAndRefresh").value = "\u5173\u95ed\u53ca\u66f4\u65b0 ";
    document.getElementById("btnCreate").value = "\u5efa\u7acb ";
	}
function writeTitle()
	{
	document.write("<title>\u5efa\u7acb\u8d44\u6599\u5939 </title>")
	}
