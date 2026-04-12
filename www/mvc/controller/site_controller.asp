<%
Function RenderHomePage()
    Dim json
    Dim viewModel
    Dim metadata
    Dim highlights
    Dim templateEngine
    Dim htmlOutput

    Set json = Server.CreateObject("G3JSON")
    Set viewModel = json.NewObject()

    Set metadata = GetSiteMetadata()
    Set highlights = GetHighlights()

    viewModel("PageTitle") = "MVC Home"
    viewModel("Heading") = metadata("SiteName")
    viewModel("Tagline") = metadata("Tagline")
    viewModel("Highlights") = highlights
    viewModel("CurrentYear") = CStr(Year(Now()))

    Set templateEngine = Server.CreateObject("G3TEMPLATE")
    htmlOutput = templateEngine.Render("view/home.html", viewModel)

    RenderHomePage = htmlOutput

    Set templateEngine = Nothing
    Set highlights = Nothing
    Set metadata = Nothing
    Set viewModel = Nothing
    Set json = Nothing
End Function

Function RenderExamplePage()
    Dim json
    Dim viewModel
    Dim Items
    Dim templateEngine
    Dim htmlOutput

    Set json = Server.CreateObject("G3JSON")
    Set viewModel = json.NewObject()
    Set Items = GetExampleItems()

    viewModel("PageTitle") = "MVC Example"
    viewModel("Heading") = "MVC Example Page"
    viewModel("Message") = "This page is rendered by controller logic and data from model functions."
    viewModel("Items") = Items

    Set templateEngine = Server.CreateObject("G3TEMPLATE")
    htmlOutput = templateEngine.Render("view/example.html", viewModel)

    RenderExamplePage = htmlOutput

    Set templateEngine = Nothing
    Set Items = Nothing
    Set viewModel = Nothing
    Set json = Nothing
End Function
%>