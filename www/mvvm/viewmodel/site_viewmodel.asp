<%
Function BuildHomeViewModel()
    Dim json
    Dim viewModel
    Dim modelData
    Dim metricItems

    Set json = Server.CreateObject("G3JSON")
    Set viewModel = json.NewObject()

    Set modelData = LoadHomeModelData()
    Set metricItems = LoadHomeMetricItems()

    viewModel("PageTitle") = "MVVM Home"
    viewModel("Header") = modelData("SiteName")
    viewModel("Subtitle") = modelData("Summary")
    viewModel("Metrics") = metricItems

    Set BuildHomeViewModel = viewModel

    Set metricItems = Nothing
    Set modelData = Nothing
    Set viewModel = Nothing
    Set json = Nothing
End Function

Function BuildExampleViewModel()
    Dim json
    Dim viewModel
    Dim records

    Set json = Server.CreateObject("G3JSON")
    Set viewModel = json.NewObject()
    Set records = LoadExampleModelData()

    viewModel("PageTitle") = "MVVM Example"
    viewModel("Header") = "MVVM Example Page"
    viewModel("Description") = "The ViewModel shapes model records into fields consumed directly by the view."
    viewModel("Rows") = records

    Set BuildExampleViewModel = viewModel

    Set records = Nothing
    Set viewModel = Nothing
    Set json = Nothing
End Function

Function RenderMvvmHome()
    Dim templateEngine
    Dim viewModel
    Dim htmlOutput

    Set templateEngine = Server.CreateObject("G3TEMPLATE")
    Set viewModel = BuildHomeViewModel()

    htmlOutput = templateEngine.Render("view/home.html", viewModel)
    RenderMvvmHome = htmlOutput

    Set viewModel = Nothing
    Set templateEngine = Nothing
End Function

Function RenderMvvmExample()
    Dim templateEngine
    Dim viewModel
    Dim htmlOutput

    Set templateEngine = Server.CreateObject("G3TEMPLATE")
    Set viewModel = BuildExampleViewModel()

    htmlOutput = templateEngine.Render("view/example.html", viewModel)
    RenderMvvmExample = htmlOutput

    Set viewModel = Nothing
    Set templateEngine = Nothing
End Function
%>