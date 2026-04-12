<%
Function GetSiteMetadata()
    Dim json
    Dim metadata

    Set json = Server.CreateObject("G3JSON")
    Set metadata = json.NewObject()

    metadata("SiteName") = "AxonASP MVC Sample"
    metadata("Tagline") = "Classic ASP architecture with explicit model and controller layers"

    Set GetSiteMetadata = metadata

    Set metadata = Nothing
    Set json = Nothing
End Function

Function GetHighlights()
    Dim highlights

    highlights = Array( _
        "Controller handles routing and rendering decisions", _
        "Model encapsulates data retrieval", _
        "Views are rendered with G3Template" _
        )

    GetHighlights = highlights
End Function

Function GetExampleItems()
    Dim Items

    Items = Array("Product Catalog", "Order List", "Audit Timeline")

    GetExampleItems = Items
End Function
%>