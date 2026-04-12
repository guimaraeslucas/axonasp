<%
Function LoadHomeModelData()
    Dim json
    Dim data

    Set json = Server.CreateObject("G3JSON")
    Set data = json.NewObject()

    data("SiteName") = "AxonASP MVVM Sample"
    data("Summary") = "Model provides source data for the ViewModel composition"

    Set LoadHomeModelData = data

    Set data = Nothing
    Set json = Nothing
End Function

Function LoadHomeMetricItems()
    Dim metrics

    metrics = Array( _
        "Templates rendered by G3Template", _
        "ViewModel normalizes display fields", _
        "Model remains independent from HTML" _
        )

    LoadHomeMetricItems = metrics
End Function

Function LoadExampleModelData()
    Dim records

    records = Array("Invoice #1001", "Invoice #1002", "Invoice #1003")

    LoadExampleModelData = records
End Function
%>