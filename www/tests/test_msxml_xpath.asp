<%
Dim xmlDoc, books, expensive, firstBook, titleNode
Set xmlDoc = Server.CreateObject("MSXML2.DOMDocument")

xmlDoc.LoadXML "<catalog>" & _
               "<book id='1' price='15.99'><title>XML</title></book>" & _
               "<book id='2' price='25.99'><title>Web</title></book>" & _
               "</catalog>"

Set books = xmlDoc.SelectNodes("//book")
Response.Write "books=" & books.Count & "<br>"

Set expensive = xmlDoc.SelectNodes("//book[@price > '20']")
Response.Write "expensive=" & expensive.Count & "<br>"

Set expensive = xmlDoc.SelectNodes("//book[@price > 20 and starts-with(title, 'W')]")
Response.Write "expensive_startswith_w=" & expensive.Count & "<br>"

Set firstBook = xmlDoc.SelectSingleNode("/catalog/book[1]")
If Not IsEmpty(firstBook) Then
    Set titleNode = firstBook.SelectSingleNode("title")
    If Not IsEmpty(titleNode) Then
        Response.Write "first_title=" & titleNode.Text & "<br>"
    End If
End If

Dim following, unionNodes
Set following = xmlDoc.SelectNodes("/catalog/book[1]/following-sibling::book")
Response.Write "following_books=" & following.Count & "<br>"

Set unionNodes = xmlDoc.SelectNodes("//book | //title")
Response.Write "union_nodes=" & unionNodes.Count & "<br>"

Dim nsDoc, nsBook
Set nsDoc = Server.CreateObject("MSXML2.DOMDocument")
nsDoc.LoadXML "<root xmlns:b='urn:book'><b:book id='1'><b:title>Namespaced</b:title></b:book></root>"
nsDoc.SelectionNamespaces = "xmlns:b='urn:book'"
Set nsBook = nsDoc.SelectSingleNode("//b:book")
If Not IsEmpty(nsBook) Then
    Response.Write "ns_book=1<br>"
Else
    Response.Write "ns_book=0<br>"
End If
%>
