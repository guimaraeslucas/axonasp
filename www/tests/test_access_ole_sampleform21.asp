<%@ Language=VBScript %>
<!-- #include file="../asplite/asplite.asp" -->
<%
Option Explicit
On Error Resume Next

Response.ContentType = "text/plain"

' Initialize the database plugin to match default handler behavior
Dim db
Set db = aspL.plugin("database")
db.path = "db/sample.mdb"

' Match sampleform21 behavior using the same includes
aspL("default/asp/includes/jQueryUiFunctions.resx")
aspL("default/asp/datatables/includes.resx")

Dim contact
Set contact = New cls_contact

contact.iId = ""
contact.sText = "TEST_SAVE_" & Replace(Replace(CStr(Now()), ":", ""), " ", "_")
contact.iNumber = 93
contact.dDate = dateFromPicker("15/07/1973")
contact.iCountryID = 143

Dim okInsert
okInsert = contact.save

If Err.Number <> 0 Then
    Response.Write "ERROR AFTER INSERT: " & Err.Description
    Response.End
End If

Dim newId
newId = contact.iId

Dim contactRead
Set contactRead = New cls_contact
contactRead.pick newId

If Err.Number <> 0 Then
    Response.Write "ERROR AFTER PICK: " & Err.Description
    Response.End
End If

Dim okRead
okRead = (CStr(contactRead.sText) = CStr(contact.sText))

contactRead.sText = "TEST_UPDATE_" & Replace(Replace(CStr(Now()), ":", ""), " ", "_")
contactRead.iNumber = 94
contactRead.dDate = dateFromPicker("16/07/1973")
contactRead.iCountryID = 143

Dim okUpdateCall
okUpdateCall = contactRead.save

If Err.Number <> 0 Then
    Response.Write "ERROR AFTER UPDATE: " & Err.Description
    Response.End
End If

Dim contactVerify
Set contactVerify = New cls_contact
contactVerify.pick newId

If Err.Number <> 0 Then
    Response.Write "ERROR AFTER VERIFY PICK: " & Err.Description
    Response.End
End If

Dim okUpdate
okUpdate = (CStr(contactVerify.sText) = CStr(contactRead.sText))

If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description
Else
    If okInsert And okRead And okUpdateCall And okUpdate Then
        Response.Write "PASS"
    Else
        Response.Write "FAIL"
    End If
End If

On Error GoTo 0
%>
