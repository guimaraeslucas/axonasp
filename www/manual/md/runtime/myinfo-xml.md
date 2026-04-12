# MyInfo.xml

## Overview

`MyInfo.xml` is a legacy XML configuration file originally introduced by Microsoft Personal Web Server (PWS). AxonASP provides compatibility support for this file through the `MSWC.MyInfo` object. When an ASP page creates an instance of `MSWC.MyInfo`, the server reads `MyInfo.xml` from the web root and exposes its XML element values as readable object properties.

## File Location

Place `MyInfo.xml` at the root of the configured web root directory. The file is automatically blocked from direct HTTP access by the `blocked_files` setting in `axonasp.toml` — clients cannot download or view the file through the browser.

## File Format

The file uses a flat XML structure with a `<MyInfo>` root element. Each child element name becomes an accessible property on the object:

```xml
<MyInfo>
  <PersonalName>Site Administrator</PersonalName>
  <PersonalAddress>123 Server Lane, Datacenter City</PersonalAddress>
  <PersonalPhone>+1 555-0100</PersonalPhone>
  <PersonalMail>admin@example.com</PersonalMail>
  <PersonalWords>Welcome to our site!</PersonalWords>
  <CompanyName>My Company</CompanyName>
  <URL1>http://www.example.com</URL1>
  <URLWords1>Company Website</URLWords1>
  <URL2>http://blog.example.com</URL2>
  <URLWords2>Company Blog</URLWords2>
</MyInfo>
```

## Usage

Instantiate the object with `Server.CreateObject` using PROGID `MSWC.MyInfo`:

```asp
<%
Dim mi
Set mi = Server.CreateObject("MSWC.MyInfo")

Response.Write mi.PersonalName & "<br>"
Response.Write mi.CompanyName & "<br>"
Response.Write mi.URL(1) & "<br>"
Response.Write mi.URLWords(1) & "<br>"
%>
```

## Properties

All XML element names in `MyInfo.xml` are exposed as properties on the object. Property names are case-insensitive.

| Property | Type | Description |
|----------|------|-------------|
| PersonalName | String | Administrator or site owner name |
| PersonalAddress | String | Physical or mailing address |
| PersonalPhone | String | Contact phone number |
| PersonalMail | String | Contact email address |
| PersonalWords | String | Welcome message or site tagline |
| CompanyName | String | Organization name |
| URL(n) | String | Indexed URL entry. Access as a method call with the numeric index: `mi.URL(1)` |
| URLWords(n) | String | Indexed URL label. Access as a method call with the numeric index: `mi.URLWords(1)` |
| *(any element)* | String | Any custom element added to `MyInfo.xml` is accessible as a property with the same name |

## Remarks

- Properties are read-only. Assigning a value to a property has no effect and does not modify the XML file.
- The object reads `MyInfo.xml` once at construction time. Changes to the file are reflected on the next `Server.CreateObject("MSWC.MyInfo")` call (or next page request).
- `URL` and `URLWords` entries use a numeric suffix in the XML (`URL1`, `URLWords1`, `URL2`, `URLWords2`). They are accessed in ASP code as method calls with a numeric argument: `mi.URL(1)`, `mi.URLWords(2)`.
- All other custom properties defined as XML child elements can be read directly as named properties.
- The file is resolved using `Server.MapPath("MyInfo.xml")`, which places it relative to the current web root.
- If the XML contains duplicate element names, only the last occurrence is returned.

## Code Example

Display site information from `MyInfo.xml` on a contact page:

```asp
<%
Dim mi
Set mi = Server.CreateObject("MSWC.MyInfo")
%>
<p><strong>Site Administrator:</strong> <%= mi.PersonalName %></p>
<p><strong>Company:</strong> <%= mi.CompanyName %></p>
<p><strong>Contact:</strong> <a href="mailto:<%= mi.PersonalMail %>"><%= mi.PersonalMail %></a></p>
<p><strong>Phone:</strong> <%= mi.PersonalPhone %></p>
<p><strong>Message:</strong> <%= mi.PersonalWords %></p>
<ul>
<%
Dim i
For i = 1 To 5
    Dim url, label
    url = mi.URL(i)
    label = mi.URLWords(i)
    If url <> "" Then
%>
    <li><a href="<%= url %>"><%= label %></a></li>
<%
    End If
Next
%>
</ul>
```
