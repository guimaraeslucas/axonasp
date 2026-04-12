<%
' MSWC Library Demo
' AxonASP Server Components Demonstration
%>
<!DOCTYPE html>
<html>
<head>
<title>AxonASP MSWC Components Demo</title>
<style>
    body { background-color: #ECE9D8; font-family: Tahoma, Verdana, Segoe UI, sans-serif; margin: 0; }
    .header { background: linear-gradient(to right, #003399, #3366CC); color: white; padding: 10px 20px; border-bottom: 2px solid #808080; }
    .header h1 { margin: 0; font-size: 24px; }
    .container { padding: 20px; }
    .section { background: white; border: 1px solid #808080; margin-bottom: 20px; padding: 15px; }
    .section-title { font-weight: bold; border-bottom: 2px solid #3366CC; padding-bottom: 5px; margin-bottom: 15px; color: #003399; }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    table th, table td { border: 1px solid #CCC; padding: 8px; text-align: left; }
    table th { background-color: #EEE; }
    code { background: #F4F4F4; border-left: 3px solid #3366CC; display: block; padding: 10px; margin: 10px 0; font-family: 'Courier New', Courier, monospace; }
    .property-name { font-weight: bold; color: #003399; }
</style>
</head>
<body>

<div class="header">
    <h1>AxonASP MSWC Server Components</h1>
    <div style="font-size: 11px;">Expert Classic ASP Implementation - Retro Demo</div>
</div>

<div class="container">

    <!-- 1. MSWC.Counters -->
    <div class="section">
        <div class="section-title">1. MSWC.Counters</div>
        <p>The Counters component creates a persistent counter for each named item. Values are saved in <code>temp/counters.txt</code>.</p>
        <%
        Set cnt = Server.CreateObject("MSWC.Counters")
        cnt.Increment "PageHits"
        page_hits = cnt.Get("PageHits")
        %>
        <code>
            Set cnt = Server.CreateObject("MSWC.Counters")<br>
            cnt.Increment "PageHits"<br>
            page_hits = cnt.Get("PageHits")
        </code>
        <table>
            <tr><th width="200">Counter Name</th><th>Value</th></tr>
            <tr><td><span class="property-name">PageHits</span></td><td><%= page_hits %></td></tr>
        </table>
    </div>

    <!-- 2. MSWC.AdRotator -->
    <div class="section">
        <div class="section-title">2. MSWC.AdRotator</div>
        <p>The AdRotator component rotates banner advertisements based on a schedule file (<code>ad_demo.txt</code>).</p>
        <%
        Set ad = Server.CreateObject("MSWC.AdRotator")
        ad.TargetFrame = "_blank"
        ad.Border = 2
        %>
        <code>
            Set ad = Server.CreateObject("MSWC.AdRotator")<br>
            ad.TargetFrame = "_blank"<br>
            ad.Border = 2<br>
            Response.Write ad.GetAdvertisement("ad_demo.txt")
        </code>
        <div style="text-align:center; padding: 20px; border: 1px dashed #AAA;">
            <%= ad.GetAdvertisement("ad_demo.txt") %>
        </div>
    </div>

    <!-- 3. MSWC.BrowserType -->
    <div class="section">
        <div class="section-title">3. MSWC.BrowserType</div>
        <p>Determines the browser's capabilities by analyzing the User-Agent header.</p>
        <%
        Set browser = Server.CreateObject("MSWC.BrowserType")
        %>
        <code>
            Set browser = Server.CreateObject("MSWC.BrowserType")<br>
            Response.Write browser.browser
        </code>
        <table>
            <tr><th width="200">Capability</th><th>Value</th></tr>
            <tr><td><span class="property-name">Browser</span></td><td><%= browser.browser %></td></tr>
            <tr><td><span class="property-name">Version</span></td><td><%= browser.version %></td></tr>
            <tr><td><span class="property-name">Frames</span></td><td><%= browser.frames %></td></tr>
            <tr><td><span class="property-name">Tables</span></td><td><%= browser.tables %></td></tr>
            <tr><td><span class="property-name">Cookies</span></td><td><%= browser.cookies %></td></tr>
            <tr><td><span class="property-name">VBScript</span></td><td><%= browser.vbscript %></td></tr>
            <tr><td><span class="property-name">JavaScript</span></td><td><%= browser.javascript %></td></tr>
        </table>
    </div>

    <!-- 4. MSWC.MyInfo -->
    <div class="section">
        <div class="section-title">4. MSWC.MyInfo</div>
        <p>Reads site administrator information from <code>www/MyInfo.xml</code>.</p>
        <%
        Set myinfo = Server.CreateObject("MSWC.MyInfo")
        %>
        <code>
            Set myinfo = Server.CreateObject("MSWC.MyInfo")<br>
            Response.Write myinfo.PersonalName
        </code>
        <table>
            <tr><th width="200">Property</th><th>Value</th></tr>
            <tr><td><span class="property-name">PersonalName</span></td><td><%= myinfo.PersonalName %></td></tr>
            <tr><td><span class="property-name">PersonalMail</span></td><td><%= myinfo.PersonalMail %></td></tr>
            <tr><td><span class="property-name">CompanyName</span></td><td><%= myinfo.CompanyName %></td></tr>
            <tr><td><span class="property-name">PersonalWords</span></td><td><%= myinfo.PersonalWords %></td></tr>
            <tr><td><span class="property-name">URL(1)</span></td><td><a href="<%= myinfo.URL(1) %>"><%= myinfo.URLWords(1) %></a></td></tr>
            <tr><td><span class="property-name">URL(2)</span></td><td><a href="<%= myinfo.URL(2) %>"><%= myinfo.URLWords(2) %></a></td></tr>
        </table>
    </div>

    <!-- 5. MSWC.NextLink -->
    <div class="section">
        <div class="section-title">5. MSWC.NextLink</div>
        <p>The NextLink component creates a sequential navigation based on a list file (<code>links_demo.txt</code>).</p>
        <%
        Set nextlink = Server.CreateObject("MSWC.NextLink")
        listFile = "links_demo.txt"
        currentIdx = nextlink.GetListIndex(listFile)
        count = nextlink.GetListCount(listFile)
        %>
        <code>
            Set nextlink = Server.CreateObject("MSWC.NextLink")<br>
            nextUrl = nextlink.GetNextURL("links_demo.txt")
        </code>
        <table>
            <tr><th width="200">Method</th><th>Value</th></tr>
            <tr><td><span class="property-name">GetListCount</span></td><td><%= count %> items</td></tr>
            <tr><td><span class="property-name">GetListIndex</span></td><td>Item #<%= currentIdx %></td></tr>
            <tr><td><span class="property-name">GetNextURL</span></td><td><a href="<%= nextlink.GetNextURL(listFile) %>"><%= nextlink.GetNextURL(listFile) %></a></td></tr>
            <tr><td><span class="property-name">GetNextDescription</span></td><td><%= nextlink.GetNextDescription(listFile) %></td></tr>
            <tr><td><span class="property-name">GetPreviousURL</span></td><td><a href="<%= nextlink.GetPreviousURL(listFile) %>"><%= nextlink.GetPreviousURL(listFile) %></a></td></tr>
            <tr><td><span class="property-name">GetNthURL(2)</span></td><td><%= nextlink.GetNthURL(listFile, 2) %></td></tr>
        </table>
    </div>

    <!-- 6. MSWC.ContentRotator -->
    <div class="section">
        <div class="section-title">6. MSWC.ContentRotator</div>
        <p>Rotates pieces of HTML content from a text file (<code>content_demo.txt</code>).</p>
        <%
        Set rotator = Server.CreateObject("MSWC.ContentRotator")
        %>
        <code>
            Set rotator = Server.CreateObject("MSWC.ContentRotator")<br>
            Response.Write rotator.ChooseContent("content_demo.txt")
        </code>
        <div style="margin: 10px 0;">
            <%= rotator.ChooseContent("content_demo.txt") %>
        </div>
        <p><em>(Refresh the page to see another content piece rotated)</em></p>
    </div>

    <!-- 7. MSWC.Tools -->
    <div class="section">
        <div class="section-title">7. MSWC.Tools</div>
        <p>Utility methods for common server-side tasks.</p>
        <%
        Set tools = Server.CreateObject("MSWC.Tools")
        %>
        <code>
            Set tools = Server.CreateObject("MSWC.Tools")<br>
            Response.Write tools.FileExists("default.asp")
        </code>
        <%
        currentPath = Server.MapPath("test_mswc_demo.asp")
        ownerVal = tools.Owner(currentPath)
        %>
        <table>
            <tr><th width="200">Method</th><th>Value</th></tr>
            <tr><td><span class="property-name">FileExists("default.asp")</span></td><td><%= tools.FileExists("default.asp") %></td></tr>
            <tr><td><span class="property-name">Owner("<%= currentPath %>")</span></td><td><%= ownerVal %></td></tr>
            <tr><td><span class="property-name">PluginExists("asp")</span></td><td><%= tools.PluginExists("asp") %></td></tr>
        </table>
    </div>

    <div style="text-align: center; color: #808080; font-size: 10px; margin-top: 20px;">
        &copy; 2026 G3Pix AxonASP Server - Classic Component Implementation.
    </div>

</div>

</body>
</html>
