<%
'########################################################################################
'###
'### QuickerSite REBRANDING options: rename QuickerSite for your customers!
'###
'########################################################################################
'###
'### The parameters below will help your to rebrand QuickerSite for your customers!
'### The labels from the QuickerLabels.mdb will also be affected by the setting below. 
'### Warning: make sure to UNLOAD your IIS application after you've updated this file!
'### You can easily code this in ASP: Application.Contents.RemoveAll()
'###
'########################################################################################

'in case you need a different name for your CMS, set it here:
const MYQS_name="QuickerSite"
const MYQS_slogan="CMS the fun way!"

'You might also want to change the URL on top of each page in the backsite:
const MYQS_url="http://www.quickersite.com/"

'the META-TAG Generator that is auto-generated in most pages:
const MYQS_GENERATOR="QuickerSite CMS - visit www.quickersite.com"

'Url to your support pages for your CMS (when clicking the Help-image in the backsite)
const MYQS_urlSupport="http://www.quickersite.com/r/?sCode=SUPPORT"

'Url to your support page regarding Templates!
const MYQS_urlTemplates="http://www.quickersite.com/r/?sCode=TEMPLATES"

'Url to your support page regarding User Friendly URLS
const MYQS_urlUFLs="http://www.quickersite.com/r/ufl"

'Footer
const MYQS_FOOTER="<p>Copyright &copy; QuickerSite 2023. All Rights Reserved.<br />Visit <a style=""color:#EEE"" target=""_blank"" href=""http://www.quickersite.com"">www.quickersite.com</a></p>"

'Icon Colors
const MYQS_IconColor="#3A383E"
const MYQS_IconHoverColor="#70919A"


function artHeaderPart
'copy/paste the header from your template here
'DO NOT INCLUDE <html>,<head> or <body>-tags.
%>
<style>
.art-button {transition: 0.3s;}
</style>
    <link rel="stylesheet" href="<%=C_DIRECTORY_QUICKERSITE%>/backsiteTemplate31/style.css" type="text/css" media="screen" />
    <!--[if IE 6]><link rel="stylesheet" href="<%=C_DIRECTORY_QUICKERSITE%>/backsiteTemplate31/style.ie6.css" type="text/css" media="screen" /><![endif]-->
    <!--[if IE 7]><link rel="stylesheet" href="<%=C_DIRECTORY_QUICKERSITE%>/backsiteTemplate31/style.ie7.css" type="text/css" media="screen" /><![endif]-->

    <script type="text/javascript" src="<%=C_DIRECTORY_QUICKERSITE%>/backsiteTemplate31/jquery.js"></script>
    <script type="text/javascript" src="<%=C_DIRECTORY_QUICKERSITE%>/backsiteTemplate31/script.js"></script>
	
<%
end function

function artTopBannerAndMenu
'start from the <body>-tag, end with the post-content-part (comment out the art-postheader-part)
%>
<body>
    <div id="art-main">
    <div class="cleared reset-box"></div>
    <div class="art-header">
        <div class="art-header-position">
            <div class="art-header-wrapper">
                <div class="cleared reset-box"></div>
                <div class="art-header-inner">
                <div class="art-headerobject"></div>
                <div class="art-logo">
                                 <h1 class="art-logo-name"><a href="#"><%=MYQS_name%></a></h1>
                                                 <h2 class="art-logo-text"><%=MYQS_slogan%></h2>
                                </div>
                </div>
            </div>
        </div>
        
    </div>
    <div class="cleared reset-box"></div>
<div class="art-bar art-nav">
<div class="art-nav-outer">
<div class="art-nav-wrapper">
<div class="art-nav-inner">
	<div class="art-nav-center">
	<ul class="art-hmenu">
       
        		<%=getBSArtMenu%>
        	</ul>
        	</div>
</div>
</div>
</div>
</div>
<div class="cleared reset-box"></div>
<div class="art-box art-sheet">
        <div class="art-box-body art-sheet-body">
            <div class="art-layout-wrapper">
                <div class="art-content-layout">
                    <div class="art-content-layout-row">
                        <div class="art-layout-cell art-content">
<div class="art-box art-post">
    <div class="art-box-body art-post-body">
<div class="art-post-inner art-article">
                                <!--<h2 class="art-postheader">Welcome
                                </h2>-->
                                                <div class="art-postcontent">

<%
end function

function artFooterPart
'copy the end of the postcontent untill right before </body>. Do NOT include the </body> closing tag!
%>
     
                                         
                                                    
                                                     </div>
                <div class="cleared"></div>
</div>

		<div class="cleared"></div>
    </div>
</div>

                          <div class="cleared"></div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="cleared"></div>
            <div class="art-footer">
                <div class="art-footer-body">
                    <!--<a href="#" class="art-rss-tag-icon" title="RSS"></a>-->
                            <div class="art-footer-text">
                                <%=MYQS_FOOTER%>
                                                            </div>
                    <div class="cleared"></div>
                </div>
            </div>
    		<div class="cleared"></div>
        </div>
    </div>
  
   
    <div class="cleared"></div>
</div>
       
                                    

<%
end function

function artTopBannerAndMenuLOGIN
%>

<body>
   <div id="art-main">
    <div class="cleared reset-box"></div>
    <div class="art-header">
        <div class="art-header-position">
            <div class="art-header-wrapper">
                <div class="cleared reset-box"></div>
                <div class="art-header-inner">
                <div class="art-headerobject"></div>
                <div class="art-logo">
                                 <h1 class="art-logo-name"><a href="#"><%=MYQS_name%></a></h1>
                                                 <h2 class="art-logo-text"><%=MYQS_slogan%></h2>
                                </div>
                </div>
            </div>
        </div>
        
    </div>
    <div class="cleared reset-box"></div>
    <div class="art-box art-sheet">
        <div class="art-box-body art-sheet-body">
            <div class="art-layout-wrapper">
                <div class="art-content-layout">
                    <div class="art-content-layout-row">
                        <div class="art-layout-cell art-content">
<div class="art-box art-post">
    <div class="art-box-body art-post-body">
<div class="art-post-inner art-article">
                                <!--<h2 class="art-postheader">Welcome
                                </h2>-->
    <div class="art-postcontent">
<%
end function

function artFooterPartLOGIN
%>
         
           </div>
                <div class="cleared"></div>
</div>

		<div class="cleared"></div>
    </div>
</div>

                          <div class="cleared"></div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="cleared"></div>
            <div class="art-footer">
                <div class="art-footer-body">
                    <!--<a href="#" class="art-rss-tag-icon" title="RSS"></a>-->
                            <div class="art-footer-text">
                                <%=MYQS_FOOTER%>
                                                            </div>
                    <div class="cleared"></div>
                </div>
            </div>
    		<div class="cleared"></div>
        </div>
    </div>
    <div class="cleared"></div>
    <div class="cleared"></div>
</div>                                 
        
        
        
       
<%
end function

function artHeaderPartLOGIN
%>
    <style>
.art-button {transition: 0.3s;}
</style>
     <link rel="stylesheet" href="<%=C_DIRECTORY_QUICKERSITE%>/backsiteTemplateLOGIN31/style.css" type="text/css" media="screen" />
    <!--[if IE 6]><link rel="stylesheet" href="<%=C_DIRECTORY_QUICKERSITE%>/backsiteTemplateLOGIN31/style.ie6.css" type="text/css" media="screen" /><![endif]-->
    <!--[if IE 7]><link rel="stylesheet" href="<%=C_DIRECTORY_QUICKERSITE%>/backsiteTemplateLOGIN31/style.ie7.css" type="text/css" media="screen" /><![endif]-->

    <script type="text/javascript" src="<%=C_DIRECTORY_QUICKERSITE%>/backsiteTemplateLOGIN31/jquery.js"></script>
    <script type="text/javascript" src="<%=C_DIRECTORY_QUICKERSITE%>/backsiteTemplateLOGIN31/script.js"></script>
    

<%
end function
%>
