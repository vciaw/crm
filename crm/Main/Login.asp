<!--#include file="../data/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<title><%=title%></title>
<link rel="shortcut icon" href="<%=SiteUrl&skinurl%>images/Ecrm.ico" type="image/x-icon" /> 
 <link rel="stylesheet" href="<%=SiteUrl%>assets/css/css.css">
<link rel="stylesheet" href="<%=SiteUrl%>assets/css/reset.css">
<link rel="stylesheet" href="<%=SiteUrl%>assets/css/supersized.css">
<link rel="stylesheet" href="<%=SiteUrl%>assets/css/style.css">
<!-- HTML5 shim, for IE6-8 support of HTML5 elements -->
        <!--[if lt IE 9]>
            <script>
(function(l,f){function m(){var a=e.elements;return"string"==typeof a?a.split(" "):a}function i(a){var b=n[a[o]];b||(b={},h++,a[o]=h,n[h]=b);return b}function p(a,b,c){b||(b=f);if(g)return b.createElement(a);c||(c=i(b));b=c.cache[a]?c.cache[a].cloneNode():r.test(a)?(c.cache[a]=c.createElem(a)).cloneNode():c.createElem(a);return b.canHaveChildren&&!s.test(a)?c.frag.appendChild(b):b}function t(a,b){if(!b.cache)b.cache={},b.createElem=a.createElement,b.createFrag=a.createDocumentFragment,b.frag=b.createFrag();
a.createElement=function(c){return!e.shivMethods?b.createElem(c):p(c,a,b)};a.createDocumentFragment=Function("h,f","return function(){var n=f.cloneNode(),c=n.createElement;h.shivMethods&&("+m().join().replace(/[\w\-]+/g,function(a){b.createElem(a);b.frag.createElement(a);return'c("'+a+'")'})+");return n}")(e,b.frag)}function q(a){a||(a=f);var b=i(a);if(e.shivCSS&&!j&&!b.hasCSS){var c,d=a;c=d.createElement("p");d=d.getElementsByTagName("head")[0]||d.documentElement;c.innerHTML="x<style>article,aside,dialog,figcaption,figure,footer,header,hgroup,main,nav,section{display:block}mark{background:#FF0;color:#000}template{display:none}</style>";
c=d.insertBefore(c.lastChild,d.firstChild);b.hasCSS=!!c}g||t(a,b);return a}var k=l.html5||{},s=/^<|^(?:button|map|select|textarea|object|iframe|option|optgroup)$/i,r=/^(?:a|b|code|div|fieldset|h1|h2|h3|h4|h5|h6|i|label|li|ol|p|q|span|strong|style|table|tbody|td|th|tr|ul)$/i,j,o="_html5shiv",h=0,n={},g;(function(){try{var a=f.createElement("a");a.innerHTML="<xyz></xyz>";j="hidden"in a;var b;if(!(b=1==a.childNodes.length)){f.createElement("a");var c=f.createDocumentFragment();b="undefined"==typeof c.cloneNode||
"undefined"==typeof c.createDocumentFragment||"undefined"==typeof c.createElement}g=b}catch(d){g=j=!0}})();var e={elements:k.elements||"abbr article aside audio bdi canvas data datalist details dialog figcaption figure footer header hgroup main mark meter nav output progress section summary template time video",version:"3.7.0",shivCSS:!1!==k.shivCSS,supportsUnknownElements:g,shivMethods:!1!==k.shivMethods,type:"default",shivDocument:q,createElement:p,createDocumentFragment:function(a,b){a||(a=f);
if(g)return a.createDocumentFragment();for(var b=b||i(a),c=b.frag.cloneNode(),d=0,e=m(),h=e.length;d<h;d++)c.createElement(e[d]);return c}};l.html5=e;q(f)})(this,document);
</script>
        <![endif]-->
<script language="javascript">
function check()
{
	var obj = document.loginForm;
	if (obj.item1.value == '')
	{
		obj.item1.focus();
		return false;
	}
	if (obj.item2.value == '')
	{
		obj.item2.focus();
		return false;
	}
	return true;
}
</script>
</head>
<body>
<%
action = Trim(Request("action"))
Fromurl = Trim(Request("Fromurl"))
Select Case action
Case "login"
    Call login()
Case Else
    Call loginForm()
End Select

Sub login()
    Dim account,password
	account = Trim(Request("item1"))
	password = Lcase(Request("item2"))
	password = md5(password,16)
	If account = "" Or password = "" Then
	    Response.Redirect("login.asp?errMsg=2")
		Response.End()
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [user] Where uAccount = '" & account & "'",conn,3,1
	If rs.RecordCount <> 1 Then
	    Response.Redirect("login.asp?errMsg=1")
		Response.End()
	End If
	If password <> rs("uPassword") Then
	    Response.Redirect("login.asp?errMsg=1")
		Response.End()
	End If
	qxflag = rs("uqxflag")
	if mid(qxflag, 1, 1) <> "1" then
	    Response.Redirect("login.asp?errMsg=3")
		Response.End()
	End If
	
	'写入Session
	Session("CRM_account") = account
	Session("CRM_name") = rs("uName")
	Session("CRM_uId") = rs("uId")
	Session("CRM_level") = rs("uLevel")
	Session("CRM_group") = rs("uGroup")
	Session("CRM_qx") = rs("uqxflag")
	Session("CRM_MR") = rs("uManagerange")
	Session("CRM_Accsql") = Accsql
	Session("Data_Source") = Data_Source
	Session("Data_User") = Data_User
	Session("Data_Password") = Data_Password
	Session("Data_Catalog") = Data_Catalog
	Session("Data_MDBPath") = Data_MDBPath
	Session("CRM_url") = Fromurl
	
	if Keeponline =1 then
	'写入Cookies
	Response.Cookies(CookieKey)("CRM_account") = account
	Response.Cookies(CookieKey)("CRM_name") = ""&rs("uName")&""
	Response.Cookies(CookieKey)("CRM_uId") = ""&rs("uId")&""
	Response.Cookies(CookieKey)("CRM_level") = ""&rs("uLevel")&""
	Response.Cookies(CookieKey)("CRM_group") = ""&rs("uGroup")&""
	Response.Cookies(CookieKey)("CRM_qx") = ""&rs("uqxflag")&""
	Response.Cookies(CookieKey)("CRM_MR") = ""&rs("uManagerange")&""
	Response.Cookies(CookieKey)("CRM_Accsql") = Accsql
	Response.Cookies(CookieKey)("Data_Source") = Data_Source
	Response.Cookies(CookieKey)("Data_User") = Data_User
	Response.Cookies(CookieKey)("Data_Password") = Data_Password
	Response.Cookies(CookieKey)("Data_Catalog") = Data_Catalog
	Response.Cookies(CookieKey)("Data_MDBPath") = Data_MDBPath
	Response.Cookies(CookieKey)("CRM_url") = Fromurl
	end if

	rs.Close
	
	if YnUserLog=1 then
	    sqlgh="UPDATE Client SET cYn = 0 WHERE  cuser  in (select uname from [user] where uLevel<>9 ) and cType<>'"&CRTypeEnd&"' and (DATEDIFF(d, cLastUpdated, { fn NOW() }) >" &gdzy&")"
        conn.execute(sqlgh)
		sqlgs="UPDATE Client SET cYn = 1 WHERE cType='"&CRTypeEnd&"'"
		conn.execute(sqlgs)
		conn.execute ("insert into userlog(olname,olstarttime,olip) values('"&Session("CRM_name")&"','"&now()&"','"&Request.ServerVariables("REMOTE_ADDR") &"')")
	end if
	
	Set rs = Nothing
	if Fromurl <>"" then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('"&Fromurl&"');</script>"
	else
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
	end if
End Sub

Sub loginForm()
    Dim errMsg
	errMsg = CInt(ABS(Request("errMsg")))
	Select Case errMsg
	Case 2
	    errMsg = "<script>alert("""&alertlogin01&""");history.back(1);</script>"
	Case 1
	    errMsg = "<script>alert("""&alertlogin02&""");history.back(1);</script>"
	Case 3
	    errMsg = "<script>alert("""&alertlogin03&""");history.back(1);</script>"
	Case Else
	    errMsg = ""
	End Select
%>
	
    
      <div class="page-container">
            <h1>CRM客户管理系统</h1>
          <form name="loginForm" method="post" action="?action=login" onsubmit="return check();">
                	<INPUT  id="item1" type="text" name="item1" class="username"  value="" placeholder="用户名">
                
               
               
               <INPUT  id="item2"  type="password" name="item2" class="password"   value="" placeholder="密码">
               
               
                <button type="submit"  id="submitid" >登录</button>
                	<% = errMsg %>
            </form>
            
        </div>
    
        <% End Sub %>
        <!-- Javascript -->
        <script src="<%=SiteUrl%>assets/js/jquery-1.8.2.min.js"></script>
        <script src="<%=SiteUrl%>assets/js/supersized.3.2.7.min.js"></script>
        <script>
		 jQuery(function($){

    $.supersized({

        // Functionality
        slide_interval     : 4000,    // Length between transitions
        transition         : 1,    // 0-None, 1-Fade, 2-Slide Top, 3-Slide Right, 4-Slide Bottom, 5-Slide Left, 6-Carousel Right, 7-Carousel Left
        transition_speed   : 1000,    // Speed of transition
        performance        : 1,    // 0-Normal, 1-Hybrid speed/quality, 2-Optimizes image quality, 3-Optimizes transition speed // (Only works for Firefox/IE, not Webkit)

        // Size & Position
        min_width          : 0,    // Min width allowed (in pixels)
        min_height         : 0,    // Min height allowed (in pixels)
        vertical_center    : 1,    // Vertically center background
        horizontal_center  : 1,    // Horizontally center background
        fit_always         : 0,    // Image will never exceed browser width or height (Ignores min. dimensions)
        fit_portrait       : 1,    // Portrait images will not exceed browser height
        fit_landscape      : 0,    // Landscape images will not exceed browser width

        // Components
        slide_links        : 'blank',    // Individual links for each slide (Options: false, 'num', 'name', 'blank')
        slides             : [    // Slideshow Images
                                 {image : '<%=SiteUrl%>assets/img/backgrounds/1.jpg'},
                                 {image : '<%=SiteUrl%>assets/img/backgrounds/2.jpg'},
                                 {image : '<%=SiteUrl%>assets/img/backgrounds/3.jpg'}
                             ]

    });

});
		</script>
		
		
		
		
		
		
		
		
        <script src="<%=SiteUrl%>assets/js/scripts.js"></script>
    

</body>
</html>