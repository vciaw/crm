<!--#include file="../data/Mssql.asp"-->
<!--#include file="../data/Config.asp"-->
<!--#include file="../data/Md5.asp" -->
<!--#include file="../data/Function.asp"--><%
	If Session("CRM_account") <> "" And Session("CRM_name") <> "" And IsNumeric(Session("CRM_level")) Then 
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
	end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "/www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="/www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><%=title%></title>
<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=0" />
<link rel="apple-touch-icon" href="img/57.png" />
<link rel="apple-touch-icon" sizes="72x72" href="img/72.png" />
<link rel="apple-touch-icon" sizes="114x114" href="img/114.png" />
<meta name="keywords" content="" />
<meta name="description" content="" />
<link rel="stylesheet" type="text/css" href="style/reset.css" /> 
<link rel="stylesheet" type="text/css" href="style/root.css" /> 
<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript" src="js/toogle.js"></script>
</head>
<body>
<!-- start header -->
    <div id="header">
         <a href="#"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
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
	    Response.Redirect("Login.asp?errMsg=2")
		Response.End()
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [user] Where uAccount = '" & account & "'",conn,3,1
	If rs.RecordCount <> 1 Then
	    Response.Redirect("Login.asp?errMsg=1")
		Response.End()
	End If
	If password <> rs("uPassword") Then
	    Response.Redirect("Login.asp?errMsg=1")
		Response.End()
	End If
	qxflag = rs("uqxflag")
	if mid(qxflag, 1, 1) <> "1" then
	    Response.Redirect("Login.asp?errMsg=3")
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
		conn.execute ("insert into userlog(olname,olstarttime,olip) values('"&Session("CRM_name")&"','"&now()&"','EasyCrm Mobile')")
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
    <!-- start page -->
    <div class="page">
    
    		
            <div class="simplebox">
            	<h1 class="titleh">系统登录</h1>
                <div class="content">
                	
                  <form name="loginForm" method="post" action="?action=login" onsubmit="return check();">
                    <div class="form-line">
                   	  <label class="st-label">帐号</label>
                      <input type="text" name="item1" id="item1"  style=" width:80%;" value="" />
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label">密码</label>
                      <input type="text" name="item2" id="item2"  style=" width:80%;" value=""  />
                    </div>
                    
                    <div class="form-line">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input type="reset" name="button" id="button2" value="&nbsp;重 置&nbsp;" class="reset-button" />
                    </div>

                  </form>
                </div>
			</div>
            
			<%=Footer%>
            
    <div class="clear"></div>
    </div>
    <!-- end page -->
</body>
</html>
<script type="text/javascript" src="js/frame.js"></script>
<% End Sub %>