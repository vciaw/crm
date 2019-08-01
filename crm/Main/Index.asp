<!--#include file="../data/conn.asp" --><%
Dim url
If Session("CRM_account") <> "" And Session("CRM_name") <> "" And IsNumeric(Session("CRM_level")) Then 
url = "main.asp"  
else
Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('../index.asp');</script>"
end if
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><%=title%></title>
<link rel="shortcut icon" href="<%=SiteUrl&skinurl%>images/Ecrm.ico" type="image/x-icon" /> 
<style type="text/css">
html{ overflow:hidden;}
table,td,th, body,dt,dd,dl{ margin:0; padding:0; border:none;}
</style>
</head>
<body scroll="no">
<table cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
        <td colspan="2" height="33"><iframe src="header.asp" name="header" target="menu" width="100%" height="33" scrolling="no" frameborder="0"></iframe></td>
    </tr>
    <tr>
        <td valign="top" rowspan="2" width="110"><iframe src="menu.asp" name="menu" target="main" width="110" height="100%" scrolling="no" frameborder="0"></iframe></td>
        <td valign="top" height="100%">
        	<table  cellpadding="0" cellspacing="0" width="100%" height="100%">
                <tr>
                    <td valign="top" width="100%"><iframe src="<% = url %>" id="main" name="main" width="100%" height="100%" frameborder="0" scrolling="yes" style="overflow:visible;"></iframe></td>
                </tr>
            </table>
        
        </td>
    </tr>
</table>

</body>
</html>