<!--#include file="../data/Mssql.asp"-->
<!--#include file="../data/Config.asp"-->
<!--#include file="../data/Md5.asp" -->
<!--#include file="../data/Function.asp"-->
<!--#include file="../data/CheckInput.asp"-->
<!--#include file="../lang/zh-cn/lang.asp"-->
<!--#include file="../data/Config/Show_Listall.asp"-->
<!--#include file="../data/Config/Show_Client.asp"-->
<!--#include file="../data/Config/Must_Client.asp"-->
<!--#include file="../data/Config/Show_Must_Linkmans.asp"-->
<!--#include file="../data/Config/Show_Must_Records.asp"-->
<!--#include file="../data/Config/Show_Must_Order.asp"-->
<!--#include file="../data/Config/Show_Must_Order_Products.asp"-->
<!--#include file="../data/Config/Show_Must_Products.asp"-->
<!--#include file="../data/Config/Show_Must_Hetong.asp"-->
<!--#include file="../data/Config/Show_Must_Service.asp"-->
<!--#include file="../data/Config/Show_Must_Expense.asp"-->
<%
	If Session("CRM_account") <> "" And Session("CRM_name") <> "" And IsNumeric(Session("CRM_level")) Then 
	else
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('Login.asp');</script>"
	end if
	
Function Header()
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
<%
end Function
%>
<%
Function Footer()
%>
<%
end Function
%>
<%
Function Oter()
%>
<%
end Function

'·­Ò³Àà
Function pagelist(baseURL, PN, TotalPages, TotalRecords)
    If baseURL = "" Or IsNull(baseURL) Then Exit Function
    If InStr(baseURL, "?") And Right(baseURL, 1) <> "?" And Right(baseURL, 1) <> "&" Then
        baseURL = baseURL & "&"
    Else
        baseURL = baseURL & "?"
    End If
    Dim strList
        strList = "<div class=""pager"">"
        If PN < 2 Then
        strList = strList & "<a class='backpage grey'>Ê×Ò³</a>"
        Else
        strList = strList & "<a href='" & baseURL & "PN=1' class='backpage'>Ê×Ò³</a>"
        End If
    If PN > 2 Then s1 = PN - 2 Else s1 = 1
    If PN < TotalPages - 2 Then s2 = PN + 2 Else s2 = TotalPages
    For j = s1 To s2
       If j = PN Then
        strList = strList & "<a href='javascript:;' class='active'><b>" & j & "</b></a>"
       Else
        strList = strList & "<a href='" & baseURL & "PN=" & j & "'><b>" & j & "</b></a>"
       End If
    Next
        If PN = TotalPages Then
        strList = strList & "<a class='nextpage greys'>Î²Ò³</a>"
        Else
        strList = strList & "<a href='" & baseURL & "PN=" & TotalPages & "' class='nextpage'>Î²Ò³</a>"
        End If
        strList = strList & "</div>"
        pagelist = strList
End Function
%>