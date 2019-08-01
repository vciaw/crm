<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "../data/config.asp"-->
<!--#include file = "../data/function.asp"-->
<!--#include file = "../data/cmd.asp"-->
<%
Response.Expires = WebCachTime
Response.Charset="utf-8"
Call CheckLogin()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../styles/webimpage.css" type="text/css" rel="stylesheet" media="all">
<script type="text/javascript" src="../js/webimhelper.js"></script>
<script type="text/javascript" src="../js/webimpage.js"></script>
<title>添加联系人</title>
<script type="text/javascript">
var uid = 1;
</script>
</head>
<body> 
<%
If CInt(Session("userpower"))>2 Then 
	Response.Write("匿名用户无此功能！")
	Response.End
End If
%>
<%
If Request.QueryString("email")<>"" Then
	email = GetSafeStr(Request.QueryString("email"))
	If email = Session("useremail") Then
		strResult = "<span class='red'>您不能加自己为好友!</span>"
	Else
		Call DataBegin()
		toid = GetUserIdByEmail(email)
		If oConn.ExeCute("select count(*) from [userfriend] where userid="&Session("userid")&" and friendid="&toid)(0)>0 Then
			strResult = "<span class='red'>您已经添加过这位好友!</span>"
		ElseIf  oConn.ExeCute("select count(*) from usersysmsg where fromid="&Session("userid")&" and toid="&toid)(0)>0 Then
			strResult = "<span class='red'>请耐心等待好友回复，不要重复发送!</span>"
		Else
			oConn.Execute("insert into usersysmsg (fromid,toid,msgcontent,typeid,msgaddtime) values ('"&Session("userid")&"','"&toid&"','"&Session("useremail")&"','7','"&Now()&"')")
			strResult = "发送成功，请等待验证结果!"
		End If
		Call DataEnd()
	End If
%>
<table width="200" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="180" align="center"><%=strResult%><br /><br /><br />
	<a href="javascript:location.href='<%=Request.ServerVariables("Http_REFERER")%>'">←返回上页</a></td>
  </tr>
  <tr>
    <td height="85" align="center">
        <input class="button1" type="button" name="btnCancel" id="btnCancel" onclick="winClose(event);" value="关闭" /></td>
  </tr>
</table>
<%End If%>
</body>
</html>