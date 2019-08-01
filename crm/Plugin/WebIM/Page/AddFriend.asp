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
function chkEmail()
{
	if($("tb1").style.display=="none")return true;
	var email = $F("tbEmail").trim();
	if(email=="")
	{
		setTip("Email","请填写email地址","red");
		return false;
	}
	else if(!validEmail(email))
	{
		setTip("Email","错误的email地址","red");
		return false;
	}
	else if(!exsitEmail(email))
	{
		setTip("Email","不存在这样的用户","red");	
		return false;
	}
	setTip("Email","OK","gray");
	return true;
}
function chkAll()
{
	if(chkEmail())
	{
		showLoading();
		document.forms[0].submit();
	}
}
function setTip(s,msg,cn)
{
	var oSpan = $("span"+s);
	oSpan.className = cn;
	Elem.Value(oSpan,msg);
}
function validEmail(email)
{
    var regex = /^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/;
    return regex.test(email);
}
function exsitEmail(email)
{
	var ajax = new Ajax();
	ajax.send("../data/service.asp?t=1","email="+email,null,"POST",false);
	return parseInt(Xml.First($T(ajax.req.responseXML,"result").item(0),"num"))!=0;
}
</script>
</head>
<body> 
<div style="float:left;width:348px">
<%
If CInt(Session("userpower"))>2 Then 
	Response.Write("匿名用户无此功能！")
	Response.End
End If
t = CInt(Request.QueryString("t"))
email = Request.QueryString("tbEmail")
gender = Request.QueryString("gender")
face = Request.QueryString("face")
status = Request.QueryString("status")
p = Request.QueryString("p")
If t>0 Then
	If t = 1 Then
		sql = "select * from [user] where useremail = '"&Request.QueryString("tbEmail")&"'"
	ElseIf t = 2 Then
		sql = "select * from [user] where 1 "
		If gender <> "-1" Then sql = sql&" and usergender = "&gender
		If face <> "-1" Then 
			If face = "1" Then
				sql = sql&" and userface <> 'default.gif'"
			Else
				sql = sql&" and userface = 'default.gif'"
			End If
		End If
		If status <> "-1" Then 
			If status = "1" Then
				sql = sql&" and userstatus <6"
			Else
				sql = sql&" and userstatus >5"
			End If
		End If
	End If
%>
<table align="center" width="98%" border="0" cellpadding="0" cellspacing="1" style="background-color:#bed6e0">
<%
Call DataBegin()
oRs.Open sql,oConn,1,1
If Not (oRs.Bof And oRs.Eof) Then
	oRs.PageSize = 10	
	If p="" Or IsNumeric(p)=False Then p = 1
	p = CInt(p)
	oRs.AbsolutePage = p
	For I = 1 To oRs.PageSize
	If I Mod 2 = 1 Then
		TdColor = "fff"
	Else
		TdColor = "e0edff"
	End If
%>
	<tr>
		<td style="height:22px;background-color:#<%=TdColor%>;width:25px;text-align:center"><script>document.write("<img title='"+["联机","忙碌","马上回来","离开","通话中","外出就餐","脱机","脱机"][<%=oRs("userstatus")%>]+"' src='../images/m"+[0,1,2,2,1,2,3,3][<%=oRs("userstatus")%>]+".gif' />")</script></td>
		<td style="background-color:#<%=TdColor%>;text-indent:3px"><a title="<%=oRs("usersign")%>" href="javascript:void(0)"><%=oRs("username")%></a>[<%If oRs("usergender") = "1" Then Response.Write("男") Else Response.Write("女") End If%>]</td>
		<td style="background-color:#<%=TdColor%>;width:80px;text-align:center"><a title="<%=oRs("useremail")%>" href="addfriendsave.asp?email=<%=oRs("useremail")%>">加为好友</a></td>
	</tr>
<%
	oRs.MoveNext
	If oRs.Eof Then Exit For
	Next
Else
	Response.Write("<tr><td  style=""height:25px;background-color:#fff"" colspan=""6"">&nbsp;没有任何记录！</td></tr>")
End If
%>
</table>
</div>
<div style="float:left;padding:7px 0 0 5px">
	<a href="addfriend.asp">←重新查找</a>
</div>
<%If oRs.PageCount > 0 Then%>
<div style="float:right;padding:5px 5px 0 0">
	到
	<select style="font-size:11px" onchange="showLoading();location.href='?face=<%=face%>&email=<%=email%>&status=<%=status%>&gender=<%=gender%>&t=<%=t%>&p='+this.value" >
	<%
		For Q = 1 To oRs.PageCount
			Response.Write("<option value="""&Q&"""")
			If p=Q Then Response.Write(" selected=""selected"" ")
			Response.Write(">"&Q&"</option>")
		Next
	%>
	</select>
	页&nbsp;
	共<%=p%>/<%=oRs.PageCount%>页
	<%If p<>1 Then%>	
		<a onclick="showLoading();" href="?face=<%=face%>&email=<%=email%>&status=<%=status%>&gender=<%=gender%>&t=<%=t%>&p=1">首页</a>
		<a onclick="showLoading();" href="?face=<%=face%>&email=<%=email%>&status=<%=status%>&gender=<%=gender%>&t=<%=t%>&p=<%=p-1%>">上页</a>
	<%Else%>
		<span class="gray">首页</span>
		<span class="gray">上页</span>
	<%End If%>
	<%If p<>oRs.PageCount Then%>	
		<a onclick="showLoading();" href="?face=<%=face%>&email=<%=email%>&status=<%=status%>&gender=<%=gender%>&t=<%=t%>&p=<%=p+1%>">下页</a>
		<a onclick="showLoading();" href="?face=<%=face%>&email=<%=email%>&status=<%=status%>&gender=<%=gender%>&t=<%=t%>&p=<%=oRs.PageCount%>">末页</a>
	<%Else%>
		<span class="gray">下页</span>
		<span class="gray">末页</span>
	<%End If%>
</div>
<%
End If
oRs.Close()
Set oRs = Nothing
Call DataEnd()
%>
<%
Else
%>
<form action="addfriend.asp" method="get" name="form1" id="form1"> 
<div style="height:40px;text-indent:10px;line-height:40px"><strong>查找方式</strong></div>
<div style="height:30px;text-indent:10px"><label for="t1"><input type="radio" id="t1" checked="checked" value="1" name="t"/>精确查找<label></div>
<div style="height:30px;text-indent:10px"><label for="t2"><input type="radio" id="t2" value="2" name="t"/>按条件查找<label></div>
<div style="height:85px;padding:15px" id="tb1">
	<div style="height:30px">
		Email地址：<input name="tbEmail" type="text" class="input1" id="tbEmail" maxlength="50" onblur="chkEmail()"/>
	</div>
	<div style="height:30px">
		<span id="spanEmail">示例：quguangyu@gmail.com</span></td>
	</div>
</div>
<div style="height:85px;display:none;padding:15px" id="tb2">
	<div style="height:30px">
		性别：<select name="gender"><option value="-1">不限</option><option value="1">男</option><option value="2">女</option></select>&nbsp;&nbsp;自定义头像：<select name="face"><option value="-1">不限</option><option value="1">是</option><option value="2">不是</option></select>
	</div>
	<div style="height:30px">
		在线状态：<select name="status"><option value="-1">不限</option><option value="1">在线</option><option value="2">不在线</option></select>
	</div>
  </tr>
</div>
<div style="height:35px;text-align:center">
        <input class="button1" type="button" name="btnSubmit" id="btnSubmit" value="确定" onclick="chkAll()"/>&nbsp;&nbsp; 
        <input class="button1" type="button" name="btnCancel" id="btnCancel" onclick="winClose(event);" value="取消" /></td>
</div>   
</form>

<script type="text/javascript">
$("t1").onclick = $("t2").onclick = function(){
	Elem.Hid("tb1","tb2");
	Elem.Show("tb"+this.value);
};
</script>
<%End If%>
</div>
</body>
</html>