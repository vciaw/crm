<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
'获取当前页码
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 

If mid(Session("CRM_qx"), 5, 1) = 1 Then

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>

<body oncontextmenu=self.event.returnValue=false> 
<style>body{padding:35px 0 48px;}</style>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：系统管理 > 员工管理</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
			<div class="MenuboxS">
				 <ul>
					<li class="hover"><span><a href="?">员工列表</a></span></li>
					<li class=""><span><a href="#" onclick='User_Add()' style="cursor:pointer">新增员工</a></span></li>
				</ul>
			</div>
		</td>
	</tr>
</table>
<script>function User_Add() {$.dialog.open('GetUser.asp?action=User&sType=Add', {title: '新窗口', width: 800, height: 400,fixed: true}); };</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10">
<%
action = Trim(Request("action"))
Select Case action
Case "delete"
    Call deleteData()
Case Else
    Call Main()
End Select

Sub Main()
%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td width="90" class="td_l_c">编号</td>
					<td class="td_l_l">帐号</td>
					<td width="100" class="td_l_c">姓名</td>
					<td width="100" class="td_l_c">部门</td>
					<td width="100" class="td_l_c">等级</td>
					<td width="90" class="td_l_c">管理</td>
				</tr>
	<%
	PN = CLng(ABS(Request("PN")))
	If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
	intPageSize = DataPageSize
	pagenum = intPageSize*(PN-1)
	Set rs = Server.CreateObject("ADODB.Recordset")
	IF PN=1 THEN
	rs.Open "Select top "&intPageSize&" * From [user] Order By uId asc ",conn,1,1 
	ELSE
	rs.Open "Select top "&intPageSize&" * From [user] where uId > ( SELECT Max(uId) FROM ( SELECT TOP "&pagenum&" uId FROM [user] ORDER BY uId asc ) AS T ) Order By uId Asc ",conn,1,1
	END IF
	Set Rsstr=conn.Execute("Select count(uId) As RecordSum From [user] ",1,1)
	TotalRecords=Rsstr("RecordSum") 
	if Int(TotalRecords/intPageSize)=TotalRecords/intPageSize then
	TotalPages=TotalRecords/intPageSize
	else
	TotalPages=Int(TotalRecords/intPageSize)+1
	end if
	Rsstr.Close:Set Rsstr=Nothing
	If PN > TotalPages Then PN = TotalPages
	Do While Not rs.BOF And Not rs.EOF
	%>
							<tr class="tr">
								<td class="td_l_c"><%=rs("uId")%></td>
								<td class="td_l_l"><a href='#' onclick='User_InfoEdit<%=rs("uId")%>()' style="cursor:pointer"><%=rs("uAccount")%></a></td>
								<td class="td_l_c"><%=rs("uName")%></td>
								<td class="td_l_c"><%=EasyCrm.getNewItem("system_group","gId",rs("uGroup"),"gName")%></td>
								<td class="td_l_c"><%=EasyCrm.getNewItem("system_level","lId",rs("uLevel"),"lName")%></td>
								<td class="td_l_c"><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='User_InfoEdit<%=rs("uId")%>()' style="cursor:pointer" /> <input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='User_InfoDel<%=rs("uId")%>()' style="cursor:pointer" /></td>
							</tr>
							<script>function User_InfoEdit<%=rs("uId")%>() {$.dialog.open('GetUser.asp?action=User&sType=Edit&Id=<%=rs("uId")%>', {title: '新窗口', width: 800,height: 400, fixed: true}); };</script>
							<script>function User_InfoDel<%=rs("uId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {art.dialog.open('?action=delete&uId=<%=rs("uId")%>');return false;},cancel: true }); };</script>
	<%
	rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	%>
			</table>
		</td>
	</tr>
</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<%=EasyCrm.pagelist("User.asp", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<%
end sub

Sub deleteData()
	uId = Trim(Request("uId"))
	If uId = "" Then Exit Sub
	conn.execute("DELETE FROM [user] where uId = "&uId&" ")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
<script src="../data/calendar/WdatePicker.js"></script>
</body>
</html>
<%
else
Response.write"<script>alert("""&alert31&""");location.href=""../"";</script>"
end if
%><% Set EasyCrm = nothing %>
