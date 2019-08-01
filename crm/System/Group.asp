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
		<td class="top_left td_t_n td_r_n">当前位置：系统管理 > 部门管理</td>
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
					<li class="hover"><span><a href="?">部门列表</a></span></li>
					<li class=""><span><a href="#" onclick='Group_Add()' style="cursor:pointer">新增部门</a></span></li>
				</ul>
			</div>
		</td>
	</tr>
</table>
<script>function Group_Add() {$.dialog.open('GetGroup.asp?action=Group&sType=Add', {title: '新窗口', width: 400, height: 170,fixed: true}); };</script>

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
					<td class="td_l_l">部门名称</td>
					<td width="90" class="td_l_c">管理</td>
				</tr>
	<%
	PN = CLng(ABS(Request("PN")))
	If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
	intPageSize = DataPageSize
	pagenum = intPageSize*(PN-1)
	Set rs = Server.CreateObject("ADODB.Recordset")
	IF PN=1 THEN
	rs.Open "Select top "&intPageSize&" * From [system_group] Order By gId asc ",conn,1,1 
	ELSE
	rs.Open "Select top "&intPageSize&" * From [system_group] where gId > ( SELECT Max(gId) FROM ( SELECT TOP "&pagenum&" gId FROM [system_group] ORDER BY gId asc ) AS T ) Order By gId Asc ",conn,1,1
	END IF
	Set Rsstr=conn.Execute("Select count(gId) As RecordSum From [system_group] ",1,1)
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
								<td class="td_l_c"><%=rs("gId")%></td>
								<td class="td_l_l"><a href='#' onclick='Group_InfoEdit<%=rs("gId")%>()' style="cursor:pointer"><%=rs("gName")%></a></td>
								<td class="td_l_c"><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Group_InfoEdit<%=rs("gId")%>()' style="cursor:pointer" /> <input type="button" class="button_info_del" value=" " title="<%=L_Del%>" <%if EasyCrm.getCountItem("User","uid","uid"," and uGroup = "&rs("gId")&" ")=0 then%> onclick='Group_InfoDel<%=rs("gId")%>()'<%else%>onClick="art.dialog({title: '提示',time: 2,icon: 'warning',content: '有关联数据，无法删除！'}); "<%end if%> style="cursor:pointer" /></td>
							</tr>
							<script>function Group_InfoEdit<%=rs("gId")%>() {$.dialog.open('GetGroup.asp?action=Group&sType=Edit&Id=<%=rs("gId")%>', {title: '新窗口', width: 400,height: 170, fixed: true}); };</script>
							<script>function Group_InfoDel<%=rs("gId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {art.dialog.open('?action=delete&gId=<%=rs("gId")%>');return false;},cancel: true }); };</script>
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
			<%=EasyCrm.pagelist("Group.asp", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<%end Sub

Sub deleteData()
	gId = Trim(Request("gId"))
	If gId = "" Then Exit Sub
	conn.execute("DELETE FROM [system_group] where gId = "&gId&" ")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
</body>
</html>
<%
else
Response.write"<script>alert("""&alert31&""");location.href=""../"";</script>"
end if
%><% Set EasyCrm = nothing %>