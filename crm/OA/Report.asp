<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 66, 1) = 1 Then %>
<%
	oType=Request.QueryString("oType")
	oClass=Request.QueryString("oClass")
	oIsread=Request.QueryString("oIsread")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script src="<%=SiteUrl&skinurl%>Js/Common.js" type="text/javascript"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>

<body> 
<style>body{padding:35px 0 48px;}</style>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Page_OA%> > <%=L_Page_Report%></td>
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
					<li <%if oType="A" or oType="" then%>class="hover"<%end if%>><span><a href="?action=main">全部</a></span></li>
					<li <%if oType="B" then%>class="hover"<%end if%>><span><a href="?action=main&oType=B&oIsread=0">未读</a></span></li>
					<li <%if oType="C" then%>class="hover"<%end if%>><span><a href="?action=main&oType=C&oclass=<%=L_Ribao%>"><%=L_Ribao%></a></span></li>
					<li <%if oType="D" then%>class="hover"<%end if%>><span><a href="?action=main&oType=D&oclass=<%=L_Zhoubao%>"><%=L_Zhoubao%></a></span></li>
					<li <%if oType="E" then%>class="hover"<%end if%>><span><a href="?action=main&oType=E&oclass=<%=L_Yuebao%>"><%=L_Yuebao%></a></span></li>
					<li <%if oType="F" then%>class="hover"<%end if%>><span><a href="?action=main&oType=F&oclass=<%=L_Jibao%>"><%=L_Jibao%></a></span></li>
					<li <%if oType="G" then%>class="hover"<%end if%>><span><a href="?action=main&oType=G&oclass=<%=L_Nianbao%>"><%=L_Nianbao%></a></span></li>
					<% If mid(Session("CRM_qx"), 67, 1) = 1 Then %>
					<li ><span><a href="#" onclick='Report_Add()' style="cursor:pointer">写报告</a></span></li>
					<%end if%>
				</ul>
			</div>
		</td>
	</tr>
</table>

<script>function Report_Add() {$.dialog.open('GetUpdate.asp?action=Report&sType=Add', {title: '新窗口', width: 800, height: 500,fixed: true}); };</script>
<%
action = Trim(Request("action"))
Select Case action
Case "add"
    Call add()
Case "saveadd"
    Call saveadd()
Case "savereply"
    Call savereply()
Case "view"
    Call view()
Case "delete"
    Call deleteData()
Case Else
	Call main()
End Select

Sub main()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr> 
		<td valign="top" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" colspan=2 style="padding:10px 10px 0;" class="td_n">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_t">
								<td class="td_l_l" COLSPAN=6><B>信息列表</B></td>
							</tr>
							
							<tr class="tr_b">
								<td width="80" class="td_l_c"><%=L_Report_oIsread%></td>
								<td width="100" class="td_l_c"><%=L_Report_oClass%></td>
								<td class="td_l_l"><%=L_Report_oTitle%></td>
								<td width="100" class="td_l_c"><%=L_Report_oUser%></td>
								<td width="150" class="td_l_c"><%=L_Report_oTime%></td>
								<td width="100" class="td_l_c"><%=L_Top_Manage%></td>
							</tr>
							<%
								If Session("CRM_level") = 9 Then									'管理员管理所有的
								sql = sql & " "
								elseIf Session("CRM_level") < 9 and Session("CRM_level") > 1 Then 	'部门经理可以看到别人提交给自己的，和自己提交的
								sql = sql & " and ( oSendto like '%"&Session("CRM_name")&"%' or oUser = '"&Session("CRM_name")&"' ) "
								else								
								sql = sql & " and oUser = '"&Session("CRM_name")&"' " 				'普通员工只能看到自己提交的报告
								end if
								
								if oClass <>"" then
								sql = sql & " and oClass ='"&oClass&"'"
								end if
								
								if oIsread <>"" then
								sql = sql & " and oIsread ="&oIsread&""
								end if
								
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [OA_Report] where 1 = 1 "&sql&" Order By Id Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [OA_Report] where 1 = 1 "&sql&" and Id < ( SELECT Min(Id) FROM ( SELECT TOP "&pagenum&" Id FROM [OA_Report]  where 1 = 1 "&sql&" ORDER BY Id desc ) AS T ) Order By Id Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(Id) As RecordSum From [OA_Report] where 1 = 1 "&sql&" ",1,1)
							
							TotalRecords=Rsstr("RecordSum") 
							if Int(TotalRecords/intPageSize)=TotalRecords/intPageSize then
							TotalPages=TotalRecords/intPageSize
							else
							TotalPages=Int(TotalRecords/intPageSize)+1
							end if
							Rsstr.Close
							Set Rsstr=Nothing
							If PN > TotalPages Then PN = TotalPages
							If rs.RecordCount > 0 Then
							Do While Not rs.BOF And Not rs.EOF
							%>
							<tr class="tr">
								<%if rs("oIsread") = 0 then%>
								<td class="td_l_c"><img src="<%=SiteUrl&skinurl%>images/ico/message_new.png"></td>
								<%else%>
								<td class="td_l_c"><img src="<%=SiteUrl&skinurl%>images/ico/message_old.png"></td>
								<%end if%>
								
								<td class="td_l_c"><a href="?action=main&oClass=<%=rs("oClass")%>"><%=rs("oClass")%></a></td>
								
								<td class="td_l_l"><a onclick='Report_InfoView<%=rs("Id")%>()' style="cursor:pointer"><%=rs("oTitle")%></a></td>
								<td class="td_l_c"><%=rs("oUser")%></td>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("oTime"),1)%></td>
								
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 68, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Pizhu%>"  onclick='Report_InfoView<%=rs("Id")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 69, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Report_InfoDel<%=rs("Id")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Report_InfoView<%=rs("Id")%>() {$.dialog.open('GetUpdate.asp?action=Report&sType=Reply&id=<%=rs("id")%>', {title: '查看', width: 900,height: 500, fixed: true}); };</script>
							
							<script>function Report_InfoDel<%=rs("Id")%>()
							{
								art.dialog({
									content: '<%=Alert_del_YN%>',
									icon: 'error',
									ok: function () {
										art.dialog.open('?action=delete&Id=<%=rs("id")%>');
										art.dialog.close();
									},
									cancelVal: '关闭',
									cancel: true
								});
							};
							</script>
							<%
							rs.MoveNext
							Loop
							else
							%>
							<tr><td class="td_l_l" colspan=16><%=L_Notfound%></td></tr>
							<%
							end if
							rs.Close
							Set rs = Nothing
							%>
						</table> 
					</td>
				</tr>
			</table> 
        </td>
	</tr>
</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<%=EasyCrm.pagelist("Report.asp?oClass="&oClass&"&oIsread="&oIsread&"", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<% End Sub	

Sub deleteData()
	id = Trim(Request("id"))
	If id = "" Then
	Exit Sub
	End If
	conn.execute("DELETE FROM [OA_Report] where Id = "&Id&" ")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
</body>
</html><%else%>无权限<%end if%>
<% Set EasyCrm = nothing %>