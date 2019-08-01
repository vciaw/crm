<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))

'获取当前页码
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 

If Action = "searchItem" Then
    Dim TimeBegin,TimeEnd
	TimeBegin = Trim(Request("TimeBegin"))
	TimeEnd = Trim(Request("TimeEnd"))
	Session("Search_Log_TimeBegin") = Trim(Request("TimeBegin"))
	Session("Search_Log_TimeEnd") = Trim(Request("TimeEnd"))
	Dim sql
    sql = ""	
	
	if Accsql =1 then
	If TimeBegin <> "" Then
		sql = sql & " And olstarttime >= '" & TimeBegin & "' "
	End If
			
	If TimeEnd <> "" Then
		sql = sql & " And olstarttime <= '" & TimeEnd & "' "
	End If
	else
	If TimeBegin <> "" Then
		sql = sql & " And olstarttime >= #" & TimeBegin & "# "
	End If
			
	If TimeEnd <> "" Then
		sql = sql & " And olstarttime <= #" & TimeEnd & "# "
	End If
	End If
	
End If

If TimeBegin = "" And TimeEnd = "" Then
    If Session("CRM_Log_User") <> "" Then
        sql = Session("CRM_Log_User")
	End If
Else
    Session("CRM_Log_User") = sql
End If

If action = "killSession" Then
	Session("CRM_Log_User") = ""
	Session("Search_Log_TimeBegin")=""
	Session("Search_Log_TimeEnd")=""
	sql=""
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" oncontextmenu=self.event.returnValue=false><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>
<%
Select Case action
Case "delete"
    Call deleteData()
Case "clear"
    Call clear()
Case Else
	Call main()
End Select

Sub main()
%>
<body> 
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：日志管理 > 登录日志管理</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan=2 class="Search_All td_n">
				<form name="searchForm" method="post" action="?Action=searchItem">
				<input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" value="<%=Session("Search_Log_TimeBegin")%>" style="width:150px;" onFocus="WdatePicker()" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" value="<%=Session("Search_Log_TimeEnd")%>" style="width:150px;" onFocus="WdatePicker()" />　<input type="submit" name="Submit" class="button222" value=" <%=L_Search%> ">　<input type="button" name="button" class="button223" value=" <%=L_Clear%> " onClick=window.location.href="?action=killSession" />
				</form>
		</td>
	</tr>
	<tr>
		<td valign="top" colspan=2 style="padding:0 10px 10px 10px;" class="td_n">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
				<tr class="tr_t">
					<td width="80" class="td_l_c">编号</td>
					<td class="td_l_l">账户</td>
					<td width="150" class="td_l_c">时间</td>
					<td width="120" class="td_l_c">登录IP</td>
					<td width="50" class="td_l_c">管理</td>
				</tr>
				<%
				PN = CLng(ABS(Request("PN")))
				If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
				intPageSize = DataPageSize
				pagenum = intPageSize*(PN-1)
				Set rs = Server.CreateObject("ADODB.Recordset")
				IF PN=1 THEN
				rs.Open "Select top "&intPageSize&" * From [userlog] where 1=1 "&sql&" Order By Id desc ",conn,1,1 
				ELSE
				rs.Open "Select top "&intPageSize&" * From [userlog] where 1=1 "&sql&" and Id < ( SELECT Min(Id) FROM ( SELECT TOP "&pagenum&" Id FROM [userlog] where 1=1 "&sql&" ORDER BY Id desc ) AS T ) Order By Id desc ",conn,1,1
				END IF
				Set Rsstr=conn.Execute("Select count(Id) As RecordSum From [userlog] where 1=1 "&sql&"",1,1)
				TotalRecords=Rsstr("RecordSum") 
				if Int(TotalRecords/intPageSize)=TotalRecords/intPageSize then
				TotalPages=TotalRecords/intPageSize
				else
				TotalPages=Int(TotalRecords/intPageSize)+1
				end if
				Rsstr.Close:Set Rsstr=Nothing
				If PN > TotalPages Then PN = TotalPages
				If rs.RecordCount > 0 Then
				Do While Not rs.BOF And Not rs.EOF
				%>
				<tr class="tr">
					<td class="td_l_c"><%=rs("id")%></td>
					<td class="td_l_l"><%=rs("olname")%></a></td>
					<td class="td_l_c"><%=rs("olstarttime")%></td>
					<td class="td_l_c"><%=rs("olip")%></td>
					<td class="td_l_c"><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='LogUser_InfoDel<%=rs("id")%>()' style="cursor:pointer" /></td>
				</tr>
				<script>function LogUser_InfoDel<%=rs("id")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {art.dialog.open('?action=delete&id=<%=rs("id")%>&PN=<%=PNN%>');art.dialog.close();},cancel: true }); };</script>
				<%
				rs.MoveNext
				Loop
				else
				%>
							<tr><td class="td_l_l" colspan="5"><%=L_Notfound%></td></tr>
				<%
				end if
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
			<span class="r"><input name="Back" type="button" id="Back" class="button247" value="清空记录" onClick=window.location.href="?action=clear"></span>
			<%=EasyCrm.pagelist("Log_user.asp", PN,TotalPages,TotalRecords)%> 
		</td>
	</tr>
</table>
</div>
<%
End Sub

Sub deleteData()
    Dim id
	id = Trim(Request("id"))
	If id = "" Then
	Exit Sub
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [userlog] Where id = " & id,conn,3,2
	If rs.RecordCount > 0 Then
	    id = rs("id")
		rs.Delete
		rs.Update
	End If
	rs.Close
	Set rs = Nothing	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

End Sub

Sub clear()
	conn.execute("delete from userlog")
	Response.Redirect("?PN="&PNN)
End Sub
%>
</body>
</html>
<% Set EasyCrm = nothing %>
<script src="../data/calendar/WdatePicker.js"></script>