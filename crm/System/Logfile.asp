<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))

'获取当前页码
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 

If Action = "searchItem" Then
    Dim TimeBegin,TimeEnd,User,lAction,cId
	TimeBegin = Trim(Request("TimeBegin"))
	TimeEnd = Trim(Request("TimeEnd"))
	lUser = Trim(Request("User"))
	lAction = Trim(Request("lAction"))
	cId = Trim(Request("cId"))
	Session("Search_LogFile_TimeBegin") = Trim(Request("TimeBegin"))
	Session("Search_LogFile_TimeEnd") = Trim(Request("TimeEnd"))
	Session("Search_LogFile_User") = Trim(Request("User"))
	Session("Search_LogFile_lAction") = Trim(Request("lAction"))
	Dim sql
    sql = ""	
	
	If cId <> "" Then
		sql = sql & " And lcId = '" & cId & "' "
	End If		
	
	If lUser <> "" Then
		sql = sql & " And lUser = '" & lUser & "' "
	End If	
	
	If lAction <> "" Then
		sql = sql & " And lAction = '" & lAction & "' "
	End If
	
	if Accsql =1 then
		If TimeBegin <> "" and TimeEnd <> "" Then 
			sql = sql & " And lTime >= '" & TimeBegin & "' And lTime <= '" & TimeEnd & "' "
		End If
		If TimeBegin <> "" and TimeEnd = "" Then
			sql = sql & " And lTime = '" & TimeEnd & "' )"
		End If
	else
		If TimeBegin <> "" and TimeEnd <> "" Then 
			sql = sql & " And lTime >= #" & TimeBegin & "# And lTime <= #" & TimeEnd & "# "
		End If
		If TimeBegin <> "" and TimeEnd = "" Then
			sql = sql & " And lTime = #" & TimeEnd & "# )"
		End If
	End If
	
End If

If TimeBegin = "" And TimeEnd = "" And lUser = "" And lAction = "" And lcId = "" Then
    If Session("CRM_LogFile_User") <> "" Then
        sql = Session("CRM_LogFile_User")
	End If
Else
    Session("CRM_LogFile_User") = sql
End If

If action = "killSession" Then
	Session("CRM_LogFile_User") = ""
	Session("Search_LogFile_TimeBegin")=""
	Session("Search_LogFile_TimeEnd")=""
	Session("Search_LogFile_User")=""
	Session("Search_LogFile_lAction")=""
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
Case "ViewReason"
    Call ViewReason()
Case Else
	Call main()
End Select

Sub main()
%>
<body> 
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：日志管理 > 操作记录管理</td>
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
				<input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" value="<%=Session("Search_LogFile_TimeBegin")%>" style="width:100px;" onFocus="WdatePicker()" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" value="<%=Session("Search_LogFile_TimeEnd")%>" style="width:100px;" onFocus="WdatePicker()" />　业务员 : <% = EasyCrm.UserList(2,"User",Session("Search_LogFile_User")) %>　行为 : 
				<select name="lAction">
					<option value="">请选择</option>
					<option value="新增">新增</option>
					<option value="修改">修改</option>
					<option value="删除">删除</option>
				</select> 
				　
				编号 : <input name="cId" type="text" class="int" id="cId" size="5" >　
				<input type="submit" name="Submit" class="button222" value=" <%=L_Search%> ">　
				<input type="button" name="button" class="button223" value=" <%=L_Clear%> " onClick=window.location.href="?action=killSession" />
				</form>
				<script language="JavaScript">
				<!--
				for(var i=0;i<document.getElementById('lAction').options.length;i++){
					if(document.getElementById('lAction').options[i].value == "<% = Session("Search_LogFile_lAction") %>"){
					document.getElementById('lAction').options[i].selected = true;}}
				-->
				</script>
		</td>
	</tr>
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 ">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td width="80" class="td_l_c">编号</td>
					<td width="150" class="td_l_c">时间</td>
					<td class="td_l_l">企业名称</td>
					<td width="80" class="td_l_c">行为</td>
					<td width="80" class="td_l_c">数据表</td>
					<td width="80" class="td_l_c">原因</td>
					<td width="80" class="td_l_c">业务员</td>
					<td width="50" class="td_l_c">管理</td>
				</tr>
	<%
	PN = CLng(ABS(Request("PN")))
	If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
	intPageSize = DataPageSize
	pagenum = intPageSize*(PN-1)
	Set rs = Server.CreateObject("ADODB.Recordset")
	IF PN=1 THEN
	rs.Open "Select top "&intPageSize&" * From [Logfile] where 1=1 "&sql&" Order By lid desc ",conn,1,1 
	ELSE
	rs.Open "Select top "&intPageSize&" * From [Logfile] where 1=1 "&sql&" and lid < ( SELECT Min(lid) FROM ( SELECT TOP "&pagenum&" lid FROM [Logfile] where 1=1 "&sql&" ORDER BY lid desc ) AS T ) Order By lid desc ",conn,1,1
	END IF
	Set Rsstr=conn.Execute("Select count(lid) As RecordSum From [Logfile] where 1=1 "&sql&" ",1,1)
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
								<td class="td_l_c"><%=rs("lid")%></td>
								<td class="td_l_c"><%=rs("lTime")%></td>
								<td class="td_l_l"><a href='javascript:void(0)' onclick='Client_InfoView<%=rs("lcId")%>()' style="cursor:pointer"><%=EasyCrm.getNewItem("Client","cId",rs("lcId"),"cCompany")%></a></td>
								<td class="td_l_c"><%=rs("lAction")%></td>
								<td class="td_l_c"><%=rs("lClass")%></td>
								<td class="td_l_c"><%if rs("lReason")<>"" then%><input type="button" class="button226" value="查看" onclick='Logfile_InfoView<%=rs("lId")%>()' style="cursor:pointer" /><%end if%></td>
								<td class="td_l_c"><%=rs("lUser")%></td>
								<td class="td_l_c"><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Logfile_InfoDel<%=rs("lId")%>()' style="cursor:pointer" /></td>
							</tr>
							<script>function Client_InfoView<%=rs("lcId")%>() {$.dialog.open('../main/GetUpdate.asp?action=Client&sType=InfoView&cId=<%=rs("lcId")%>', {title: '查看', width: 900,height: 500, fixed: true}); };</script>
							<script>function Logfile_InfoView<%=rs("lId")%>() {
								art.dialog(
									{ 
										title: '操作原因', 
										icon: 'question',
										content: '<%=rs("lReason")%>',
										drag: false,
										resize: false
									}
								); 
							};</script>
							
							<script>function Logfile_InfoDel<%=rs("lId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {art.dialog.open('?action=delete&lid=<%=rs("lId")%>&PN=<%=PNN%>');art.dialog.close();},cancel: true }); };</script>
	<%
	rs.MoveNext
	Loop
	else
	%>
							<tr><td class="td_l_l" colspan="8"><%=L_Notfound%></td></tr>
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
			<%=EasyCrm.pagelist("Logfile.asp", PN,TotalPages,TotalRecords)%> 
		</td>
	</tr>
</table>
</div>
<%
end sub
%>


<%
Sub deleteData()
    Dim lid
	lid = Trim(Request("lid"))
	If lid = "" Then
	Exit Sub
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From Logfile Where lid = " & lid,conn,3,2
	If rs.RecordCount > 0 Then
	    lid = rs("lid")
		rs.Delete
		rs.Update
	End If
	rs.Close
	Set rs = Nothing	
	'Response.Redirect("?PN="&PNN)
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

Sub ViewReason()
	lid = Trim(Request("lid"))
%>

			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="GetUpdate.asp?action=AreaData&sType=SaveBigClassAdd" method="post">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l"><B>详细内容</B></td>
							</tr>
							<tr>
								<td class="td_l_l"></td>
							</tr>
						</table>
					</td>
				</tr>
				
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdb10"> 
						<div style="float:left;padding:10px 0;">
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
						</div>
					</td> 
				</tr>
				</form>
			</table>
<%
End Sub

Sub clear()
	conn.execute("delete from Logfile")
	Response.Redirect("?PN="&PNN)
End Sub
%>
</body>
</html>
<script src="../data/calendar/WdatePicker.js"></script><% Set EasyCrm = nothing %>