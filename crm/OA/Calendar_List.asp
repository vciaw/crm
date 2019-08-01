<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 73, 1) = 1 Then %>
<%
action = Trim(Request("action"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Calendar_List.asp"
Session("CRM_pagenum") = PNN

If Action = "searchItem" Then
    Dim calendaruser,TimeBegin,TimeEnd
	calendaruser = Trim(Request("User"))
	TimeBegin = Trim(Request("TimeBegin"))
	TimeEnd = Trim(Request("TimeEnd"))
	Session("CRM_Calendar_User") = Trim(Request("User"))
	Session("CRM_Calendar_TimeBegin") = Trim(Request("TimeBegin"))
	Session("CRM_Calendar_TimeEnd") = Trim(Request("TimeEnd"))
	Dim sql
    sql = ""	
	
    If calendaruser <> "" Then
	    sql = sql & " And calendaruser = '" & calendaruser & "'"
	End If	
	
	if Accsql =1 then
	If TimeBegin <> "" and  TimeEnd <> "" Then
	    sql = sql & " And calendarDate >= '" & TimeBegin & "' And calendarDate <= '" & TimeEnd & "' "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And calendarDate = '" & TimeBegin & "' "
	End If
	else
	If TimeBegin <> "" and TimeEnd <> "" Then
	    sql = sql & " And calendarDate >= #" & TimeBegin & "# And calendarDate <= #" & TimeEnd & "# "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And calendarDate = #" & TimeBegin & "# "
	End If
	end if

	If Session("CRM_level") < 9 Then
		sql = sql & " And calendaruser In (" & arrUser & ")"
	End If
	
End If

If calendaruser = "" And TimeBegin = "" And TimeEnd = "" Then
    If Session("CRM_Search_Calendar") <> "" Then
        sql = Session("CRM_Search_Calendar")
	Else
	    If Session("CRM_level") < 9 Then
		    sql = " And calendaruser In (" & arrUser & ")"
		End If
	End If
Else
    Session("CRM_Search_Calendar") = sql
End If

If action = "killSession" Then
	Session("CRM_Search_Calendar") = ""
	Session("CRM_Calendar_User") = ""
	Session("CRM_Calendar_TimeBegin") = ""
	Session("CRM_Calendar_TimeEnd") = ""
	If Session("CRM_level") < 9 Then
		sql = " And calendaruser In (" & arrUser & ")"
	else
		sql=""
	end if
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>
</head>

<body oncontextmenu=self.event.returnValue=false> 
<style>body{padding:35px 0 48px;}</style> 
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Page_OA%> > <%=L_Page_Calendar%></td>
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
					<li class=""><span><a href="#" onclick="location.href='Calendar.asp';" style="cursor:pointer">月历视图</a></span></li>
					<li class="hover"><span><a href="?" >列表视图</a></span></li>
				</ul>
			</div>
		</td>
		<form name="searchForm" action="?Action=searchItem" method="post">
		<td class="td_l_r pdr10" COLSPAN="6" style="border-right:0;padding-top:5px;">
			<span class="tips01" style="float:left;padding:0 10px;height:34px;text-align:left;position:fixed;right:10px;top:43px;color:#000;">
			<% if Session("CRM_level") = 9 then %><% = EasyCrm.UserList(2,"User",Session("CRM_Calendar_User")) %><%else%><% = EasyCrm.UserList(1,"User",Session("CRM_Calendar_User")) %><%end if%>　
			<input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" style="width:100px;" value="<%=Session("CRM_Calendar_TimeBegin")%>" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" style="width:100px;" value="<%=Session("CRM_Calendar_TimeEnd")%>" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" />　
			<input type="submit" name="Submit" class="button222" value=" <%=L_Search%> ">　
			<input type="button" name="button" class="button223" value=" <%=L_Clear%> " onClick=window.location.href="?action=killSession" />
			</span>
		</td>
		</form> 
	</tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr> 
		<td valign="top" class="td_n pdl10 pdr10 pdt10 pdb10">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" colspan=2 class="td_n">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_t">
								<td width="60" class="td_l_c"><%=L_Calendar_ID%></td>
								<td width="100" class="td_l_c"><%=L_Calendar_calendarDate%></td>
								<td class="td_l_l"><%=L_Calendar_calendarText%></td>
								<td width="100" class="td_l_c"><%=L_Calendar_calendaruser%></td>
								<td width="100" class="td_l_c"><%=L_Top_Manage%></td>
							</tr>
							
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [calendar] where 1 = 1 "&sql&" Order By Id Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [calendar] where 1 = 1 "&sql&" and Id < ( SELECT Min(Id) FROM ( SELECT TOP "&pagenum&" Id FROM [calendar]  where 1 = 1 "&sql&" Order BY Id desc ) AS T ) Order By Id Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(Id) As Servicenum From [calendar] where 1 = 1 "&sql&" ",1,1)
							
							TotalService=Rsstr("Servicenum") 
							if Int(TotalService/intPageSize)=TotalService/intPageSize then
							TotalPages=TotalService/intPageSize
							else
							TotalPages=Int(TotalService/intPageSize)+1
							end if
							Rsstr.Close
							Set Rsstr=Nothing
							If PN > TotalPages Then PN = TotalPages
							If rs.RecordCount > 0 Then
							Do While Not rs.BOF And Not rs.EOF
							%>
							<tr class="tr">
								<td class="td_l_c"><%=rs("Id")%></td>
								<td class="td_l_c"><%=rs("calendarDate")%></td>
								<td class="td_l_l" onclick='InfoEdit<%=rs("Id")%>()' style="cursor:pointer"><%=left(rs("calendarText"),20)%></td>
								<td class="td_l_c"><%=rs("calendaruser")%></td>
								<td class="td_l_c"><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='InfoEdit<%=rs("Id")%>()' style="cursor:pointer" /> <input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='InfoDel<%=rs("Id")%>()' style="cursor:pointer" /></td>
							</tr>
							
							<script>function InfoEdit<%=rs("id")%>() {$.dialog.open('GetUpdate.asp?action=Calendar&sType=Edit&id=<%=rs("id")%>', {title: '编辑', width: 400, height: 280,fixed: true}); };</script>
							
							<script>function InfoDel<%=rs("Id")%>() {$.dialog.open('GetUpdate.asp?action=Calendar&sType=Del&id=<%=rs("Id")%>');art.dialog.close(); };</script>
							
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
        </td>
	</tr>
</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd ">
			<%=EasyCrm.pagelist("Calendar_List.asp", PN,TotalPages,TotalService)%>
		</td> 
	</tr>
</table>
</div>
</body>
</html><%else%>无权限<%end if%>
<% Set EasyCrm = nothing %>
<script src="../data/calendar/WdatePicker.js"></script>