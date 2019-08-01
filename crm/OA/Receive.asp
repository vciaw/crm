<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 61, 1) = 1 Then %>
<%
	oType	=Request.QueryString("oType")
	Sendname=Request.QueryString("username")
	oIsread	=Request.QueryString("Isread")
	id		=Request.QueryString("id")
	if oIsread = "" then oIsread = 0
if Accsql = 1 then 
	Nowdate = "Getdate"
else
	Nowdate = "Now"
end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script src="<%=SiteUrl&skinurl%>Js/Common.js" type="text/javascript"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Float.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>
<body> 
<%if id="" then%>
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Page_OA%> > <%=L_Page_Receive%></td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdt10">   
			<div class="MenuboxS">
				 <ul>
					<li <%if oType="A" or oType="" then%>class="hover"<%end if%>><span><a href="?action=Main">收件箱</a></span></li>
					<li <%if oType="B" then%>class="hover"<%end if%>><span><a href="?action=Send&oType=B">已发送</a></span></li>
					<% If mid(Session("CRM_qx"), 62, 1) = 1 Then %>
					<li <%if oType="C" then%>class="hover"<%end if%>><span><a href="#" onclick='Receive_InfoAdd()' style="cursor:pointer">写信息</a></span></li>
					<%end if%>
				</ul>
			</div>
		</td>
	</tr>
</table>
<script>function Receive_InfoAdd() {$.dialog.open('../OA/Receive.asp?action=Add&oType=C&id=-1', {title: '新增', width: 900,height: 400, fixed: true}); };</script>
<%end if%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<%
action = Trim(Request("action"))
Select Case action
Case "Add"
    Call Add()
Case "SaveAdd"
    Call SaveAdd()
Case "Reply"
    Call Reply()
Case "SaveReply"
    Call SaveReply()
Case "View"
    Call View()
Case "Delete"
    Call Delete()
Case "Send"
    Call Send()
Case "SendView"
    Call SendView()
Case "SendDelete"
    Call SendDelete()
Case "ReceiveClear"
    Call ReceiveClear()
Case "SendClear"
    Call SendClear()
Case Else
	Call Main()
End Select

Sub Main() '收件箱
%>
	<tr>
		<td valign="top" colspan=2 style="padding:10px 10px 0;" class="td_n">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
				<tr class="tr_t">
					<td class="td_l_l" COLSPAN=5><B>【<%if oIsread="" or oIsread=0 then%>未读<%else%>已读<%end if%>】信息列表</B></td>
				</tr>
				<tr class="tr_b">
					<td width="80" class="td_l_c"><a href="?action=main<%if oIsread="" or oIsread=0 then%>&Isread=1<%else%>&Isread=0<%end if%>"><font color=red><%=L_Mms_oIsread_0%></font>/<%=L_Mms_oIsread_1%></a></td>
					<td width="100" class="td_l_c"><%=L_Mms_oSender%></td>
					<td class="td_l_l"><%=L_Mms_oTitle%></td>
					<td width="150" class="td_l_c"><%=L_Mms_oTime%></td>
					<td width="100" class="td_l_c"><%=L_Top_Manage%></td>
				</tr>
				<%
					Dim sql
					sql = ""
					if Sendname <>"" then
					sql = sql & " and oSender ='"&Sendname&"'"
					end if
					if oIsread = 1 then
					sql = sql & " and oIsread ="&oIsread&" and oReceiver = '"&Session("CRM_name")&"'"
					else
					sql = sql & " and ( oIsread = 0 and oReceiver = '"&Session("CRM_name")&"' and oAttime is null ) or ( oAttime is not null and oAttime < "&Nowdate&"()+ 0.007 and oIsread = 0 and oReceiver = '"&Session("CRM_name")&"' ) "
					end if
					
				PN = CLng(ABS(Request("PN")))
				If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
				intPageSize = DataPageSize
				pagenum = intPageSize*(PN-1)
				Set rs = Server.CreateObject("ADODB.Recordset")
				
				IF PN=1 THEN
				rs.Open "Select top "&intPageSize&" * From [OA_mms_Receive] where 1 = 1 "&sql&" Order By Id Desc",conn,1,1
				ELSE
				rs.Open "Select top "&intPageSize&" * From [OA_mms_Receive] where 1 = 1 "&sql&" and Id < ( SELECT Min(Id) FROM ( SELECT TOP "&pagenum&" Id FROM [OA_mms_Receive]  where 1 = 1 "&sql&" ORDER BY Id desc ) AS T ) Order By Id Desc ",conn,1,1
				END IF
				Set Rsstr=conn.Execute("Select count(Id) As RecordSum From [OA_mms_Receive] where 1 = 1 "&sql&" ",1,1)
				
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
					
					<td class="td_l_c"><a href="?action=Main&username=<%=rs("oSender")%>" title="只看<%=rs("oSender")%>的信息" ><%=rs("oSender")%></a></td>
					
					<td class="td_l_l"><a onclick='Receive_InfoEdit<%=rs("Id")%>()' style="cursor:pointer" ><%=rs("oTitle")%></a></td>
					<td class="td_l_c"><%=EasyCrm.FormatDate(rs("oTime"),1)%></td>
					
					<td class="td_l_c"><input type="button" class="button_info_edit" value=" " title="<%=L_Reply%>"  onclick='Receive_InfoEdit<%=rs("Id")%>()' style="cursor:pointer" /> <input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Report_InfoDel<%=rs("Id")%>()' style="cursor:pointer" /></td>
				</tr>
				<script>function Receive_InfoView<%=rs("Id")%>() {$.dialog.open('../OA/Receive.asp?action=View&id=<%=rs("id")%>', {title: '查看', width: 900,height: 500, fixed: true}); };</script>
				<script>function Receive_InfoEdit<%=rs("Id")%>() {$.dialog.open('../OA/Receive.asp?action=Reply&id=<%=rs("Id")%>', {title: '回复', width: 900,height: 470, fixed: true}); };</script>
				
				<script>function Report_InfoDel<%=rs("Id")%>()
				{
					art.dialog({
						content: '<%=Alert_del_YN%>',
						icon: 'error',
						ok: function () {
							art.dialog.open('../OA/Receive.asp?action=Delete&Id=<%=rs("id")%>');
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
				<tr><td class="td_l_l" colspan=5><%=L_Notfound%></td></tr>
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
			<span class=r><input name="Back" type="button" class="button227" id="Back" value=" <%=l_Clear%> " onClick="location.href='?action=ReceiveClear';"></span>
			<%=EasyCrm.pagelist("Receive.asp?Action=main&username="&Sendname&"&Isread="&oIsread&"", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<% End Sub	

Sub Send() '发件箱
%>
	<tr>
		<td valign="top" colspan=2 style="padding:10px 10px 0;" class="td_n">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
				<tr class="tr_t">
					<td class="td_l_l" COLSPAN=4><B>信息列表</B></td>
				</tr>
				<tr class="tr_b">
					<td width="100" class="td_l_c"><%=L_Mms_oReceiver%></td>
					<td class="td_l_l"><%=L_Mms_oTitle%></td>
					<td width="150" class="td_l_c"><%=L_Mms_oTime%></td>
					<td width="100" class="td_l_c"><%=L_Top_Manage%></td>
				</tr>
				<%
					sql = ""
					sql = sql & " and oSender ='"&Session("CRM_name")&"'"
					if Sendname <>"" then
					sql = sql & " and oReceiver ='"&Sendname&"'"
					end if
					
				PN = CLng(ABS(Request("PN")))
				If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
				intPageSize = DataPageSize
				pagenum = intPageSize*(PN-1)
				Set rs = Server.CreateObject("ADODB.Recordset")
				
				IF PN=1 THEN
				rs.Open "Select top "&intPageSize&" * From [OA_mms_send] where 1 = 1 "&sql&" Order By Id Desc",conn,1,1
				ELSE
				rs.Open "Select top "&intPageSize&" * From [OA_mms_send] where 1 = 1 "&sql&" and Id < ( SELECT Min(Id) FROM ( SELECT TOP "&pagenum&" Id FROM [OA_mms_send]  where 1 = 1 "&sql&" ORDER BY Id desc ) AS T ) Order By Id Desc ",conn,1,1
				END IF
				Set Rsstr=conn.Execute("Select count(Id) As RecordSum From [OA_mms_send] where 1 = 1 "&sql&" ",1,1)
				
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
					<%if UBound(split(rs("oReceiver"),",")) > 0 then%>
					<td class="td_l_c"><%=L_More%></td>
					<%else%>
					<td class="td_l_c"><%=rs("oReceiver")%></td>
					<%end if%>
					
					<td class="td_l_l" onmouseover="tip.start(this)" tips="<%=rs("oContent")%>"><%=rs("oTitle")%></td>
					<td class="td_l_c"><%=EasyCrm.FormatDate(rs("oTime"),1)%></td>
					
					<td class="td_l_c"><% If mid(Session("CRM_qx"), 64, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Send_InfoDel<%=rs("Id")%>()' style="cursor:pointer" /><%end if%></td>
				</tr>
				
				<script>function Send_InfoDel<%=rs("Id")%>()
				{
					art.dialog({
						content: '<%=Alert_del_YN%>',
						icon: 'error',
						ok: function () {
							art.dialog.open('../OA/Receive.asp?action=SendDelete&Id=<%=rs("id")%>');
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
				<tr><td class="td_l_l" colspan=4><%=L_Notfound%></td></tr>
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
			<% If mid(Session("CRM_qx"), 64, 1) = 1 Then %>
			<span class=r><input name="Back" type="button" class="button227" id="Back" value=" <%=l_Clear%> " onClick="location.href='?action=SendClear';"></span><%end if%>
			<%=EasyCrm.pagelist("Receive.asp?Action=Send&oType=B", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<% End Sub	%>

<% Sub Add() %>
	<script language="JavaScript">
	<!-- 客户档案必填项提示
	function CheckInput()
	{
		if(document.getElementById('oReceiver').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Mms_oReceiver & alert04%>'});document.getElementById('oReceiver').focus();return false;}
		if(document.getElementById('oTitle').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Mms_oTitle & alert04%>'});document.getElementById('oTitle').focus();return false;}
		if(document.getElementById('oContent').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Mms_oContent & alert04%>'});document.getElementById('oContent').focus();return false;}
	}
	-->
	</script>
	<form name="Save" action="?action=SaveAdd" method="post" onSubmit="return CheckInput();">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
				</tr>
				<tr>
					<td class="td_l_c title" width="100"><%=L_Mms_oReceiver%> <font color="#FF0000">*</font></td>
					<td class="td_r_l"> <input type="text" class="int" name="oReceiver" id="oReceiver" value="" readonly style="width:80%" maxlength="50"> <span class="info_help " onclick="AddReceive()" style="cursor:pointer;">&nbsp;<a>选择收件人</a></span>
					<script>function AddReceive() {$.dialog.open('../OA/GetUpdate.asp?action=Receiver&sType=Receiver', {title: '快速选择', width: 640, height: 500,fixed: true}); };</script></td>
				</tr>
				<tr>
					<td class="td_l_c title"><%=L_Mms_oTitle%> <font color="#FF0000">*</font></td>
					<td class="td_r_l"> <input type="text" class="int" name="oTitle" id="oTitle" style="width:80%" maxlength="50"> </td>
				</tr>
				<tr>
					<td class="td_l_c title" valign="top"><%=L_Mms_oContent%> <font color="#FF0000">*</font></td>
					<td class="td_r_l" style="padding:10px;"> <textarea name="oContent" id="oContent" class="int" style="width:81%;height:200px;"></textarea></td>
				</tr>
			</table>
		</td> 
	</tr>
	<tr>
		<td valign="top" class="td_n pdl10 pdr10"> 
			<div style="float:left;padding:10px 0 0;">
				<input name="oSender" type="hidden" value="<%=Session("CRM_name")%>">
				<input type="submit" name="Submit" class="button45" value=" <%=l_Submit%> " >　
				<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
			</div>
		</td> 
	</tr>
    </form>

<% End Sub

Sub Reply()
  Dim ONtitle,ONcontent,ONedittime
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From OA_mms_Receive where id="&id,conn,1,1
	oSender = rs("oSender")
	oReceiver = rs("oReceiver")
	oContent = rs("oContent")
	oTime = rs("oTime")
	oTitle = Replace(rs("oTitle"),""&L_Reply&"：","")
	rs.Close
	Set rs = Nothing
	conn.execute "UPDATE OA_mms_Receive SET oIsread='1' Where id="&id
%>
	<form name="Save" action="?action=SaveReply" method="post" onSubmit="return CheckInput();">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="120" >
				<tr class="tr_t"> 
					<td class="td_l_l" style="border-right:0;"><B><%=L_Mms_oSender%>：<%=oSender%> </B></td>
					<td class="td_l_r"><B><%=EasyCrm.FormatDate(oTime,1)%> </B></td>
				</tr>
				<tr>
					<td class="td_l_c title"><%=L_Mms_oTitle%></td>
					<td class="td_r_l"> <%=oTitle%> </td>
				</tr>
				<tr>
					<td class="td_l_c title" valign="top"><%=L_Mms_oContent%></td>
					<td class="td_r_l"> <%=oContent%></td>
				</tr>
			</table>
		</td> 
	</tr>
	<% If mid(Session("CRM_qx"), 63, 1) = 1 Then %>
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="120" >
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="2"><B><%=L_Top_Mms_reply%> </B></td>
				</tr>
				<tr>
					<td class="td_l_c title" width="120" valign="top"><%=L_Mms_oContent%> <font color="#FF0000">*</font></td>
					<td class="td_r_l" style="padding:10px;"> <textarea name="oContent" id="oContent" style="width:99%;height:200px;"></textarea></td>
				</tr>
			</table>
		</td> 
	</tr>
	<%end if%>
	<tr>
		<td valign="top" class="td_n pdl10 pdr10"> 
			<div style="float:left;padding:10px 0 0;">
				<input name="oTitle" type="hidden" value="<%=L_Reply%>：<%=oTitle%>">
				<input name="oReceiver" type="hidden" value="<%=oSender%>">
				<input name="oSender" type="hidden" value="<%=Session("CRM_name")%>">
				<% If mid(Session("CRM_qx"), 63, 1) = 1 Then %>
				<input type="submit" name="Submit" class="button45" value=" <%=l_Submit%> " >　
				<%end if%>
				<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
			</div>
		</td> 
	</tr>
    </form>
<% End Sub 

Sub SaveAdd()
    Dim oReceiver,oSender,oTitle,oContent,oIsread
	oReceiver = Request.Form("oReceiver")
	oSender = Request.Form("oSender")
	oTitle = Request.Form("oTitle")
	oContent = Request.Form("oContent")
	conn.execute("insert into [OA_mms_send] (oReceiver,oSender,oTitle,oContent,oIsread,oTime) values('"&oReceiver&"','"&oSender&"','"&oTitle&"','"&oContent&"',0,'"&Now()&"')")
	arrReceive=split(oReceiver,",")
	for i=0 to ubound(arrReceive)
	conn.execute("insert into [OA_mms_Receive] (oReceiver,oSender,oTitle,oContent,oIsread,oTime) values('"&arrReceive(i)&"','"&oSender&"','"&oTitle&"','"&oContent&"',0,'"&Now()&"')")
	next
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub

Sub SaveReply()
	Dim oReceiver,oSender,oTitle,oContent,oIsread
	oReceiver = Request.Form("oReceiver")
	oSender = Request.Form("oSender")
	oTitle = Request.Form("oTitle")
	oContent = Request.Form("oContent")
	conn.execute("insert into [OA_mms_send] (oReceiver,oSender,oTitle,oContent,oIsread,oTime) values('"&oReceiver&"','"&oSender&"','"&oTitle&"','"&oContent&"',0,'"&Now()&"')")
	conn.execute("insert into [OA_mms_Receive] (oReceiver,oSender,oTitle,oContent,oIsread,oTime) values('"&oReceiver&"','"&oSender&"','"&oTitle&"','"&oContent&"',0,'"&Now()&"')")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub

Sub Delete()
	If Id = "" Then Exit Sub
	conn.execute("DELETE FROM [OA_mms_Receive] where Id = "&Id&" ")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

Sub SendDelete()
	If Id = "" Then Exit Sub
	conn.execute("DELETE FROM [OA_mms_send] where Id = "&Id&" ")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

Sub ReceiveClear()
	conn.execute("DELETE FROM [OA_mms_Receive] where oReceiver ='"&Session("CRM_name")&"' ")
	Response.Write "<script>location.href='?action=Main&oType=A';</script>"
End Sub

Sub SendClear()
	conn.execute("DELETE FROM [OA_mms_send] where oSender ='"&Session("CRM_name")&"' ")
	Response.Write "<script>location.href='?action=Send&oType=B';</script>"
End Sub

%>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html><%else%>无权限<%end if%><% Set EasyCrm = nothing %>
