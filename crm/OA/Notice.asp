<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 56, 1) = 1 Then %>
<%
action = Trim(Request("action"))
oType=Request.QueryString("oType")
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Notice.asp"
Session("CRM_pagenum") = PNN
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
<%=AddNoticeInput()%>
</head>

<body oncontextmenu=self.event.returnValue=false> 
<style>body{padding:35px 0 48px;}</style>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Page_OA%> > <%=L_Page_Notice%></td>
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
					<li <%if oType="" then%>class="hover"<%end if%>><span><a href="?">所有公文</a></span></li>
					<%
					Set rs = Server.CreateObject("ADODB.Recordset")
					rs.Open "Select SelectID,Select_NoticeClass From [SelectData] where Select_NoticeClass <>'' and Select_NoticeClass<>'Null'",conn,1,1
					If rs.RecordCount > 0 Then
					Do While Not rs.BOF And Not rs.EOF
					%>
					<li <%if oType=""&rs("SelectID")&"" then%>class="hover"<%end if%> ><span><a href="?action=Main&otype=<%=rs("SelectID")%>"><%=rs("Select_NoticeClass")%></a></span></li>
					<%
					rs.MoveNext
					Loop
					end if
					rs.Close
					Set rs = Nothing
					%>
					<% If mid(Session("CRM_qx"), 57, 1) = 1 Then %>
					<li class=""><span><a href="#" onclick='Notice_InfoAdd()' style="cursor:pointer">添加公文</a></span></li>
					<%end if%>
				</ul>
			</div>
		</td>
	</tr>
</table>

<script>function Notice_InfoAdd() {$.dialog.open('GetUpdate.asp?action=Notice&sType=Add', {title: '新窗口', width: 800, height: 480,fixed: true}); };</script>

<%
Select Case action
Case "setstar"
    Call setstar()
Case "setstarno"
    Call setstarno()
Case "delete"
    Call deleteData()
Case Else
	Call main()
End Select

Sub main()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" colspan=2 style="padding:10px 10px 0;" class="td_n">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_t">
								<td class="td_l_l" COLSPAN=6><B>信息列表</B></td>
							</tr>
							<tr class="tr_b">
								<td width="60" class="td_l_c"><%=L_Notice_ONid%></td>
								<td width="60" class="td_l_c"><%=L_Notice_ONStar%></td>
								<td width="120" class="td_l_c"><%=L_Notice_ONclass%></td>
								<td class="td_l_l"><%=L_Notice_ONtitle%></td>
								<td width="150" class="td_l_c"><%=L_Notice_ONaddtime%></td>
								<td width="100" class="td_l_c"><%=L_Top_Manage%></td>
							</tr>
							<%
							if oType <>"" then
								sql = sql & " and ONclass ='"&EasyCrm.getNewItem("SelectData","SelectId",""&oType&"","Select_NoticeClass")&"'"
							end if
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [OA_Notice] where 1 = 1 "&sql&" Order By ONid Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [OA_Notice] where 1 = 1 "&sql&" and ONid < ( SELECT Min(ONid) FROM ( SELECT TOP "&pagenum&" ONid FROM [OA_Notice]  where 1 = 1 "&sql&" Order BY ONid desc ) AS T ) Order By ONid Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(ONid) As Servicenum From [OA_Notice] where 1 = 1 "&sql&" ",1,1)
							
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
								<td class="td_l_c"><%=rs("ONid")%></td>
								<% If mid(Session("CRM_qx"), 58, 1) = 1 Then %>
								<%if rs("ONStar") = 0 then%>
								<td class="td_l_c"><a href="?action=setstar&ONid=<%=rs("ONid")%>"><img src="<%=SiteUrl&Skinurl%>images/ico/starno.png" border=0></a></td>
								<%else%>
								<td class="td_l_c"><a href="?action=setstarno&ONid=<%=rs("ONid")%>"><img src="<%=SiteUrl&Skinurl%>images/ico/star.png" border=0></a></td>
								<%end if%>
								<%else%>
								<%if rs("ONStar") = 0 then%>
								<td class="td_l_c"><img src="<%=SiteUrl&Skinurl%>images/ico/starno.png" border=0></td>
								<%else%>
								<td class="td_l_c"><img src="<%=SiteUrl&Skinurl%>images/ico/star.png" border=0></td>
								<%end if%>
								<%end if%>
								<td class="td_l_c"><%=rs("ONclass")%></td>
								<td class="td_l_l"><a onclick='Notice_InfoView<%=rs("ONId")%>()' style="cursor:pointer"><%=rs("ONtitle")%></a></td>
								<td class="td_l_c"><%=rs("ONaddtime")%></td>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 58, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Notice_InfoEdit<%=rs("ONId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 59, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Notice_InfoDel<%=rs("ONId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Notice_InfoView<%=rs("ONId")%>() {$.dialog.open('GetUpdate.asp?action=Notice&sType=View&Id=<%=rs("ONId")%>', {title: '查看', width: 800,height: 480, fixed: true}); };</script>
							
							<script>function Notice_InfoEdit<%=rs("ONId")%>() {$.dialog.open('GetUpdate.asp?action=Notice&sType=Edit&Id=<%=rs("ONId")%>', {title: '编辑', width: 800,height: 480, fixed: true}); };</script>
							
							<script>function Notice_InfoDel<%=rs("ONId")%>(){art.dialog({content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {
								art.dialog.open('?action=delete&ONid=<%=rs("ONId")%>');art.dialog.close();},cancelVal: '关闭',cancel: true});};
							</script>
							<%
							rs.MoveNext
							Loop
							else
							%>
							<tr><td class="td_l_l" colspan="6"><%=L_Notfound%></td></tr>
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
			<%=EasyCrm.pagelist("Notice.asp?action=Main&otype="&otype&"", PN,TotalPages,TotalService)%>
		</td> 
	</tr>
</table>
</div>

<% End Sub

Sub setstar() 
  Dim ONId
    ONId = Trim(Request("ONId"))
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From OA_Notice Where ONId = " & ONId,conn,3,2
	If rs.RecordCount > 0 Then
		sql="UPDATE OA_Notice SET ONStar='1' Where ONId In (" & ONId & ")"
		conn.execute sql
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
    Response.Write("<script>location.href='Notice.asp' ;</script>")
End Sub

Sub setstarno() 
  Dim ONId
    ONId = Trim(Request("ONId"))
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From OA_Notice Where ONId = " & ONId,conn,3,2
	If rs.RecordCount > 0 Then
		sql="UPDATE OA_Notice SET ONStar='0' Where ONId In (" & ONId & ")"
		conn.execute sql
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
    Response.Write("<script>location.href='Notice.asp' ;</script>")
End Sub

Sub deleteData()
    Dim ONid
	ONid = Trim(Request("ONid"))
	If ONid = "" Then
	Exit Sub
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From OA_Notice Where ONid = " & ONid,conn,3,2
	If rs.RecordCount > 0 Then
	    ONid = rs("ONid")
		rs.Delete
		rs.Update
	End If
	rs.Close
	Set rs = Nothing	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
</table>
</body>
</html><%else%>无权限<%end if%>
<% Set EasyCrm = nothing %>