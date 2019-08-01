<!--#include file="../data/conn.asp"--><!--#include file="../UpLoad/UpLoad.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
Action = Trim(Request("Action"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
If subAction = "searchItem" Then
    Dim sClass,sTitle,sShare,sUser,TimeBegin,TimeEnd
	sClass = Trim(Request("sclass"))
	sTitle = Trim(Request("stitle"))
	sShare = Trim(Request("sShare"))
	sUser = Trim(Request("User"))
	TimeBegin = Trim(Request("TimeBegin"))
	TimeEnd = Trim(Request("TimeEnd"))
	Session("Search_Soft_sClass") = Trim(Request("sclass"))
	Session("Search_Soft_sTitle") = Trim(Request("stitle"))
	Session("Search_Soft_sShare") = Trim(Request("sShare"))
	Session("Search_Soft_sUser") = Trim(Request("User"))
	Session("Search_Soft_TimeBegin") = Trim(Request("TimeBegin"))
	Session("Search_Soft_TimeEnd") = Trim(Request("TimeEnd"))
	
	Dim sql
    sql = ""	
	
    If sClass <> "" Then
		sql = sql & " And s_class = '" & sClass & "'"
	End If
	
    If sTitle <> "" Then
		sql = sql & " And s_title Like '%" & sTitle & "%'"
	End If
	
    If sUser <> "" Then
		sql = sql & " And s_user = '" & sUser & "'"
	End If
			
	if Accsql =1 then
	If TimeBegin <> "" and  TimeEnd <> "" Then
	    sql = sql & " And s_time >= '" & TimeBegin & "' And s_time <= '" & TimeEnd & "' "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And DATEDIFF(d,s_time,'"&TimeBegin&"')=0 "
	End If
	else
	If TimeBegin <> "" and TimeEnd <> "" Then
	    sql = sql & " And s_time >= #" & TimeBegin & "# And s_time <= #" & TimeEnd & "# "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And DateDiff('d', s_time, '"&TimeBegin&"') =0"
	End If
	end if
			
	If sShare <> "" Then
	    sql = sql & " And s_share = " & sShare & " "
	End If

	If Session("CRM_level") < 9 Then
		sql = sql & " And s_user In (" & arrUser & ")"
	End If
	
End If

If sClass = "" And sTitle = "" And sUser = "" And TimeBegin = "" And TimeEnd = "" And sShare = "" Then
    If Session("CRM_Search_Soft") <> "" Then
        sql = Session("CRM_Search_Soft")
	Else
	    If Session("CRM_level") < 9 Then
		sql = " And (s_share = 1 or s_user In (" & arrUser & "))"
		End If
	End If
Else
    Session("CRM_Search_Soft") = sql
End If

If subAction = "killSession" Then
	Session("CRM_Search_Soft") = ""
	Session("Search_Soft_sClass") = ""
	Session("Search_Soft_sTitle") = ""
	Session("Search_Soft_sShare") = ""
	Session("Search_Soft_sUser") = ""
	Session("Search_Soft_TimeBegin") = ""
	Session("Search_Soft_TimeEnd") = ""
	If Session("CRM_level") < 9 Then
		sql = " And (s_share = 1 or s_user In (" & arrUser & "))"
	else
		sql=""
	end if
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Float.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
<script language="JavaScript">
<!--
function CheckAll(form) {
for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
}
-->
</script>
</head>

<body style="padding-top:35px;"> 
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：办公OA > 文件柜 </td>
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
					<li <%if otype="Main" or otype="" then%>class="hover"<%end if%> id="CheckB"><span><a href="?otype=Main">文件列表</a></span></li>
					<li class="" id="CheckA"><span><a href="javascript:void(0)" style="cursor:pointer">高级搜索</a></span></li>
					<li class="" id="CheckC"><span><a href="#" onclick='Soft_InfoAdd()' style="cursor:pointer">上传文件</a></span></li>
				</ul>
			</div>
		</td>
	</tr>
</table>

<script>function Soft_InfoAdd() {$.dialog.open('../OA/GetUpdate.asp?action=Soft&sType=Add', {title: '新增', width: 500, height: 280,fixed: true}); };</script>

<%
	userfolder = Session("CRM_account")
	filefolder = Server.MapPath("../soft/"&userfolder)
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(filefolder) then '如果文件夹不存在则创建
	fso.CreateFolder(filefolder) 
	end if
	
Select Case action
	Case "CheckSub"		'批量操作
    Call CheckSubject()
	Case "YNShare"
		Call YNShare()
	Case "RealDel"
		Call RealDel()
	Case "checkalldel"
		Call checkalldel()
	Case Else
		Call Main()
End Select

Sub Main() 
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr> 
		<td valign="top" class="td_n">
		
			<div id="SearchBox" style="position: absolute; width:100%; height:450px; background:#ffffff; display:none; z-index:10;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0" style="background:#ffffff;">
					<tr>
						<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
							<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
								<tr class="tr_t"> 
									<td class="td_l_l" COLSPAN="6" style="border-right:0;"><B><%=L_Top_Search%></B></td>
								</tr>
							</table>
							<form name="searchForm" action="?subAction=searchItem" method="post">
							<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
								<col width="120" />
								<tr>
									<td class="td_l_r title" style="border-top:0;">文件名</td>
									<td class="td_r_l" style="border-top:0;"><input name="stitle" type="text" class="int" id="stitle" size="30" value="<%=Session("Search_Soft_sTitle")%>" ></td>
								</tr>
								<tr>
									<td class="td_l_r title">创建时间</td>
									<td class="td_r_l"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Soft_TimeBegin")%>" /> ~ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Soft_TimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title">文件分类</td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_SoftClass","sclass",""&Session("Search_Soft_sclass")&"") %></td>
								</tr>
								<tr>
									<td class="td_l_r title">是否共享</td>
									<td class="td_r_l"> <select name="sShare"  class="int"><option value=""><%=L_Select%></option><option value="1"><%=L_Shi%></option><option value="0"><%=L_Fou%></option></select></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cUser%></td>
									<td class="td_r_l">
										<% If Session("CRM_level") = 9 Then %>
										<% = EasyCrm.UserList(2,"User",Session("Search_Soft_sUser")) %>
										<% Else %>
										<% = EasyCrm.UserList(1,"User",Session("Search_Soft_sUser")) %>
										<% End If %>
									</td>
								</tr>
								<tr>
									<td class="td_r_l" colspan="4">
										<input type="submit" name="Submit" class="button42" value="<%=L_Search%>">
										<input type="button" name="button" class="button43" value="<%=L_Clear%>" onClick=window.location.href="?SubAction=killSession" />
									</td>
								</tr>
							</table>   
							</form>
						</td> 
					</tr>
				</table>
			</div>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" colspan=2 style="padding:10px 10px 0;" class="td_n">
						
						<form name="Search" action="?action=CheckSub&SubAction=Search&PN=<%=PNN%>" method="post">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1" style="border-bottom:0px;">
							<tr class="tr_t"style="border-bottom:0px;"> 
								<td class="td_l_l">
									<span class="tips01" style="float:left;display:none;padding:0 10px;height:34px;line-height:34px;text-align:left;position:fixed;right:10px;top:43px;color:#000;" id="CheckSub">
										<input type="submit" name="Checkexecute" id="Checkexecute" class="button242" value="批量共享">
										<input type="submit" name="Checkexecute" id="Checkexecute" class="button243" value="撤销共享">
										<input type="submit" name="Checkexecute" id="Checkexecute" class="button247" value="批量删除">
									</span>
								<B>信息列表</B>
								</td>
							</tr>
						</table> 
						<script type="text/javascript">
						 function getBlock(ck, d) {
							var c = document.getElementById(ck);
							var d = document.getElementById(d);
							if (c.checked == true) {
								d.style.display = "block"
							} 
							//else {d.style.display = "none"}
						}
						</script>

						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_b">
								<td width="40" class="td_l_c"><input name="chkall" type="checkbox" id="chkall" onclick="CheckAll(this.form);getBlock('chkall','CheckSub')" value="checkbox"></td>
								<td width="60" class="td_l_c">编号</td>
								<td width="100" class="td_l_c">分类</td>
								<td class="td_l_l">文件名</td>
								<td width="80" class="td_l_c">下载</td>
								<td width="80" class="td_l_c">共享</td>
								<td width="80" class="td_l_c">发布人</td>
								<td width="150" class="td_l_c">时间</td>
								<td width="100" class="td_l_c">管理</td>
							</tr>
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [OA_soft] where 1 = 1 "&sql&" Order By sId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [OA_soft] where 1 = 1 "&sql&" and sId < ( SELECT Min(sId) FROM ( SELECT TOP "&pagenum&" sId FROM [OA_soft]  where 1 = 1 "&sql&" ORDER BY sId desc ) AS T ) Order By sId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(sId) As RecordSum From [OA_soft] where 1 = 1 "&sql&" ",1,1)
						
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
								<td class="td_l_c"><input type="checkbox" name="sId" id="sId<%=rs("sId")%>" value="<%=rs("sId")%>" onclick="getBlock('sId<%=rs("sId")%>','CheckSub')"></td>
								<td class="td_l_c"><%=rs("sId")%></td>
								<td class="td_l_c"><%=rs("s_class")%></td>
								<td class="td_l_l"><%=rs("s_title")%></td>
								<%if rs("s_file")<>"" then%>
								<td class="td_l_c"><a href="<%=rs("s_file")%>">下载</a></td>
								<%else%>
								<td class="td_l_c">无</td>
								<%end if%>
								<%If Session("CRM_level")=9 or rs("s_user")=Session("CRM_name") Then%>
								<%if rs("s_share")=0 then%>
								<td class="td_l_c"><a href="?action=YNShare&sid=<%=rs("sid")%>&s_share=1&PN=<%=PNN%>" style="color:#f00;"><img src="<%=SiteUrl&skinurl%>images/ico/no.gif" border=0></a></td>
								<%else%>
								<td class="td_l_c"><a href="?action=YNShare&sid=<%=rs("sid")%>&s_share=0&PN=<%=PNN%>" style="color:#f00;"><img src="<%=SiteUrl&skinurl%>images/ico/yes.gif" border=0></a></td>
								<%end if%>
								<%else%>
								<%if rs("s_share")=0 then%>
								<td class="td_l_c"><img src="<%=SiteUrl&skinurl%>images/ico/no.gif" border=0></td>
								<%else%>
								<td class="td_l_c"><img src="<%=SiteUrl&skinurl%>images/ico/yes.gif" border=0></td>
								<%end if%>
								<%end if%>
								<td class="td_l_c"><%=rs("s_user")%></td>
								<td class="td_l_c"><%=rs("s_time")%></td>
								<td class="td_l_c">
									<input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>" onclick='Soft_InfoEdit<%=rs("sId")%>()' style="cursor:pointer" />
									<input type="button" class="button_info_del" value=" " title="<%=L_RealDel%>" onclick='Soft_InfoDel<%=rs("sId")%>()' style="cursor:pointer" />
								</td>
							</tr>
							<script>function Soft_InfoEdit<%=rs("sId")%>() {$.dialog.open('../OA/GetUpdate.asp?action=Soft&sType=Edit&Id=<%=rs("sId")%>', {title: '查看', width: 500,height: 280, fixed: true}); };</script>
							<script>function Soft_InfoDel<%=rs("sId")%>() //彻底删除
							{
								art.dialog({
									content: '<%=Alert_del_YN%>',
									icon: 'warning',
									ok: function () {
										art.dialog.open('?action=RealDel&sId=<%=rs("sId")%>');
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
							<tr><td class="td_l_l" colspan="9"><%=L_Notfound%></td></tr>
							<%
							end if
							rs.Close
							Set rs = Nothing
							%>
							
						</table> 
						</form>
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
			<%if sql<>"" then%><span class="r"><input name="Back" type="button" id="Back" class="button227" value="清空" onClick=window.location.href="?SubAction=killSession"></span><%end if%>
			<%=EasyCrm.pagelist("index.asp", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>

<script language="JavaScript">
<!--
for(var i=0;i<document.getElementById('sShare').options.length;i++){
    if(document.getElementById('sShare').options[i].value == "<% = Session("Search_Soft_sShare") %>"){
    document.getElementById('sShare').options[i].selected = true;}}
-->
</script>
<% end sub 

Sub CheckSubject()
	id = Trim(Request("sId"))
	PN = CLng(ABS(Request("PN")))
If Request("Checkexecute")="批量共享" Then
	conn.execute "UPDATE OA_soft SET s_share='1' Where sid In (" & id & ")"
    Response.Write("<script>location.href='index.asp?PN="&PN&"' ;</script>")
elseIf Request("Checkexecute")="撤销共享" Then
	conn.execute "UPDATE OA_soft SET s_share='0' Where sid In (" & id & ")"
    Response.Write("<script>location.href='index.asp?PN="&PN&"' ;</script>")
elseIf Request("Checkexecute")="批量删除" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [OA_soft] Where sId in (" & id &") ",conn,3,2
		Do While Not rs.BOF And Not rs.EOF
			Dim rss 
			Set rss = Server.CreateObject("ADODB.Recordset")
			rss.Open "Select * From [OA_soft] Where sId in (" & id &") ",conn,3,2
			s_file=rss("s_file")
				If s_file <> "" Then '删除数据库信息同时删除文件
					Set fso = CreateObject("Scripting.FileSystemObject")
					IF fso.FileExists(server.MapPath(s_file)) Then
					fso.DeleteFile(server.MapPath(s_file))
					End IF
				End If
			rss.Delete
			rss.Update
		rs.MoveNext
		Loop
		rs.Close
	Set rs = Nothing
    Response.Write("<script>location.href='index.asp?PN="&PN&"' ;</script>")
end if
end sub 

Sub YNShare() '共享切换
  Dim sid,s_share,PN
    sid = Trim(Request("sid"))
    s_share = Trim(Request("s_share"))
    PN = Trim(Request("PN"))
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From OA_soft Where sid = " & sid,conn,3,2
	If rs.RecordCount > 0 Then
		sql="UPDATE OA_soft SET s_share='"&s_share&"' Where sid In (" & sid & ")"
		conn.execute sql
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
	Response.Write "<script>location.href='index.asp?PN="&PN&"';</script>"
End Sub

Sub RealDel() '彻底删除
  Dim sId
    sId = Trim(Request("sId"))
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From OA_soft Where sId = " & sId,conn,3,2
	file=rs("s_file")
		If file <> "" Then
			Set fso = CreateObject("Scripting.FileSystemObject")
			IF fso.FileExists(server.MapPath(file)) Then
			fso.DeleteFile(server.MapPath(file))
			End IF
		End If
	rs.Delete
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

%>
</table>
</body>
</html>
<% Set EasyCrm = nothing %>
<script src="../data/calendar/WdatePicker.js"></script>