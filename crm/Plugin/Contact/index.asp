<!--#include file="../../data/conn.asp" --><!--#include file="config.asp" --><!--#include file="../../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "index.asp"
Session("CRM_pagenum") = PNN
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<link href="<%=SiteUrl&skinurl%>chosen/chosen.css" rel="stylesheet" />
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Float.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>

<body>

<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：功能插件 > 通讯录</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<%if inStr(Plugin_contact_manage,session("CRM_name"))>0 or Session("CRM_level") = 9 then%>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
            <div class="MenuboxS">
              <ul>
					<li <%if otype="Main" or otype="" then%>class="hover"<%end if%> id="CheckB"><span><a href="?action=Main&otype=Main">信息列表</a></span></li>
					<li class="" id="CheckC"><span><a href="#" onclick='Plugin_Contact_InfoAdd()' style="cursor:pointer">新增信息</a></span></li>
					<% if inStr(Plugin_Contact_manage,session("CRM_name"))>0 or Session("CRM_level") = 9 then %>
					<li <%if otype="Manage" then%>class="hover"<%end if%>><span><a href="?action=Manage&otype=Manage">高级管理</a></span></li>
					<%end if%>
              </ul>
            </div>
		</td>
	</tr>
</table>
<script>function Plugin_Contact_InfoAdd() {$.dialog.open('GetUpdate.asp?action=Add', {title: '新增', width: 700, height: 420,fixed: true}); };</script>

<%
action = Trim(Request("action"))
Select Case action
Case "Manage"
    Call infoManage()
Case "Managesave"
    Call infoManagesave()
Case "delete"
    Call infodelete()
Case Else
    Call Main()
End Select
%>

<%
Sub Main()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr> 
		<td valign="top" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" colspan=2 style="padding:10px 10px 0;" class="td_n">
						
						<form name="Search" action="?action=CheckSub&SubAction=Search&PN=<%=PNN%>" method="post">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1" style="border-bottom:0px;">
							<tr class="tr_t"style="border-bottom:0px;"> 
								<td class="td_l_l">
								<B>信息列表</B>
								</td>
							</tr>
						</table> 
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_b">
								<td width="80" class="td_l_c">编号</td>
								<td width="100" class="td_l_c">分类</td>
								<td class="td_l_l">公司名称</td>
								<td width="100" class="td_l_c">联系人</td>
								<td width="100" class="td_l_c">电话</td>
								<td width="100" class="td_l_c">职位</td>
								<td width="100" class="td_l_c">部门</td>
								<td width="100" class="td_l_c">ＱＱ</td>
								<td width="100" class="td_l_c">主营</td>
								<td width="90" class="td_l_c"><%=L_Top_Manage%></td>
							</tr>
							<%
						    PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [Plugin_Contact] where 1 = 1 "&sql&" Order By id Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Plugin_Contact] where 1 = 1 "&sql&" and id < ( SELECT Min(id) FROM ( SELECT TOP "&pagenum&" id FROM [Plugin_Contact]  where 1 = 1 "&sql&" ORDER BY id desc ) AS T ) Order By id Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(id) As RecordSum From [Plugin_Contact] where 1 = 1 "&sql&" ",1,1)
						
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
								<td class="td_l_c"><%=rs("Id")%></td>
								<td class="td_l_c"><%=rs("cClass")%></td>
								<td class="td_l_l"><%=rs("cCompany")%></td>
								<td class="td_l_c"><%=rs("cLinkman")%></td>
								<td class="td_l_c"><%=rs("cTel")%></td>
								<td class="td_l_c"><%=rs("cZhiwei")%></td>
								<td class="td_l_c"><%=rs("cGroup")%></td>
								<td class="td_l_c"><%=rs("cQQ")%></td>
								<td class="td_l_c"><%=rs("cProducts")%></td>
								<td class="td_l_c">
									<input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Plugin_Contact_InfoEdit<%=rs("Id")%>()' style="cursor:pointer" /> 
									<%if Session("CRM_level") = 9 then%><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Plugin_Contact_InfoDel<%=rs("Id")%>()' style="cursor:pointer" /><%end if%>
								</td>
							</tr>
							<script>function Plugin_Contact_InfoEdit<%=rs("Id")%>() {$.dialog.open('GetUpdate.asp?action=InfoEdit&Id=<%=rs("Id")%>', {title: '编辑', width: 700,height: 420, fixed: true}); };</script>
							<script>function Plugin_Contact_InfoDel<%=rs("Id")%>()
							{
								art.dialog({
									content: '<%=Alert_del_YN%>',
									icon: 'error',
									ok: function () {
										art.dialog.open('?action=delete&Id=<%=rs("Id")%>');
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
							<tr><td class="td_l_l" colspan="10"><%=L_Notfound%></td></tr>
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

<%
end sub

Sub infoManage()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr> 
		<td valign="top" class="td_n">
		
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" colspan=2 style="padding:10px 10px 0;" class="td_n">
						
						<form name="Managesave" action="?action=Managesave" method="post" onSubmit="return CheckInput();">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1" style="border-bottom:0px;">
							<tr class="tr_t"style="border-bottom:0px;"> 
								<td class="td_l_l">
								<B>高级配置</B>
								</td>
							</tr>
						</table> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr">
								<td class="td_l_c title" style="border-top:0;">分类</td>
								<td class="td_r_l" style="border-top:0;">
									<input name="Plugin_contact_class" type="text" class="int" id="Plugin_contact_class" size="40" value="<%=Plugin_contact_class%>"> <span class="info_help help01">不同分类之间用半角逗号分割，结尾不含逗号，下同。</span>
								</td>
							</tr>
							<tr class="tr">
								<td class="td_l_c title">职位</td>
								<td class="td_r_l">
									<input name="Plugin_contact_zhiwei" type="text" class="int" id="Plugin_contact_zhiwei" size="40" value="<%=Plugin_contact_zhiwei%>">
								</td>
							</tr>
							<tr class="tr">
								<td class="td_l_c title">部门</td>
								<td class="td_r_l">
									<input name="Plugin_contact_group" type="text" class="int" id="Plugin_contact_group" size="40" value="<%=Plugin_contact_group%>">
								</td>
							</tr>
							<tr >
								<td class="td_l_c title">管理权限</td>
								<td class="td_r_l" style="padding:10px;">
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
									<col width="100" />
									<%
										Set rsg = Server.CreateObject("ADODB.Recordset")
										rsg.Open "Select * From [system_group]",conn,1,1
										Do While Not rsg.BOF And Not rsg.EOF
									%>
										<tr> 
											<td class="td_l_c title"><%=rsg("gName")%></td>
											<td  class="td_l_l">
											<%
												Set rsm = Server.CreateObject("ADODB.Recordset")
												rsm.Open "Select * From [user] where uGroup="&rsg("gId")&" ",conn,1,1
												Do While Not rsm.BOF And Not rsm.EOF
											%>
											<input type="checkbox" name="Plugin_contact_manage" value="<%=rsm("uName")%>" <%if inStr(Plugin_contact_manage,rsm("uName"))>0 then%>checked<%end if%>> <%=rsm("uName")%>　
											<%
												rsm.MoveNext
												Loop
												rsm.Close
												Set rsm = Nothing
											%>
											</td>
										</tr> 
									<%
										rsg.MoveNext
										Loop
										rsg.Close
										Set rsg = Nothing
									%>
									</table>
								</td>
							</tr>
							<tr>
								<td class="td_r_l" colspan="4">
								<input type="submit" name="Submit" class="button45" value=" <%=L_Edit%> ">　
								<input name="Back" type="button" id="Back" class="button43" value=" <%=L_Back%> " onClick="history.back();">
								</td>
							</tr>
						</table>   
						</form>
					</td>
				</tr>
			</table>
        </td>
	</tr>
</table>

<%
End Sub

Sub infoManagesave()
	Plugin_contact_class = replace(Trim(Request.Form("Plugin_contact_class")),CHR(34),"'")
	Plugin_contact_zhiwei = replace(Trim(Request.Form("Plugin_contact_zhiwei")),CHR(34),"'")
	Plugin_contact_group = replace(Trim(Request.Form("Plugin_contact_group")),CHR(34),"'")
	Plugin_contact_manage = replace(Trim(Request.Form("Plugin_contact_manage")),CHR(34),"'")
	Dim TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "Dim Plugin_contact_class,Plugin_contact_zhiwei,Plugin_contact_group,Plugin_contact_manage" & VbCrLf
	
	TempStr = TempStr & "'通讯录配置" & VbCrLf
	TempStr = TempStr & "Plugin_contact_class="& Chr(34) & Plugin_contact_class & Chr(34) &" '分类" & VbCrLf
	TempStr = TempStr & "Plugin_contact_zhiwei="& Chr(34) & Plugin_contact_zhiwei & Chr(34) &" '职位" & VbCrLf
	TempStr = TempStr & "Plugin_contact_group="& Chr(34) & Plugin_contact_group & Chr(34) &" '部门" & VbCrLf
	TempStr = TempStr & "Plugin_contact_manage="& Chr(34) & Plugin_contact_manage & Chr(34) &" '权限" & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf
	ADODB_SaveToFile TempStr,"Config.asp"
	Response.Write "<script>location.href='?action=List&otype=Main';</script>"
End Sub

Sub infodelete()
    Dim Id
	Id = CLng(ABS(Request("Id")))
	PN = Request("PN")
	If Not IsNumeric(Id) Or Id <= 0 Then Response.Write "<script>alert(""不存在"");</script>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Plugin_Contact] Where Id = " & Id,conn,3,2
	If rs.RecordCount <> 1 Then Response.Write "<script>alert(""不存在"");</script>"
	Id = rs("Id")
	rs.Delete
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("index.asp?PN="&PN&"")
End Sub

Sub ADODB_SaveToFile(ByVal strBody,ByVal File)
	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If
	
	If FSFlag = 1 Then
		Set WriteFile = fs.CreateTextFile(Server.MapPath(File),True)
		WriteFile.Write strBody
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请使用FTP等功能，将<font color=Red >data/config.asp</font>文件内容替换成框中内容"
			Err.Clear
			Set objStream = Noting
			Exit Sub
		End If
		With objStream
			.Type = 2
			.Open
			.Charset = "GB2312"
			.Position = objStream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	End If
End Sub

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>
		</td>
	</tr>
</table>
<script src="../../data/calendar/WdatePicker.js"></script>
<%else%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10">   
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>错误提示</B></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr>
					<td class="td_r_l" style="border-top:0;">
						您无权使用该插件！
					</td>
				</tr>
			</table>   
		</td>
	</tr>
</table>

<%end if%>
</body>
</html><% Set EasyCrm = nothing %>