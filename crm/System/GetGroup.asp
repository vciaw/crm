<!--#include file="../Data/Conn.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>
<body>
<style>body{padding-bottom:55px;}</style>
<%
action = Trim(Request("action"))
sType = Trim(Request("sType"))
tipinfo = Trim(Request("tipinfo"))

Select Case action
Case "Setting"
    Call Setting()
Case "Products"
    Call Products()
Case "AreaData"
    Call AreaData()
Case "CustomField"
    Call CustomField()
Case "SelectData"
    Call SelectData()
Case "User"
    Call User()
Case "Group"
    Call Group()
Case "Level"
    Call Level()
Case "InfoList"
    Call InfoList()
End Select


Sub Group()
	if tipinfo<>"" then
		Response.Write("<script>art.dialog({title: 'Error',time: 1.5,icon: 'warning',content: '"&tipinfo&"'});</script>")
	end if
%>

	<script language="JavaScript">
	<!-- ��������ʾ
	function CheckInput()
	{
		if(document.getElementById('gId').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '���ű�Ų���Ϊ�գ�'});document.getElementById('gId').focus();return false;}
		if(document.getElementById('gName').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '�������Ʋ���Ϊ�գ�'});document.getElementById('gName').focus();return false;}
	}
	-->
	</script>
<%
if sType="Add" then
%>
		<form name="Save" action="GetGroup.asp?action=Group&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>�������� </B></td>
							</tr>
							<tr> 
								<td class="td_l_r title">���ű��</td>
								<td class="td_l_l"><input name="gId" type="text" class="int" id="gId" size="10" maxlength="2" onkeyup='this.value=this.value.replace(/\D/gi,"")' >  <span class="info_help help01">�ޣ����� 1 - 99</span></td>
							</tr>
							<tr> 
								<td class="td_l_r title">��������</td>
								<td class="td_l_l"><input name="gName" type="text" class="int" id="gName" size="30" maxlength="16" > </td>
							</tr>
						</table>
					</td>
				</tr>
				
				<tr>
					<td valign="top" class="td_n pdl10 pdr10"> 
						<div style="float:left;padding:10px 0 0;">
							<input type="submit" name="Submit" class="button45" value="����">��
							<input name="Back" type="button" id="Back" class="button43" value="�ر�" onClick="art.dialog.close();">
						</div>
					</td> 
				</tr>
			</table>
		</form>
<%
elseif sType="SaveAdd" then
	gId = Trim(Request("gId"))
	gName = Trim(Request("gName"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [system_group] Where gId = " & gId & " Or gName = '" & gName & "' ",conn,3,1
	If rs.RecordCount > 0 Then
		Response.Write("<script>location.href='GetGroup.asp?action=Group&sType=Add&tipinfo=���ű�Ż����������ظ�';</script>")
	Response.End()
	End If
	rs.Close
	rs.Open "Select Top 1 * From [system_group]",conn,3,2
	rs.AddNew
	rs("gId") = gId
	rs("gName") = gName
	rs.Update
	rs.Close
	Set rs = Nothing
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
elseif sType="Edit" then
	gId = Request("Id")
%>
		<form name="Save" action="GetGroup.asp?action=Group&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>�޸Ĳ��� </B></td>
							</tr>
							<tr> 
								<td class="td_l_r title">���ű��</td>
								<td class="td_l_l"><input name="gId" type="text" class="int" id="gId" size="10" maxlength="2" value="<%=gId%>" onkeyup='this.value=this.value.replace(/\D/gi,"")' >  <span class="info_help help01">�ޣ����� 1 - 99</span></td>
							</tr>
							<tr> 
								<td class="td_l_r title">��������</td>
								<td class="td_l_l"><input name="gName" type="text" class="int" id="gName" size="30" maxlength="16" value="<%=EasyCrm.getNewItem("system_group","gID",""&gID&"","gName")%>" > </td>
							</tr>
						</table>
					</td>
				</tr>
				
				<tr>
					<td valign="top" class="td_n pdl10 pdr10"> 
						<div style="float:left;padding:10px 0 0;">
							<input name="gIdOld" type="hidden" id="gIdOld" value="<%=gID%>">
							<input name="gNameOld" type="hidden" id="gNameOld" value="<%=EasyCrm.getNewItem("system_group","gID",""&gID&"","gName")%>">
							<input type="submit" name="Submit" class="button45" value="����">��
							<input name="Back" type="button" id="Back" class="button43" value="�ر�" onClick="art.dialog.close();">
						</div>
					</td> 
				</tr>
			</table>
		</form>
<%
elseif sType="SaveEdit" then

	gId = Trim(Request("gId"))
	gIdOld = Trim(Request("gIdOld"))
	gName = Trim(Request("gName"))
	gNameOld = Trim(Request("gNameOld"))
	
	
	if gId = gIdOld then '���û���²��ű��
		if gName <> gNameOld then
			'���ֻ�޸Ĳ������ƣ��ж��Ƿ����������������ظ�
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open "Select * From [system_group] Where gName = '" & gName & "' ",conn,1,1
			If rs.RecordCount > 0 Then
				Response.Write("<script>location.href='GetGroup.asp?action=Group&sType=Edit&Id="&gId&"&tipinfo=�����������ظ�';</script>")
			Response.End()
			else
			conn.execute("update [system_group] set gName = '"&gName&"' where gName = '"&gNameOld&"' ")
			End If
			rs.Close
		end if
	else '��������˲��ű�ţ�ͬ�������û���Ϳͻ���
	
		'������ű�������������ظ�
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [system_group] Where gId = " & gId & " and gName <> '" & gNameOld & "' ",conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetGroup.asp?action=Group&sType=Edit&Id="&gIdOld&"&tipinfo=���ű�����ظ�';</script>")
		Response.End()
		End If
		rs.Close
		
		if gName <> gNameOld then 
			'��������˲������ƣ����жϲ��������Ƿ����Ĳ����ظ�
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open "Select * From [system_group] Where gName = '" & gName & "' and gId="&gId&" ",conn,1,1
			If rs.RecordCount > 0 Then
				Response.Write("<script>location.href='GetGroup.asp?action=Group&sType=Edit&Id="&gId&"&tipinfo=�����������ظ�';</script>")
			Response.End()
			else
			conn.execute("update [system_group] set gId = '"&gId&"',gName='"&gName&"' where gId = "&gIdOld&" ")
			End If
			rs.Close
		else '���ֻ�޸Ĳ��ű�ţ��򲻿��ǲ�������
			conn.execute("update [system_group] set gId = '"&gId&"' where gId = "&gIdOld&" ")
		end if
			conn.execute("update [user] set uGroup = '"&gId&"' where uGroup = "&gIdOld&" ")
			conn.execute("update [client] set cGroup = '"&gId&"' where cGroup = "&gIdOld&" ")
	end if
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	
end if
%>
<%
End Sub
%>

<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>