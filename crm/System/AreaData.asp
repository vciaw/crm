<!--#include file="../data/conn.asp" --><% If mid(Session("CRM_qx"), 5, 1) = 1 Then %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>
<body style="padding-top:35px;">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">��ǰλ�ã�ϵͳ���� > ��������</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="ˢ��" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="����" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="ǰ��" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
			<div class="MenuboxS">
				 <ul>
					<li class="hover"><span><a href="?">�����б�</a></span></li>
					<li class=""><span><a href="#" onclick='AreaData_BClass_Add()' style="cursor:pointer">��������</a></span></li>
					<li class=""><span><a href="#" onclick='AreaData_Import()' style="cursor:pointer">��������</a></span></li>
				</ul>
			</div>
		</td>
	</tr>
</table>
<script>function AreaData_BClass_Add() {$.dialog.open('GetAreaData.asp?action=AreaData&sType=BigClassAdd', {title: '������������', width: 400,height: 145, fixed: true}); };</script>
<script>function AreaData_Import() {$.dialog.open('GetAreaData.asp?action=AreaData&sType=Import', {title: '����ȫ������', width: 600,height: 350, fixed: true}); };</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pd10">
<%
action = Trim(Request.QueryString("action"))
tipinfo = Trim(Request("tipinfo"))
Select Case action
Case "delete" 'ɾ����������
    Call deleteData()
Case Else
    Call main()
End Select

Sub main()  'Ĭ����ʾ���������б�
if tipinfo<>"" then
	Response.Write("<script>art.dialog({title: '��ʾ',time: 1,icon: 'warning',content: '"&tipinfo&"'});</script>")
end if
%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
				  <td class="td_l_l">��Ϣ�б�</td>
				  <td class="td_l_c" width="120">����</td>
				</tr>
								<%
								Set rs = Server.CreateObject("ADODB.Recordset")
								rs.Open "Select * From [AreaData] where aFId = '0' order by aId asc ",conn,3,1
								If rs.RecordCount > 0 Then
								Do While Not rs.BOF And Not rs.EOF
								%>
								<tr class="tr">
									<td class="tr_f"><a href="#" onclick='AreaData_BClass_Edit<%=rs("aId")%>()' title='�޸�' style="cursor:pointer"><%=rs("aName")%></a></td>
									<td class="td_l_r title"><input type="button" class="button_info_add" value=" " title="���С��"  onclick='AreaData_SClass_Add<%=rs("aId")%>()' style="cursor:pointer" /><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='AreaData_BClass_Edit<%=rs("aId")%>()' style="cursor:pointer" /><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onClick=window.location.href="?action=delete&aId=<%=rs("aId")%>" /></td>
								</tr>
						<script>function AreaData_BClass_Edit<%=rs("aId")%>() {$.dialog.open('GetAreaData.asp?action=AreaData&sType=BigClassEdit&aId=<%=rs("aId")%>', {title: '�༭��������', width: 400,height: 145, fixed: true}); };</script>
						<script>function AreaData_SClass_Add<%=rs("aId")%>() {$.dialog.open('GetAreaData.asp?action=AreaData&sType=SmallClassAdd&aFId=<%=rs("aId")%>', {title: '��ӵ���С��', width: 400,height: 180, fixed: true}); };</script>
								<%	'�ӷ����б�
										Set rss = Server.CreateObject("ADODB.Recordset")
										rss.Open "Select * From [AreaData] where aFId ='" & rs("aId") & "' ",conn,3,1
										Do While Not rss.BOF And Not rss.EOF
								%>
										<tr class="tr">
											<td class="td_l_l" style="padding-left:30px;">������ <a onclick='AreaData_SClass_Edit<%=rss("aId")%>()' title='�޸�' style="cursor:pointer"><%=rss("aName")%></a></td>
											<td class="td_l_r"><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>" onclick='AreaData_SClass_Edit<%=rss("aId")%>()' style="cursor:pointer" /><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onClick=window.location.href="?action=delete&aId=<%=rss("aId")%>" /></td>
										</tr>
						<script>function AreaData_SClass_Edit<%=rss("aId")%>() {$.dialog.open('GetAreaData.asp?action=AreaData&sType=SmallClassEdit&aId=<%=rss("aId")%>', {title: '�༭����С��', width: 400,height: 180, fixed: true}); };</script>
								<%
											rss.MoveNext
										Loop
										rss.Close
										Set rss = Nothing
										
									rs.MoveNext
								Loop
								else
								%>
								<tr class="tr">
									<td class="tr_f" colspan=2>�����ݣ�</td>
								</tr>
								<%
								end if
								rs.Close
								Set rs = Nothing
								%>
			</table>
<%
End Sub


Sub deleteData() 'ɾ��

	aId = Request("aId")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [AreaData] Where aFId = '"&aId&"'",conn,1,1 '�жϵ�ǰ�������Ƿ�����ӷ���
	If rs.RecordCount > 0 Then
        Response.Write("<script>location.href='?action=main&tipinfo=���ӷ��࣬��ֹɾ����';</script>")
	else
		Set rss = Server.CreateObject("ADODB.Recordset")
		rss.Open "Select * From [AreaData] Where aId = " & aId,conn,3,2
		If rss.RecordCount > 0 Then
			rss.Delete
			rss.Update
		End If
		rss.Close
		Set rss = Nothing
		Response.Redirect("?action=main")
	end if
	rs.Close
	Set rs = Nothing
End Sub
%>
		</td>
	</tr>
</table>
</body>
</html>
<%
else
Response.write"<script>alert("""&alert31&""");location.href=""../"";</script>"
end if
%>