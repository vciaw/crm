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
		<td class="top_left td_t_n td_r_n">��ǰλ�ã�ϵͳ���� > �Զ����ֶ�</td>
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
					<li class="hover"><span><a href="?">��Ϣ�б�</a></span></li>
					<li class=""><span><a href="#" onclick='CustomField_Add()' style="cursor:pointer">�����ֶ�</a></span></li>
				</ul>
			</div>
		</td>
	</tr>
</table>
<script>function CustomField_Add() {$.dialog.open('GetCustomField.asp?action=CustomField&sType=Add', {title: '�����ֶ�', width: 400,height: 420, fixed: true}); };</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pd10">
<%
action = Trim(Request.QueryString("action"))
tipinfo = Trim(Request("tipinfo"))
Select Case action
Case "delete" 'ɾ��
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
				  <td class="td_l_c" width="120">���ݱ�</td>
				  <td class="td_l_c" width="120">��ʾ��</td>
				  <td class="td_l_c" width="120">�ֶ���</td>
				  <td class="td_l_c" width="120">�ֶ�����</td>
				  <td class="td_l_c" width="120">������</td>
				  <td class="td_l_l">Ĭ��ֵ(��Ƕ��ŷָ�)</td>
				  <td class="td_l_c" width="100">�Ƿ�����ʾ</td>
				  <td class="td_l_c" width="100">�Ƿ�����</td>
				  <td class="td_l_c" width="100">����</td>
				</tr>
								<%
								Set rs = Server.CreateObject("ADODB.Recordset")
								rs.Open "Select * From [CustomField] order by Id asc ",conn,3,1
								If rs.RecordCount > 0 Then
								Do While Not rs.BOF And Not rs.EOF
								%>
								<tr class="tr">
									<td class="td_l_c">[<%=rs("cTable")%>]</td>
									<td class="td_l_c"><%=rs("cTitle")%></td>
									<td class="td_l_c"><%=rs("cName")%></td>
									<td class="td_l_c"><%=rs("cType")%></td>
									<td class="td_l_c"><%=rs("cWidth")%>PX</td>
									<td class="td_l_l"><%=rs("cContent")%></td>
									<td class="td_l_c"><%if rs("cList")= 1 then%>��<%else%>��<%end if%></td>
									<td class="td_l_c"><%if rs("cYn")= 1 then%>��<%else%>��<%end if%></td>
									<td class="td_l_c"><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='CustomField_Edit<%=rs("Id")%>()' style="cursor:pointer" /><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onClick=window.location.href="?action=delete&Id=<%=rs("Id")%>" /></td>
								</tr>
								<script>function CustomField_Edit<%=rs("Id")%>() {$.dialog.open('GetCustomField.asp?action=CustomField&sType=Edit&Id=<%=rs("Id")%>', {title: '�༭', width: 400,height: 420, fixed: true}); };</script>
								<%
									rs.MoveNext
								Loop
								else
								%>
								<tr class="tr">
									<td class="tr_f" colspan=7>�����ݣ�</td>
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
	Id = Request("Id")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [CustomField] Where Id = " & Id,conn,3,2
	If rs.RecordCount > 0 Then
		rs.Delete
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
	Response.Redirect("?action=main")
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