<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
If Session("CRM_level") < 9 Then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
end if
	'��ȡgetֵ
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	id		=	Request("id")
%><%=Header%>
<!-- start header -->
    <div id="header">
         <a href="System.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
         	<a href="#" class="button list"><img src="img/create.png" width="16" height="16" alt="icon"/></a>
         <a onClick=window.location.href="javascript:window.location.reload();" class="button create"><img src="img/reload-button.png" width="15" height="16" alt="icon" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start page -->
    <div class="page">
	<div class="simplebox">
<%
Select Case action
Case "SaveAdd" '���
    Call SaveAdd()
Case "Edit" '�޸�
    Call Edit()
Case "SaveEdit" '�����޸�
    Call SaveEdit()
Case else
    Call Main()
End Select

Sub Main()
%>
	<script language="JavaScript">
	<!-- ��������ʾ
	function CheckInput()
	{
		if(document.getElementById('gId').value == ""){alert("���ű�Ų���Ϊ�գ�"); document.all.gId.focus();return false;}
		if(document.getElementById('gName').value == ""){alert("�������Ʋ���Ϊ�գ�"); document.all.gName.focus();return false;}
	}
	-->
	</script>
                <div class="content listbox" style="margin-bottom:10px;">
				<form name="Save" action="?action=SaveAdd" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> ���ű��</label>
						<input name="gId" type="number" class="int" id="gId" size="11" maxlength="16" min="1" max="99"> �ޣ����� 1 - 99
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> ��������</label>
						<input name="gName" type="text" class="int" id="gName" size="11" maxlength="16">
                    </div>
                    
                    <div class="form-line">
                    <input type="submit" name="button" id="button" value="&nbsp;�� ��&nbsp;" class="submit-button" />
                    </div>
                </form>
                </div>

            	<h1 class="titleh">�����б�</h1>
					<table class="tabledata"> 
						<col width="40"><col ><col width="70">
						<tbody> 
                        <tr> 
							<td >���</td> 
							<td>����</td> 
							<td>����</td> 
                        </tr> 
						<%
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.Open "Select * From [system_group] order by gId asc ",conn,1,1
							Do While Not rs.BOF And Not rs.EOF
						%>
								<tr>
									<td>[<%=rs("gId")%>]</td>
									<td><%=rs("gName")%></td>
									<td><input type="button" class="reset-button" value="<%=L_Edit%>" title="<%=L_Edit%>" onClick="window.location.href='?action=Edit&id=<%=rs("gId")%>'" /></td>
								</tr>
						<%
							rs.MoveNext
							Loop
							rs.Close
							Set rs = Nothing
						%>
                    </table> 
					<blockquote>�ֻ��治�ṩ����ɾ�����ܣ����������</blockquote>
			</div>
		<%=Footer%>
            
<%
end Sub

Sub SaveAdd() '���
	gId = Trim(Request("gId"))
	gName = Trim(Request("gName"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [system_group] Where gId = " & gId & " Or gName = '" & gName & "' ",conn,3,1
	If rs.RecordCount > 0 Then
		Response.Write("<script>alert('���ű�Ż����������ظ�');history.back(1);</script>")
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
	Response.Redirect("?")
End Sub

Sub Edit()
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From system_group where gid="&id,conn,1,1
	gId = rs("gId")
	gName = rs("gName")
	rs.Close
	Set rs = Nothing
%>
            	<h1 class="titleh">�޸Ĳ���</h1>
                <div class="content" style="margin-bottom:10px;">
				<form name="Save" action="?action=SaveEdit" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> ���ű��</label>
						<input name="gId" type="number" class="int" id="gId" size="11" maxlength="16" value="<%=gId%>" min="1" max="99"> �ޣ����� 1 - 99
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> ��������</label>
						<input name="gName" type="text" class="int" id="gName" size="11" maxlength="16" value="<%=gName%>">
                    </div>
                    
                    <div class="form-line">
					<input name="gIdOld" type="hidden" id="gIdOld" value="<%=gID%>">
					<input name="gNameOld" type="hidden" id="gNameOld" value="<%=gName%>">
                    <input type="submit" name="button" id="button" value="&nbsp;�� ��&nbsp;" class="submit-button" />
					<input name="Back" type="button" id="Back" class="reset-button" value="����" onClick="location.href='Group.asp';">
                    </div>
                </form>
                </div>

<%
End Sub
Sub SaveEdit()

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
				Response.Write("<script>alert('�����������ظ�');history.back(1);</script>")
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
				Response.Write("<script>alert('���ű�����ظ�');history.back(1);</script>")
		Response.End()
		End If
		rs.Close
		
		if gName <> gNameOld then 
			'��������˲������ƣ����жϲ��������Ƿ����Ĳ����ظ�
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open "Select * From [system_group] Where gName = '" & gName & "' and gId="&gId&" ",conn,1,1
			If rs.RecordCount > 0 Then
				Response.Write("<script>alert('�����������ظ�');history.back(1);</script>")
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
	
	Response.Redirect("?")
End Sub
%>
    <div class="clear"></div>
    </div>
    <!-- end page -->
<script type="text/javascript" src="js/frame.js"></script>
</body>
</html>
