<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
If Session("CRM_level") < 9 Then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
end if
	'��ȡgetֵ
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
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
	
	<script language="JavaScript">
	<!-- ��������ʾ
	function CheckInput()
	{
		if(document.getElementById('account').value == ""){alert("��¼�˺Ų���Ϊ�գ�"); document.all.account.focus();return false;}
		if(document.getElementById('name').value == ""){alert("��ʵ��������Ϊ�գ�"); document.all.name.focus();return false;}
		if(document.getElementById('password').value == ""){alert("��¼���벻��Ϊ�գ�"); document.all.password.focus();return false;}
		if(document.getElementById('level').value == ""){alert("ϵͳ��ɫ����Ϊ�գ�"); document.all.level.focus();return false;}
		if(document.getElementById('group').value == ""){alert("�������Ų���Ϊ�գ�"); document.all.group.focus();return false;}
	}
	-->
	</script>
                <div class="content listbox" style="margin-bottom:10px;">
				<form name="Save" action="?action=SaveAdd" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> ��¼�˺�</label>
						<input name="account" type="text" class="int" id="account" size="11" maxlength="16"> ¼��󲻿��޸�
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> ��ʵ����</label>
						<input name="name" type="text" class="int" id="name" size="11" maxlength="16">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> ��¼����</label>
						<input name="password" type="password" class="int" id="password" size="11" maxlength="16">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> ϵͳ��ɫ</label>
						<% = EasyCrm.getList(2,"system_level","lId","lName","level","") %>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> ��������</label>
						<% = EasyCrm.getList(2,"system_group","gId","gName","group","") %>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">����</label>
						<input name="Birthday" type="date" id="Birthday" class="int Wdate" size="15" onFocus="WdatePicker()">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">��˾ʱ��</label>
						<input name="addtime" type="date" id="addtime" class="int Wdate" size="15" onFocus="WdatePicker()">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">�ֻ�</label>
						<input name="Mobile" type="number" class="int" id="Mobile" size="20" maxlength="16">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">E-mail</label>
						<input name="Email" type="email" class="int" id="Email">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">סַ</label>
						<input name="Address" type="text" class="int" id="Address">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">���֤</label>
						<input name="card" type="number" class="int" id="card">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">���ͻ���</label>
						<input name="ClientNum" type="number" id="ClientNum" class="int" size="10"  value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <span class="info_help help01" >&nbsp;��Ϊ�����ƣ�</span>
                    </div>
                    
                    <div class="form-line">
                    <input type="submit" name="button" id="button" value="&nbsp;�� ��&nbsp;" class="submit-button" />
                    </div>
                </form>
                </div>

            	<h1 class="titleh">Ա���б�</h1>
					<table class="tabledata"> 
						<tbody> 
                        <tr> 
							<td >���</td> 
							<td>�ʺ�</td> 
							<td>����</td> 
							<td>����</td> 
							<td>��ɫ</td> 
                        </tr> 
						<%
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.Open "Select * From [user] order by uId asc ",conn,1,1
							Do While Not rs.BOF And Not rs.EOF
						%>
								<tr>
									<td>[<%=rs("uId")%>]</td>
									<td><%=rs("uAccount")%></td>
									<td><%=rs("uName")%></td>
									<td><%=EasyCrm.getNewItem("system_group","gId",rs("uGroup"),"gName")%></td>
									<td><%=EasyCrm.getNewItem("system_level","lId",rs("uLevel"),"lName")%></td>
								</tr>
						<%
							rs.MoveNext
							Loop
							rs.Close
							Set rs = Nothing
						%>
                    </table> 
					<blockquote>�ֻ��治�ṩԱ���༭��ɾ�����ܣ����������</blockquote>
			</div>
		<%=Footer%>
            
<%

Select Case action
Case "SaveAdd" '���
    Call SaveAdd()
End Select

Sub SaveAdd() '���
	uAccount = Trim(Request("account"))
	uPassword = Lcase(Request("password"))
	uName = Trim(Request("name"))
	uLevel = CLng(Request("Level"))
	uGroup = CLng(Request("Group"))
	uBirthday = Trim(Request("Birthday"))
	uaddtime = Trim(Request("addtime"))
	uEmail = Trim(Request("Email"))
	uMobile = Trim(Request("Mobile"))
	ucard = Trim(Request("card"))
	uAddress = Trim(Request("Address"))
	uClientNum = Trim(Request("ClientNum"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [user] Where uName = '" & uName & "' or uAccount = '" & uAccount & "' ",conn,1,1
	If rs.RecordCount > 0 Then
		Response.Write("<script>alert(""��¼�����������ظ�"");history.back(1);</script>")
		Response.End()
	End If
	rs.Close
	rs.Open "Select Top 1 * From [user]",conn,3,2
	rs.AddNew
	rs("uAccount") = EasyCrm.ReName(uAccount)
	rs("uPassword") = md5(uPassword,16)
	rs("uName") = uName
	rs("uLevel") = uLevel
	rs("uGroup") = uGroup
	rs("uMobile") = uMobile
	rs("uEmail") = uEmail
	rs("uAddress") = uAddress
	if uBirthday <> "" then 
	rs("uBirthday") = uBirthday
	end if
	rs("ucard") = ucard
	if uaddtime <> "" then
	rs("uaddtime") = uaddtime
	end if
	rs("uClientNum") = uClientNum
	rs("uQxflag") = EasyCrm.getNewItem("system_level","lId",""&uLevel&"","lQxfalg")
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("?")
End Sub
%>
    <div class="clear"></div>
    </div>
    <!-- end page -->
<script type="text/javascript" src="js/frame.js"></script>
</body>
</html>
