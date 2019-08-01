<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
If Session("CRM_level") < 9 Then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
end if
	'获取get值
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
	<!-- 必填项提示
	function CheckInput()
	{
		if(document.getElementById('account').value == ""){alert("登录账号不能为空！"); document.all.account.focus();return false;}
		if(document.getElementById('name').value == ""){alert("真实姓名不能为空！"); document.all.name.focus();return false;}
		if(document.getElementById('password').value == ""){alert("登录密码不能为空！"); document.all.password.focus();return false;}
		if(document.getElementById('level').value == ""){alert("系统角色不能为空！"); document.all.level.focus();return false;}
		if(document.getElementById('group').value == ""){alert("所属部门不能为空！"); document.all.group.focus();return false;}
	}
	-->
	</script>
                <div class="content listbox" style="margin-bottom:10px;">
				<form name="Save" action="?action=SaveAdd" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> 登录账号</label>
						<input name="account" type="text" class="int" id="account" size="11" maxlength="16"> 录入后不可修改
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> 真实姓名</label>
						<input name="name" type="text" class="int" id="name" size="11" maxlength="16">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> 登录密码</label>
						<input name="password" type="password" class="int" id="password" size="11" maxlength="16">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> 系统角色</label>
						<% = EasyCrm.getList(2,"system_level","lId","lName","level","") %>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> 所属部门</label>
						<% = EasyCrm.getList(2,"system_group","gId","gName","group","") %>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">生日</label>
						<input name="Birthday" type="date" id="Birthday" class="int Wdate" size="15" onFocus="WdatePicker()">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">入司时间</label>
						<input name="addtime" type="date" id="addtime" class="int Wdate" size="15" onFocus="WdatePicker()">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">手机</label>
						<input name="Mobile" type="number" class="int" id="Mobile" size="20" maxlength="16">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">E-mail</label>
						<input name="Email" type="email" class="int" id="Email">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">住址</label>
						<input name="Address" type="text" class="int" id="Address">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">身份证</label>
						<input name="card" type="number" class="int" id="card">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">最大客户量</label>
						<input name="ClientNum" type="number" id="ClientNum" class="int" size="10"  value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <span class="info_help help01" >&nbsp;０为不限制！</span>
                    </div>
                    
                    <div class="form-line">
                    <input type="submit" name="button" id="button" value="&nbsp;保 存&nbsp;" class="submit-button" />
                    </div>
                </form>
                </div>

            	<h1 class="titleh">员工列表</h1>
					<table class="tabledata"> 
						<tbody> 
                        <tr> 
							<td >编号</td> 
							<td>帐号</td> 
							<td>姓名</td> 
							<td>部门</td> 
							<td>角色</td> 
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
					<blockquote>手机版不提供员工编辑和删除功能，以免误操作</blockquote>
			</div>
		<%=Footer%>
            
<%

Select Case action
Case "SaveAdd" '添加
    Call SaveAdd()
End Select

Sub SaveAdd() '添加
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
		Response.Write("<script>alert(""登录名或姓名有重复"");history.back(1);</script>")
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
