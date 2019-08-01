<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
If Session("CRM_level") < 9 Then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
end if
	'获取get值
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
Case "SaveAdd" '添加
    Call SaveAdd()
Case "Edit" '修改
    Call Edit()
Case "SaveEdit" '保存修改
    Call SaveEdit()
Case else
    Call Main()
End Select

Sub Main()
%>
	<script language="JavaScript">
	<!-- 必填项提示
	function CheckInput()
	{
		if(document.getElementById('gId').value == ""){alert("部门编号不能为空！"); document.all.gId.focus();return false;}
		if(document.getElementById('gName').value == ""){alert("部门名称不能为空！"); document.all.gName.focus();return false;}
	}
	-->
	</script>
                <div class="content listbox" style="margin-bottom:10px;">
				<form name="Save" action="?action=SaveAdd" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> 部门编号</label>
						<input name="gId" type="number" class="int" id="gId" size="11" maxlength="16" min="1" max="99"> 限：数字 1 - 99
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> 部门名称</label>
						<input name="gName" type="text" class="int" id="gName" size="11" maxlength="16">
                    </div>
                    
                    <div class="form-line">
                    <input type="submit" name="button" id="button" value="&nbsp;保 存&nbsp;" class="submit-button" />
                    </div>
                </form>
                </div>

            	<h1 class="titleh">部门列表</h1>
					<table class="tabledata"> 
						<col width="40"><col ><col width="70">
						<tbody> 
                        <tr> 
							<td >编号</td> 
							<td>名称</td> 
							<td>管理</td> 
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
					<blockquote>手机版不提供部门删除功能，以免误操作</blockquote>
			</div>
		<%=Footer%>
            
<%
end Sub

Sub SaveAdd() '添加
	gId = Trim(Request("gId"))
	gName = Trim(Request("gName"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [system_group] Where gId = " & gId & " Or gName = '" & gName & "' ",conn,3,1
	If rs.RecordCount > 0 Then
		Response.Write("<script>alert('部门编号或部门名称有重复');history.back(1);</script>")
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
            	<h1 class="titleh">修改部门</h1>
                <div class="content" style="margin-bottom:10px;">
				<form name="Save" action="?action=SaveEdit" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> 部门编号</label>
						<input name="gId" type="number" class="int" id="gId" size="11" maxlength="16" value="<%=gId%>" min="1" max="99"> 限：数字 1 - 99
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> 部门名称</label>
						<input name="gName" type="text" class="int" id="gName" size="11" maxlength="16" value="<%=gName%>">
                    </div>
                    
                    <div class="form-line">
					<input name="gIdOld" type="hidden" id="gIdOld" value="<%=gID%>">
					<input name="gNameOld" type="hidden" id="gNameOld" value="<%=gName%>">
                    <input type="submit" name="button" id="button" value="&nbsp;保 存&nbsp;" class="submit-button" />
					<input name="Back" type="button" id="Back" class="reset-button" value="返回" onClick="location.href='Group.asp';">
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
	if gId = gIdOld then '如果没更新部门编号
		if gName <> gNameOld then
			'如果只修改部门名称，判断是否与其它部门名称重复
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open "Select * From [system_group] Where gName = '" & gName & "' ",conn,1,1
			If rs.RecordCount > 0 Then
				Response.Write("<script>alert('部门名称有重复');history.back(1);</script>")
			Response.End()
			else
			conn.execute("update [system_group] set gName = '"&gName&"' where gName = '"&gNameOld&"' ")
			End If
			rs.Close
		end if
	else '如果更新了部门编号，同步更新用户表和客户表
	
		'如果部门编号与其它部门重复
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [system_group] Where gId = " & gId & " and gName <> '" & gNameOld & "' ",conn,1,1
		If rs.RecordCount > 0 Then
				Response.Write("<script>alert('部门编号有重复');history.back(1);</script>")
		Response.End()
		End If
		rs.Close
		
		if gName <> gNameOld then 
			'如果更改了部门名称，则判断部门名称是否与别的部门重复
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open "Select * From [system_group] Where gName = '" & gName & "' and gId="&gId&" ",conn,1,1
			If rs.RecordCount > 0 Then
				Response.Write("<script>alert('部门名称有重复');history.back(1);</script>")
			Response.End()
			else
			conn.execute("update [system_group] set gId = '"&gId&"',gName='"&gName&"' where gId = "&gIdOld&" ")
			End If
			rs.Close
		else '如果只修改部门编号，则不考虑部门名称
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
