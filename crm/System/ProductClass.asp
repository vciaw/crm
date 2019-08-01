<!--#include file="../data/conn.asp" -->
<%
If mid(Session("CRM_qx"), 4, 1) = 1 Then
Function list() '产品分类列表
    Dim strToPrint
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select * From [ProductClass] where pClassFId = '0' order by pClassId asc ",conn,3,1
    Do While Not rs.BOF And Not rs.EOF
    	strToPrint = strToPrint & "        <tr class=""tr"">" & VBCrlf
    	strToPrint = strToPrint & "          <td class=""tr_f""><a href=""?action=editFPC&pClassId="&rs("pClassId")&""">" & rs("pClassname") & "</a> <input type=""button"" class=""button_ico_Add"" value=""新增子类"" onClick=window.location.href=""?action=addSPC&FPC=" & rs("pClassId") & """ /></td>" & VBCrlf
    	strToPrint = strToPrint & "          <td class=""td_l_c""><input type=""button"" class=""button_edit"" value="""&L_Edit&""" onClick=window.location.href=""?action=editFPC&pClassId="&rs("pClassId")&""" /> <input type=""button"" class=""button_del"" value="""&L_Del&""" onClick=window.location.href=""?action=delete&pClassId="&rs("pClassId")&""" /></td>" & VBCrlf
		
			'子分类列表
			Dim rss
			Set rss = Server.CreateObject("ADODB.Recordset")
			rss.Open "Select * From [ProductClass] where pClassFid ='" & rs("pClassId") & "' ",conn,3,1
			Do While Not rss.BOF And Not rss.EOF
				strToPrint = strToPrint & "        <tr class=""tr"">" & VBCrlf
				strToPrint = strToPrint & "          <td class=""td_l_l"" style=""padding-left:30px;""><a href=""?action=editSPC&pClassId="&rss("pClassId")&""">" & rss("pClassname") & "</a></td>" & VBCrlf
				strToPrint = strToPrint & "          <td class=""td_l_c""><input type=""button"" class=""button_edit"" value="""&L_Edit&""" onClick=window.location.href=""?action=editSPC&pClassId="&rss("pClassId")&""" /> <input type=""button"" class=""button_del"" value="""&L_Del&""" onClick=window.location.href=""?action=delete&pClassId="&rss("pClassId")&""" /></td>" & VBCrlf
				rss.MoveNext
			Loop
			rss.Close
			Set rss = Nothing
			
    	rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
	list = strToPrint
End Function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
</head>
<body style="padding-top:35px;">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：系统管理 > 产品类别管理</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>


<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pd10">
<%
action = Trim(Request.QueryString("action"))
Select Case action

Case "addFPC" '新增产品大类
    Call addFatherProductClass()
Case "saveFPC" '保存产品大类
    Call saveFatherProductClass()
Case "editFPC" '修改产品大类
    Call editFatherProductClass()
Case "editsaveFPC" '保存修改产品大类
    Call editsaveFatherProductClass()
	
Case "addSPC" '新增产品小类
    Call addSonProductClass()
Case "saveSPC" '保存产品小类
    Call saveSonProductClass()
Case "editSPC" '修改产品小类
    Call editSonProductClass()
Case "editsaveSPC" '保存修改产品小类
    Call editsaveSonProductClass()
	
Case "delete" '删除产品分类
    Call deleteData()
	
Case "add"
    Call addOrEdit()
Case "save"
    Call saveData()
Case "edit"
    Call addOrEdit()
Case "restore"
    Call restore()
Case Else
    Call main()
End Select

Sub main()  '默认显示产品分类列表
%>

	<%
	Dim rsa '没有任何分类，跳转到添加产品大类，否则显示产品分类列表
	Set rsa = Server.CreateObject("ADODB.Recordset")
	rsa.Open "Select * From Productclass",conn,1,1
	If rsa.RecordCount = 0 Then
		Response.Redirect("?action=addFPC")
	else 
	%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
				  <td class="td_l_l"><input type="button" class="button_Submit" value="新增大类" onClick=window.location.href="?action=addFPC" /></td>
				  <td class="td_l_c" width="150">管理</td>
				  <% = list() %>
				</tr>
			</table>
	<%
	End If
	rsa.Close
	Set rsa = Nothing
	%>

<%
End Sub

Sub addFatherProductClass() '新增产品大类
%>
		<form name="FatherProductClass" action="?action=saveFPC" method="post">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" />
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="4"><B>产品大类</B></td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_c title">类别名称</td>
					<td class="td_r_l">
						<input name="pClassname" type="text" class="int" id="pClassname" size="30" maxlength="16" value="">
					</td>
				</tr>
				<tr> 
					<td class="td_r_l" colspan="4"><input type="submit" class="button_Submit" name="Submit" value="<%=L_Submit%>"> <input name="Back" type="button" class="button_back" id="Back" value=" <%=l_Back%> " onClick="location.href='?otype=<%=otype%>';"></td>
				</tr>
			</table>
		</form>
<%
end Sub

Sub saveFatherProductClass() '保存新增产品大类
    Dim pClassname
	pClassname = Trim(Request.Form("pClassname"))
	If pClassname = "" Then
        Response.Write("<script>alert(""产品分类名不能为空"");history.back(1);</script>")
		Exit Sub
	End If
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select  * From Productclass Where pClassname = '" & pClassname & "'",conn,3,2
	If rs.RecordCount > 0 Then
        Response.Write("<script>alert(""已存在"");history.back(1);</script>")
		rs.Close
		Set rs = Nothing
		Exit Sub
	Else
	    rs.AddNew
		rs("pClassFid") = 0
		rs("pClassname") = pClassname
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("?")
	End If
End Sub

Sub addSonProductClass() '新增产品小类
	FPC = Request("FPC") '获取大类ID
%>
		<form name="FatherProductClass" action="?action=saveSPC" method="post">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" />
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="4"><B>产品小类</B></td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_c title">上级分类</td>
					<td class="td_r_l">
						<select name="pClassFid" class="int">
							<option value="">请选择</option>
							<% 
								Set rsb = Conn.Execute("select * from ProductClass where pClassFid = '0' ")
								If Not rsb.Eof then
								Do While Not rsb.Eof
								pClassid= rsb("pClassid")
								pClassname= rsb("pClassname")
							%>
							<option value="<%=pClassid%>" <%if ""&pClassid&"" = ""&FPC&"" then%>selected<%end if%>><%=pClassname%></option>
							<%
								rsb.Movenext
								Loop
								End If
								rsb.Close
								Set rsb = Nothing 
							%>
						</select> 
					</td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_c title">类别名称</td>
					<td class="td_r_l">
						<input name="pClassname" type="text" class="int" id="pClassname" size="30" maxlength="16" value="">
					</td>
				</tr>
				<tr> 
					<td class="td_l_l" colspan="4"><input type="submit" class="button_Submit" name="Submit" value="<%=L_Submit%>"> <input name="Back" type="button" class="button_back" id="Back" value=" <%=l_Back%> " onClick="location.href='?otype=<%=otype%>';"></td>
				</tr>
			</table>
		</form>
<%
end Sub

Sub saveSonProductClass() '保存新增产品小类
    Dim pClassFid,pClassname
	pClassFid = Trim(Request.Form("pClassFid"))
	pClassname = Trim(Request.Form("pClassname"))
	If pClassFid = "" Then
        Response.Write("<script>alert(""产品大类不能为空"");history.back(1);</script>")
		Exit Sub
	End If
	If pClassname = "" Then
        Response.Write("<script>alert(""产品子类不能为空"");history.back(1);</script>")
		Exit Sub
	End If
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From Productclass Where pClassFid='"&pClassFid&"' and pClassname = '" & pClassname & "'",conn,3,2
	If rs.RecordCount > 0 Then
        Response.Write("<script>alert(""已存在"");history.back(1);</script>")
		rs.Close
		Set rs = Nothing
		Exit Sub
	Else
	    rs.AddNew
		rs("pClassFid") = pClassFid
		rs("pClassname") = pClassname
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("?")
	End If
End Sub

Sub editFatherProductClass() '产品大类修改
	Dim pClassid
		pClassid = Request("pClassid")
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From Productclass Where pClassid = " & pClassid,conn,1,1
		pClassname = rs("pClassname")
%>
		<form name="FatherProductClass" action="?action=editsaveFPC" method="post">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" />
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="4"><B>产品大类</B></td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_c title">类别名称</td>
					<td class="td_l_l">
						<input name="pClassname" type="text" class="int" id="pClassname" size="30" maxlength="16" value="<%=pClassname%>">
						<input name="pClassid" type="hidden" id="pClassid" value="<% = pClassid %>">
					</td>
				</tr>
				<tr> 
					<td class="td_l_l" colspan="4"><input type="submit" class="button_Submit" name="Submit" value="<%=L_Submit%>"> <input name="Back" type="button" class="button_back" id="Back" value=" <%=l_Back%> " onClick="location.href='?otype=<%=otype%>';"></td>
				</tr>
			</table>
		</form>
<%
		rs.Close
	Set rs = Nothing
End Sub

Sub editsaveFatherProductClass() '保存产品大类修改
	Dim pClassid,pClassname
	pClassid = Request("pClassid")
	pClassname = Trim(Request.Form("pClassname"))
	If pClassname = "" Then
        Response.Write("<script>alert(""产品大类不能为空"");history.back(1);</script>")
		Exit Sub
	End If
		
		rs.Open "Select * From Productclass Where pClassname = '" & pClassname & "' And pClassid <> " & pClassid,conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>alert(""该分类已存在！"");history.back(1);</script>")
		Response.End()
		End If
		rs.Close
			
	rs.Open "Select * From Productclass Where pClassid = " & pClassid,conn,3,2
	rs("pClassname") = pClassname
    rs.Update
    rs.Close
    Set rs = Nothing
	Response.Redirect("?")
End Sub

Sub editSonProductClass() '产品小类修改
	Dim pClassid
		pClassid = Request("pClassid")
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From Productclass Where pClassid = " & pClassid,conn,1,1
		pClassFid = rs("pClassFid")
		pClassname = rs("pClassname")
%>
		<form name="SonProductClass" action="?action=editsaveSPC" method="post">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" />
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="4"><B>产品小类</B></td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_c title">上级分类</td>
					<td class="td_l_l">
						<select name="pClassFid" class="int">
							<option value="">请选择</option>
							<% 
								Set rsb = Conn.Execute("select * from ProductClass where pClassFid = '0' ")
								If Not rsb.Eof then
								Do While Not rsb.Eof
							%>
							<option value="<%=rsb("pClassid")%>" <%if ""&pClassFid&"" = ""&rsb("pClassid")&"" then%>selected<%end if%>><%=rsb("pClassname")%></option>
							<%
								rsb.Movenext
								Loop
								End If
								rsb.Close
								Set rsb = Nothing 
							%>
						</select> 
					</td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_c title">类别名称</td>
					<td class="td_l_l">
						<input name="pClassname" type="text" class="int" id="pClassname" size="30" maxlength="16" value="<%=pClassname%>">
						<input name="pClassid" type="hidden" id="pClassid" value="<% = pClassid %>">
					</td>
				</tr>
				<tr> 
					<td class="td_l_l" colspan="4"><input type="submit" class="button_Submit" name="Submit" value="<%=L_Submit%>"> <input name="Back" type="button" class="button_back" id="Back" value=" <%=l_Back%> " onClick="location.href='?otype=<%=otype%>';"></td>
				</tr>
			</table>
		</form>
<%
		rs.Close
	Set rs = Nothing
End Sub

Sub editsaveSonProductClass() '保存产品小类修改
	Dim pClassid,pClassFid,pClassname
	pClassid = Request("pClassid")
	pClassFid = Trim(Request.Form("pClassFid"))
	pClassname = Trim(Request.Form("pClassname"))
	
	If pClassFid = "" Then
        Response.Write("<script>alert(""产品大类不能为空"");history.back(1);</script>")
		Exit Sub
	End If
	If pClassname = "" Then
        Response.Write("<script>alert(""产品子类不能为空"");history.back(1);</script>")
		Exit Sub
	End If
		
		rs.Open "Select * From Productclass Where pClassFid = '"&pClassFid&"' And pClassname = '"&pClassname&"' And pClassid <> "&pClassid,conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>alert(""该子分类已存在！"");history.back(1);</script>")
		Response.End()
		End If
		rs.Close
			
	rs.Open "Select * From Productclass Where pClassid = " & pClassid,conn,3,2
	rs("pClassFid") = pClassFid
	rs("pClassname") = pClassname
    rs.Update
    rs.Close
    Set rs = Nothing
	Response.Redirect("?")
End Sub

Sub deleteData() '删除
    Dim pClassId
	pClassId = Request("pClassId")
	If pClassId = "" Then
        Response.Write("<script>alert("""&alert03&""");history.back(1);</script>")
		Exit Sub
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From Productclass Where pClassFId = '"&pClassId&"'",conn,1,1 '判断当前分类下是否存在子分类
	If rs.RecordCount > 0 Then
        Response.Write("<script>alert(""该分类有子分类，请先删除小分类！"");history.back(1);</script>")
	else
		Dim rss
		Set rss = Server.CreateObject("ADODB.Recordset")
		rss.Open "Select * From Productclass Where pClassId = " & pClassId,conn,3,2
		If rss.RecordCount > 0 Then
			rss.Delete
			rss.Update
		End If
		rss.Close
		Set rss = Nothing
		Response.Redirect("?")
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