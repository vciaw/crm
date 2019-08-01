<!--#include file="../data/conn.asp" -->
<%
If mid(Session("CRM_qx"), 4, 1) = 1 Then

	'获取get值
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	if otype="" then otype="Select_Type"

Function list()
    Dim strToPrint
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select * From [SelectData] where "&otype&" <>'' and "&otype&" <> 'null'  ",conn,3,1
    Do While Not rs.BOF And Not rs.EOF
    	strToPrint = strToPrint & "        <tr class=""tr"">" & VBCrlf
    	strToPrint = strToPrint & "          <td class=""td_l_l""><a href=""?action=edit&otype="&otype&"&otypedataOld=" & rs(""&otype&"") & """>" & rs(""&otype&"") & "</a></td>" & VBCrlf
    	strToPrint = strToPrint & "          <td class=""td_l_c"">" & VBCrlf
    	strToPrint = strToPrint & "          <input type=""button"" class=""button_info_edit"" value="" "" title="""&L_Edit&""" onClick=""window.location.href='?action=edit&otype="&otype&"&otypedataOld=" & rs(""&otype&"") & "'"" /> " & VBCrlf
    	strToPrint = strToPrint & "          <input type=""button"" class=""button_info_del"" value="" ""  title="""&L_Del&""" onClick=""window.location.href='?action=delete&otype="&otype&"&otypedataOld=" & rs(""&otype&"") & "'"" onClick=""return confirm('"&Alert_del_YN&"');"" />" & VBCrlf
    	strToPrint = strToPrint & "          </td>" & VBCrlf
		
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
<script language="JavaScript">
<!--
function checkInput(o)
{
    var oo = eval("document.all." + o);
    var num = oo.length;
    for(var i=0;i<num;i++){
	    if(oo[i].value == ""){
		    alert("不能为空！");
			oo[i].focus();
			return false
			break;
		}
	}
}
-->
</script>
</head>

<body style="padding-top:35px;">

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：系统管理 > 下拉框管理</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10">   
            <div class="MenuboxS">
              <ul>
                <li style="margin-top:10px;" <%if otype="Select_Type" then%>class="hover"<%end if%>><span><a href="?otype=Select_Type">房产类型</a></span></li>
				 <li style="margin-top:10px;" <%if otype="Select_Mtype" then%>class="hover"<%end if%>><span><a href="?otype=Select_Mtype">客户类型</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_Source" then%>class="hover"<%end if%>><span><a href="?otype=Select_Source">客户来源</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_Star" then%>class="hover"<%end if%>><span><a href="?otype=Select_Star">客户级别</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_Zhiwei" then%>class="hover"<%end if%>><span><a href="?otype=Select_Zhiwei">职位</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_Records" then%>class="hover"<%end if%>><span><a href="?otype=Select_Records">跟单类型</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_Hetong" then%>class="hover"<%end if%>><span><a href="?otype=Select_Hetong">合同分类</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_Service" then%>class="hover"<%end if%>><span><a href="?otype=Select_Service">售后类型</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_ExpenseIN" then%>class="hover"<%end if%>><span><a href="?otype=Select_ExpenseIN">收入类型</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_ExpenseOUT" then%>class="hover"<%end if%>><span><a href="?otype=Select_ExpenseOUT">支出类型</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_SoftClass" then%>class="hover"<%end if%>><span><a href="?otype=Select_SoftClass">文件分类</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_NoticeClass" then%>class="hover"<%end if%>><span><a href="?otype=Select_NoticeClass">公文分类</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_Sex" then%>class="hover"<%end if%>><span><a href="?otype=Select_Sex">性别</a></span></li>
                <li style="margin-top:10px;" <%if otype="Select_YN" then%>class="hover"<%end if%>><span><a href="?otype=Select_YN">是/否</a></span></li>
              </ul>
            </div>
		</td>
	</tr>
	<tr>
		<td valign="top" class="td_n">
<%

Select Case action
	Case "add"
		Call addOrEdit()
	Case "save"
		Call saveData()
	Case "edit"
		Call addOrEdit()
	Case "restore"
		Call restore()
	Case "delete"
		Call deleteData()
	Case Else
		Call addOrEdit()
End Select
	
Sub addOrEdit()
    Dim otypedata,otypedataOld,strOut,strAction
	If action = "edit" Then
	    Dim rs
		otypedataOld = Trim(Request("otypedataOld"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [SelectData] Where "&otype&" = '" & otypedataOld & "'",conn,3,1
		If rs.RecordCount = 1 Then
			otypedata = rs(""&otype&"")
		End If
		rs.Close
		Set rs = Nothing
		strOut = ""&L_Edit&""
		strAction = "?action=restore&otype="&otype&""
	Else
	    strOut = ""&L_Add&""
		strAction = "?action=save&otype="&otype&""
	End If		    
%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
					<form name="menuForm" action="<% = strAction %>" method="post" onSubmit="return checkInput('menuForm');">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B><% = strOut %></B>
							  <% If action = "edit" Then %>
							  <input name="otypedataOld" type="hidden" id="otypedataOld" value="<% = otypedataOld %>">
							  <% End If %></td>
							</tr>
							<tr class="tr"> 
								<td class="td_l_c title">参　数</td>
								<td class="td_r_l"><input name="otypedata" type="text" class="int" id="otypedata" size="30" maxlength="16" value="<% = otypedata %>"></td>
							</tr>
							<tr> 
								<td class="td_r_l" colspan="4"><input type="submit" class="button45" name="Submit" value="<%=L_Submit%>"> <% If action = "edit" Then %><input name="Back" type="button" class="button43" id="Back" value=" 取消 " onClick="location.href='?otype=<%=otype%>';"><% End If %></td>
							</tr>
						</table>
					</form>
					</td>
				</tr>
			</table>
<%
End Sub
%>
		</td>
	</tr>
	<tr>
		<td valign="top" style="padding:10px;" class="td_n"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
				  <td class="td_l_c">类型</td>
				  <td width="90" class="td_l_c">管理</td>
				  <% = list() %>
				</tr>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%
Sub saveData() '保存数据
    Dim otypedata
	otypedata = Trim(Request.Form("otypedata"))
	If otypedata = "" Then
        Response.Write("<script>alert("""&alert03&""");history.back(1);</script>")
		Exit Sub
	End If
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select  * From [SelectData] Where "&otype&" = '" & otypedata & "'",conn,3,2
	If rs.RecordCount > 0 Then
        Response.Write("<script>alert("""&alert02&""");history.back(1);</script>")
		rs.Close
		Set rs = Nothing
		Exit Sub
	Else
	    rs.AddNew
		rs(""&otype&"") = otypedata
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("?otype="&otype&"")
	End If
End Sub

Sub restore() '修改数据
    Dim otypedataOld,otypedata
	otypedataOld = Trim(Request.Form("otypedataOld"))
	otypedata = Trim(Request.Form("otypedata"))
	If otypedataOld = "" Or otypedata = "" Then
        Response.Write("<script>alert("""&alert03&""");history.back(1);</script>")
		Exit Sub
	End If
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [SelectData] Where "&otype&" <> '" & otypedataOld & "'",conn,3,1
	Do While Not rs.BOF And Not rs.EOF
	    If rs(""&otype&"") = otypedata Then
        Response.Write("<script>alert("""&alert02&""");history.back(1);</script>")
		    rs.Close
		    Set rs = Nothing
		    Exit Sub
		End If
		rs.MoveNext
	Loop
	rs.Close
	
	rs.Open "Select * From [SelectData] Where "&otype&" = '" & otypedataOld & "'",conn,3,2
	If rs.RecordCount = 1 Then
	    otypedataOld = rs(""&otype&"")
		rs(""&otype&"") = otypedata
		rs.Update
	End If
	Set rs = Nothing
	Response.Redirect("?otype="&otype&"")
End Sub

Sub deleteData() '删除数据
    Dim otypedataOld
	otypedataOld = Trim(Request("otypedataOld"))
        'Response.Write("<script>alert("""&otype&""");</script>")
		'Response.end
	If otypedataOld = "" Then
        Response.Write("<script>alert("""&alert03&""");history.back(1);</script>")
		Exit Sub
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [SelectData] Where "&otype&" = '" & otypedataOld & "'",conn,3,2
	If rs.RecordCount > 0 Then
		rs.Delete
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
	Response.Redirect("?otype="&otype&"")
End Sub

else
Response.write"<script>alert("""&alert31&""");location.href=""../"";</script>"
end if
%>
