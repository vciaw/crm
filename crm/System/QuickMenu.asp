<!--#include file="../data/conn.asp" -->
<%
	'获取get值
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	tipinfo = 	Request.QueryString("tipinfo")
	if otype="" then otype="Installed"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/modify.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>

<body style="padding-top:35px;"> 
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：系统管理  > 快捷菜单</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10 pdb10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="8"><B><%=L_Top_Plugin%></B></td>
				</tr>
				<tr class="tr_f">
					<td width="60" class="td_l_c">编号</td>
					<td width="100" class="td_l_l">顺序</td>
					<td width="100" class="td_l_l">显示标题</td>
					<td class="td_l_l">实际路径</td>
					<td width="70" class="td_l_c">首页显示</td>
				</tr>
		<%
		Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [QuickMenu] Order By Id asc",conn,1,1
		If rs.RecordCount > 0 Then
		Do While Not rs.BOF And Not rs.EOF
		%>
				<tr class="tr">
					<td class="td_l_c"><%=rs("id")%></td>
					<td class="td_l_l"><span id="Sort<%=rs("id")%>" name="Sort<%=rs("id")%>">
					<input type="text"  size="5" value="<%=rs("Sort")%>" maxlength="20" style="border:1px solid #FFFFFF; ime-mode:Disabled; cursor:hand;width:80px;" onClick="forder('<%=rs("id")%>','<%=rs("Sort")%>','QuickMenu','Sort','id','Sort');"></span></td>
					<td class="td_l_l"><%=rs("Title")%></td>
					<td class="td_l_l"><%=rs("Url")%></td>
					<td class="td_l_c">
					<%if rs("QuickYN")="1" then%>
						<input type="button" class="button222" value="<%=L_Shi%>" onClick="window.location.href='?action=QuickYNSet&id=<%=rs("id")%>&QuickYN=0'" />
					<%else%>
						<input type="button" class="button227" value="<%=L_Fou%>" onClick="window.location.href='?action=QuickYNSet&id=<%=rs("id")%>&QuickYN=1'" />
					<%end if%>
					</td>
				</tr>
		<%
		rs.MoveNext
		Loop
		else
		Response.Write "<tr class=""tr""><td class=""td_l_l"" colspan=8>"&L_Notfound&"</td></tr>" & VBCrlf
		end if
		rs.Close
		Set rs = Nothing
		%>
			</table>
        </td>
	</tr>
</table>
<%
Select Case action
Case "QuickYNSet"
    Call QuickYNSet()
Case "YNSet"
    Call YNSet()
End Select

Sub QuickYNSet()
  Dim id
    id = Trim(Request("id"))
    QuickYN = Trim(Request("QuickYN"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From QuickMenu Where id = " & id,conn,3,2
	If rs.RecordCount > 0 Then
		sql="UPDATE QuickMenu SET QuickYN='"&QuickYN&"' Where id In (" & id & ")"
		conn.execute sql
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
    Response.Write("<script>location.href='?' ;</script>")
End Sub
%>
</body>
</html>
<script src="../data/calendar/WdatePicker.js"></script>