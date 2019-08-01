<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
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
		<td class="top_left td_t_n td_r_n">当前位置：插件管理 > 已安装插件</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
            <div class="MenuboxS">
              <ul>
                <li <%if otype="Installed" or otype="" then%>class="hover"<%end if%>><span><a href="?otype=Installed">已安装</a></span></li>
                <li <%if otype="AllPlugin" then%>class="hover"<%end if%>><span><a href="?otype=AllPlugin">所有插件</a></span></li>
              </ul> 
            </div>
		</td>
	</tr>
</table>

<%If otype="Installed" then%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="8"><B><%=L_Top_Plugin%></B></td>
				</tr>
				<tr class="tr_f">
					<td width="60" class="td_l_c">编号</td>
					<td width="100" class="td_l_l">顺序</td>
					<td width="100" class="td_l_l">插件名称</td>
					<td class="td_l_l">功能描述</td>
					<td width="120" class="td_l_c">插件开发</td>
					<td width="80" class="td_l_c">版本</td>
					<td width="70" class="td_l_c">首页显示</td>
					<td width="100" class="td_l_c">管理</td>
				</tr>
		<%
		Dim rsplugin
		Set rsplugin = Server.CreateObject("ADODB.Recordset")
		rsplugin.Open "Select * From plugin Order By Id asc",conn,1,1
		If rsplugin.RecordCount > 0 Then
		Do While Not rsplugin.BOF And Not rsplugin.EOF
		%>
				<tr>
					<td class="td_l_c"><%=rsplugin("id")%></td>
					<td class="td_l_c"><span id="Sort<%=rsplugin("id")%>" name="Sort<%=rsplugin("id")%>">
					<input type="text"  size="5" value="<%=rsplugin("pSort")%>" maxlength="20" style="border:1px solid #FFFFFF; ime-mode:Disabled; cursor:hand;width:80px;" onClick="forder('<%=rsplugin("id")%>','<%=rsplugin("pSort")%>','plugin','pSort','id','Sort');"></span></td>
					<td class="td_l_c"><span id="Title<%=rsplugin("id")%>" name="Title<%=rsplugin("id")%>">
					<input type="text"  size="5" value="<%=rsplugin("pTitle")%>" maxlength="20" style="border:1px solid #FFFFFF; ime-mode:Disabled; cursor:hand;width:80px;" onClick="forder('<%=rsplugin("id")%>','<%=rsplugin("pTitle")%>','plugin','pTitle','id','Title');"></span></td>
					<td class="td_l_l"><%=rsplugin("pContent")%></td>
					<td class="td_l_c"><%=rsplugin("pAuthor")%></td>
					<td class="td_l_c"><%=rsplugin("pVersion")%></td>
					<td class="td_l_c">
					<%if rsplugin("QuickYN")="1" then%>
						<input type="button" class="button222" value="<%=L_Shi%>" onClick="window.location.href='?action=QuickYNSet&id=<%=rsplugin("id")%>&QuickYN=0'" />
					<%else%>
						<input type="button" class="button227" value="<%=L_Fou%>" onClick="window.location.href='?action=QuickYNSet&id=<%=rsplugin("id")%>&QuickYN=1'" />
					<%end if%>
					</td>
					<td class="td_l_c">
					<%if rsplugin("pYn")="1" then%>
						<input type="button" class="button242" value="<%=L_Plugin_pYn_1%>" onClick="window.location.href='?action=YNSet&id=<%=rsplugin("id")%>&pYn=0'" />
					<%else%>
						<input type="button" class="button247" value="<%=L_Plugin_pYn_0%>" onClick="window.location.href='?action=YNSet&id=<%=rsplugin("id")%>&pYn=1'" />
					<%end if%>
					</td>
				</tr>
		<%
		rsplugin.MoveNext
		Loop
		else
		Response.Write "<tr class=""tr""><td class=""td_l_l"" colspan=8>"&L_Notfound&"</td></tr>" & VBCrlf
		end if
		rsplugin.Close
		Set rsplugin = Nothing
		%>
			</table>
        </td>
	</tr>
</table>
<%elseif otype="AllPlugin" then%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="5"><B>所有插件</B></td>
				</tr>
				<tr class="tr_f">
					<td width="100" class="td_l_c"><%=L_Plugin_pTitle%></td>
					<td class="td_l_l"><%=L_Plugin_pContent%></td>
					<td width="120" class="td_l_c"><%=L_Plugin_pAuthor%></td>
					<td width="80" class="td_l_c"><%=L_Plugin_pVersion%></td>
					<td width="100" class="td_l_c"><%=L_Top_Manage%></td>
				</tr>
		<%
		Set rsplugin = Server.CreateObject("ADODB.Recordset")
		rsplugin.Open "Select top 2 * From plugin Order By Id asc",conn,1,1
		If rsplugin.RecordCount > 0 Then
		Do While Not rsplugin.BOF And Not rsplugin.EOF
		%>
				<tr>
					<td class="td_l_c"><%=rsplugin("pTitle")%></td>
					<td class="td_l_l"><%=rsplugin("pContent")%></td>
					<td class="td_l_c"><%=rsplugin("pAuthor")%></td>
					<td class="td_l_c"><%=rsplugin("pVersion")%></td>
					<td class="td_l_c"><input type="button" class="button245" value="内置插件" /></td>
				</tr>
			<%
			rsplugin.MoveNext
			Loop
			else
			Response.Write "<tr class=""tr""><td class=""td_l_l"" colspan=8>"&L_Notfound&"</td></tr>" & VBCrlf
			end if
			rsplugin.Close
			Set rsplugin = Nothing
			%>

			<%
			filepath="./"
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			Set fileobj = fso.GetFolder(server.mappath(filepath))
			Set fsofolders = fileobj.SubFolders
			Set fsofile = fileobj.Files

			For Each folder in fsofolders

				set xmlobj = server.createobject("MSXML2.DOMDocument") 
				xmlobj.load(server.mappath(""&folder.name&"/info.xml"))
				xmlobj.Async = False
				set xmlnodelist = xmlobj.selectNodes("bcaster/item")
				for each node in xmlnodelist
			%>
				<tr>
					<td class="td_l_c"><%=node.getAttribute("Plugin_Title")%></td>
					<td class="td_l_l"><%=node.getAttribute("Plugin_Content")%></td>
					<td class="td_l_c"><%=node.getAttribute("Plugin_Author")%></td>
					<td class="td_l_c"><%=node.getAttribute("Plugin_Version")%></td>
					<td class="td_l_c">
						<%if node.getAttribute("Plugin_url") = EasyCrm.getNewItem("Plugin","pUrl","'"&node.getAttribute("Plugin_url")&"'","pUrl") then%>
						<input type="button" class="button242" value="已安装" />
						<%else%>
						<input type="button" class="button243" value="<%=node.getAttribute("Plugin_Yn")%>" onClick="window.location.href='<%=node.getAttribute("Plugin_install")%>'" />
						<%end if%>
					</td>
				</tr>
			<%
				next
			Next 
			%>
			</table>
        </td>
	</tr>
</table>
<%end if%>
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
	rs.Open "Select * From plugin Where id = " & id,conn,3,2
	If rs.RecordCount > 0 Then
		sql="UPDATE plugin SET QuickYN='"&QuickYN&"' Where id In (" & id & ")"
		conn.execute sql
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
    Response.Write("<script>location.href='index.asp' ;</script>")
End Sub

Sub YNSet()
  Dim id
    id = Trim(Request("id"))
    pYn = Trim(Request("pYn"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From plugin Where id = " & id,conn,3,2
	If rs.RecordCount > 0 Then
		sql="UPDATE plugin SET pYn='"&pYn&"' Where id In (" & id & ")"
		conn.execute sql
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
    Response.Write("<script>location.href='index.asp' ;</script>")
End Sub
%>
<% Set EasyCrm = nothing %>
</body>
</html>