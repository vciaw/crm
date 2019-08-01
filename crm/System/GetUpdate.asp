<!--#include file="../Data/Conn.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>
<body>
<style>body{padding-bottom:55px;}</style>
<%
action = Trim(Request("action"))
sType = Trim(Request("sType"))
tipinfo = Trim(Request("tipinfo"))

Select Case action
Case "Setting"
    Call Setting()
Case "Products"
    Call Products()
Case "AreaData"
    Call AreaData()
Case "CustomField"
    Call CustomField()
Case "SelectData"
    Call SelectData()
Case "User"
    Call User()
Case "Group"
    Call Group()
Case "Level"
    Call Level()
Case "InfoList"
    Call InfoList()
End Select

Sub Products() '产品数据更新

if tipinfo<>"" then
	Response.Write("<script>art.dialog({title: '提示',time: 1,icon: 'warning',content: '"&tipinfo&"'});</script>")
end if

if sType="ClassList" then
%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10 pdb10">
						<div style="float:left;padding-bottom:10px;width:100%;">
							<span style="float:right;"><input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" /></span>
							<input type="button" class="button45" value="新增大类"  onclick='Products_BClass_Add()' style="cursor:pointer" />　
						</div>
						<script>function Products_BClass_Add() {$.dialog.open('GetUpdate.asp?action=Products&sType=BigClassAdd', {title: '新增产品大类', width: 400,height: 145, fixed: true}); };</script>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t"> 
							  <td class="td_l_l">分类标题</td>
							  <td class="td_l_c" width="120">管理</td>
							</tr>
								<%
								Set rs = Server.CreateObject("ADODB.Recordset")
								rs.Open "Select * From [ProductClass] where pClassFId = '0' order by pClassId asc ",conn,3,1
								Do While Not rs.BOF And Not rs.EOF
								%>
								<tr class="tr">
									<td class="tr_f"><a href="javascript:void(0)" onclick='Products_BClass_Edit<%=rs("pClassId")%>()' title='修改' style="cursor:pointer"><%=rs("pClassname")%></a></td>
									<td class="td_l_r title"><input type="button" class="button_info_add" value=" " title="添加小类"  onclick='Products_SClass_Add<%=rs("pClassId")%>()' style="cursor:pointer" /><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Products_BClass_Edit<%=rs("pClassId")%>()' style="cursor:pointer" /><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onClick=window.location.href="GetUpdate.asp?action=Products&sType=ProductsClassDel&pClassId=<%=rs("pClassId")%>" /></td>
								</tr>
						<script>function Products_BClass_Edit<%=rs("pClassId")%>() {$.dialog.open('GetUpdate.asp?action=Products&sType=BigClassEdit&pClassId=<%=rs("pClassId")%>', {title: '编辑产品大类', width: 400,height: 145, fixed: true}); };</script>
						<script>function Products_SClass_Add<%=rs("pClassId")%>() {$.dialog.open('GetUpdate.asp?action=Products&sType=SmallClassAdd&pClassFid=<%=rs("pClassId")%>', {title: '添加产品小类', width: 400,height: 180, fixed: true}); };</script>
								<%	'子分类列表
										Set rss = Server.CreateObject("ADODB.Recordset")
										rss.Open "Select * From [ProductClass] where pClassFid ='" & rs("pClassId") & "' ",conn,3,1
										Do While Not rss.BOF And Not rss.EOF
								%>
										<tr class="tr">
											<td class="td_l_l" style="padding-left:30px;">┗━━ <a  href="javascript:void(0)" onclick='Products_SClass_Edit<%=rss("pClassId")%>()' title='修改' style="cursor:pointer"><%=rss("pClassname")%></a></td>
											<td class="td_l_r"><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>" onclick='Products_SClass_Edit<%=rss("pClassId")%>()' style="cursor:pointer" /><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onClick=window.location.href="GetUpdate.asp?action=Products&sType=ProductsClassDel&pClassId=<%=rss("pClassId")%>" /></td>
										</tr>
						<script>function Products_SClass_Edit<%=rss("pClassId")%>() {$.dialog.open('GetUpdate.asp?action=Products&sType=SmallClassEdit&pClassId=<%=rss("pClassId")%>', {title: '编辑产品小类', width: 400,height: 180, fixed: true}); };</script>
								<%
											rss.MoveNext
										Loop
										rss.Close
										Set rss = Nothing
										
									rs.MoveNext
								Loop
								rs.Close
								Set rs = Nothing
								%>
						</table>
					</td> 
				</tr>
			</table>
<%
elseif sType="BigClassAdd" then '添加大类
%>
		<form name="Save" action="GetUpdate.asp?action=Products&sType=SaveBigClassAdd" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">类别名称</td>
								<td class="td_l_l"><input name="pClassname" type="text" id="pClassname" class="int" size="40" /></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveBigClassAdd" then
		pClassname = Request.Form("pClassname")
		If pClassname = "" Then
			Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=BigClassAdd&tipinfo=产品分类名不能为空';</script>")
			Exit Sub
		End If
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [Productclass] Where pClassname = '"&pClassname&"' ",conn,3,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=BigClassAdd&tipinfo=已存在！';</script>")
		Response.End()
		End If
		rs.Close

    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [ProductClass] ",conn,3,2
		rs.AddNew
		rs("pClassFid") = 0
		rs("pClassname") = pClassname
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

elseif sType="BigClassEdit" then '修改大类
	pClassid = Request("pClassid")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [ProductClass] Where pClassid = " & pClassid,conn,1,1
	pClassname = rs("pClassname")
%>
		<form name="Save" action="GetUpdate.asp?action=Products&sType=SaveBigClassEdit" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">类别名称</td>
								<td class="td_l_l"><input name="pClassname" type="text" id="pClassname" class="int" size="20" value="<%=pClassname%>" /></td>
								<input name="pClassid" type="hidden" id="pClassid" value="<% = pClassid %>">
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
		rs.Close
	Set rs = Nothing
elseif sType="SaveBigClassEdit" then
		pClassid = Request.Form("pClassid")
		pClassname = Request.Form("pClassname")
		pClassnameOld = Trim(Request.Form("pClassnameOld"))
		If pClassname = "" Then
			Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=BigClassEdit&pClassId="&pClassid&"&tipinfo=产品分类名不能为空';</script>")
			Exit Sub
		End If
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [Productclass] Where pClassname = '"&pClassname&"' And pClassid <> " & pClassid,conn,3,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=BigClassEdit&pClassId="&pClassid&"&tipinfo=已存在！';</script>")
		Response.End()
		End If
		rs.Close

    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select * From [ProductClass] where pClassid="&pClassid&" ",conn,3,2
		rs("pClassFid") = 0
		rs("pClassname") = pClassname
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		conn.execute("update [Products] set pBigClass = '"&pClassname&"' where pBigClass = '"&pClassnameOld&"' ")
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

elseif sType="SmallClassAdd" then '添加小类
		pClassFid = Request("pClassFid")
%>
		<form name="Save" action="GetUpdate.asp?action=Products&sType=SaveSmallClassAdd" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title">上级分类</td>
								<td class="td_r_l">
									<select name="pClassFid" class="int">
										<option value="">请选择</option>
										<% 
											Set rsb = Conn.Execute("select * from [ProductClass] where pClassFid = '0' ")
											If Not rsb.Eof then
											Do While Not rsb.Eof
											pClassid= rsb("pClassid")
											pClassname= rsb("pClassname")
										%>
										<option value="<%=pClassid%>" <%if ""&pClassid&"" = ""&pClassFid&"" then%>selected<%end if%>><%=pClassname%></option>
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
							<tr>
								<td class="td_l_r title">类别名称</td>
								<td class="td_l_l"><input name="pClassname" type="text" id="pClassname" class="int" size="40" /></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveSmallClassAdd" then
	pClassFid = Trim(Request.Form("pClassFid"))
	pClassname = Trim(Request.Form("pClassname"))
	If pClassFid = "" Then
        Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=SmallClassAdd&pClassFid="&pClassFid&"&tipinfo=产品大类不能为空';</script>")
		Exit Sub
	End If
	If pClassname = "" Then
        Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=SmallClassAdd&pClassFid="&pClassFid&"&tipinfo=产品小类不能为空';</script>")
		Exit Sub
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Productclass] Where pClassFid='"&pClassFid&"' and pClassname = '" & pClassname & "'",conn,3,2
	If rs.RecordCount > 0 Then
        Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=SmallClassAdd&pClassFid="&pClassFid&"&tipinfo=已存在！' ;</script>")
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
	End If
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

elseif sType="SmallClassEdit" then '编辑小类
		pClassid = Request("pClassid")
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From Productclass Where pClassid = " & pClassid,conn,1,1
		pClassFid = rs("pClassFid")
		pClassname = rs("pClassname")
%>
		<form name="Save" action="GetUpdate.asp?action=Products&sType=SaveSmallClassEdit" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title">上级分类</td>
								<td class="td_r_l">
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
							<tr>
								<td class="td_l_r title">类别名称</td>
								<td class="td_l_l"><input name="pClassname" type="text" id="pClassname" class="int" size="40" value="<%=pClassname%>" /></td>
								<input name="pClassid" type="hidden" id="pClassid" value="<% = pClassid %>">
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveSmallClassEdit" then
	pClassid = Request.Form("pClassid")
	pClassFid = Trim(Request.Form("pClassFid"))
	pClassname = Trim(Request.Form("pClassname"))
	pClassnameOld = Trim(Request.Form("pClassnameOld"))
	If pClassFid = "" Then
        Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=SmallClassEdit&pClassId="&pClassid&"&tipinfo=产品大类不能为空' ;</script>")
		Exit Sub
	End If
	If pClassname = "" Then
        Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=SmallClassEdit&pClassId="&pClassid&"&tipinfo=产品小类不能为空' ;</script>")
		Exit Sub
	End If
	
	Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [Productclass] Where pClassFid = '"&pClassFid&"' And pClassname = '"&pClassname&"' And pClassid <> "&pClassid,conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=SmallClassEdit&pClassId="&pClassid&"&tipinfo=已存在！' ;</script>")
		Response.End()
		End If
		rs.Close
		
		rs.Open "Select * From Productclass Where pClassid = " & pClassid,conn,3,2
		rs("pClassFid") = pClassFid
		rs("pClassname") = pClassname
		rs.Update
		rs.Close
		Set rs = Nothing
		
		conn.execute("update [Products] set pSmallClass = '"&pClassname&"' where pSmallClass = '"&pClassnameOld&"' ")
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="ProductsClassDel" then '删除产品分类

	pClassId = Request("pClassId")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Productclass] Where pClassFId = '"&pClassId&"'",conn,1,1 '判断当前分类下是否存在子分类
	If rs.RecordCount > 0 Then
        Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=ClassList&tipinfo=有子分类，禁止删除！';</script>")
	else
		Set rss = Server.CreateObject("ADODB.Recordset")
		rss.Open "Select * From [Productclass] Where pClassId = " & pClassId,conn,3,2
		If rss.RecordCount > 0 Then
			rss.Delete
			rss.Update
		End If
		rss.Close
		Set rss = Nothing
		Response.Redirect("GetUpdate.asp?action=Products&sType=ClassList")
	end if
	rs.Close
	Set rs = Nothing
	
elseif sType="InfoAdd" then '添加产品
%>
	<script language="JavaScript">
	function CheckInput(){
		if(document.getElementById('pBigClass').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '产品大类 <%=alert04%>'});document.getElementById('pBigClass').focus();return false;}
		if(document.getElementById('Strade').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '产品小类 <%=alert04%>'});document.getElementById('Strade').focus();return false;}
		if(document.getElementById('pTitle').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '产品名 <%=alert04%>'});document.getElementById('pTitle').focus();return false;}
		if(document.getElementById('pUprice').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '价格 <%=alert04%>'});document.getElementById('pUprice').focus();return false;}
	}
	</script>
<style>body{overflow-y:hidden}</style>
		<form name="Save" action="GetUpdate.asp?action=Products&sType=SaveInfoAdd" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">分类</td>
								<td class="td_l_l">
									<select name="pBigClass" class="int" onchange="getTrade(this.options[this.selectedIndex].id);">
									<option value=""><%=L_Please_choose_01%></option>
									<% 
										Set rsb = Conn.Execute("select * from ProductClass where pClassFid = '0' ")
										If Not rsb.Eof then
										Do While Not rsb.Eof
										pClassid= rsb("pClassid")
										pClassname= rsb("pClassname")
									%>
										<option value="<%=pClassname%>" id="<%=pClassid%>"><%=pClassname%></option>
									<%
										rsb.Movenext
										Loop
										End If
										rsb.Close
										Set rsb = Nothing 
									%>
									</select> 
									<span id="Stradediv"  style="margin-left:10px;padding:0;">
										<select name="Strades" class="int">
											<option value=""><%=L_Please_choose_02%></option>
										</select>
									</span>
									<input name="Strade" type="hidden" id="Strade" class="int">
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">产品名</td>
								<td class="td_l_l"><input name="pTitle" type="text" id="pTitle" class="int" size="30" /></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Order_Products_oProItemA%></td>
								<td class="td_l_l"><input name="pItemA" type="text" id="pItemA" class="int" size="30" /></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Order_Products_oProItemB%></td>
								<td class="td_l_l"><input name="pItemB" type="text" id="pItemB" class="int" size="30" /></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Order_Products_oProItemC%></td>
								<td class="td_l_l"><input name="pItemC" type="text" id="pItemC" class="int" size="30" /></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Order_Products_oProItemD%></td>
								<td class="td_l_l"><input name="pItemD" type="text" id="pItemD" class="int" size="30" /></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Order_Products_oProItemE%></td>
								<td class="td_l_l"><input name="pItemE" type="text" id="pItemE" class="int" size="30" /></td>
							</tr>
							<tr>
								<td class="td_l_r title">单价</td>
								<td class="td_l_l"><input name="pUprice" type="text" id="pUprice" class="int" size="10" value="0" /> 元</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveInfoAdd" then
		pBigClass = Request.Form("pBigClass")
		pSmallClass = Request.Form("Strade")
		pTitle = Request.Form("pTitle")
		pItemA = Request.Form("pItemA")
		pItemB = Request.Form("pItemB")
		pItemC = Request.Form("pItemC")
		pItemD = Request.Form("pItemD")
		pItemE = Request.Form("pItemE")
		pUprice = Request.Form("pUprice")
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		'rs.Open "Select * From [Products] Where pTitle = '"&pTitle&"' ",conn,1,1
		'If rs.RecordCount > 0 Then
		'	Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=InfoAdd&tipinfo=产品名称已存在，请重新输入！';</script>")
		'Response.End()
		'End If
		'rs.Close
		
    	rs.Open "Select Top 1 * From [Products]",conn,3,2
		rs.AddNew
		rs("pBigClass") = pBigClass
		rs("pSmallClass") = pSmallClass
		rs("pTitle") = pTitle
		rs("pItemA") = pItemA
		rs("pItemB") = pItemB
		rs("pItemC") = pItemC
		rs("pItemD") = pItemD
		rs("pItemE") = pItemE
		rs("pUprice") = pUprice
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	
elseif sType="InfoEdit" then '修改产品
	id= Trim(Request("id"))
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From Products Where Id = " & Id,conn,1,1
	If rs.RecordCount <> 1 Then
		  Response.Write("<script>alert("""&alert01&""");history.back(1);</script>")
	Response.End()
	End If

	pBigClass = rs("pBigClass")
	pSmallClass = rs("pSmallClass")
	pTitle = rs("pTitle")
	pItemA = rs("pItemA")
	pItemB = rs("pItemB")
	pItemC = rs("pItemC")
	pItemD = rs("pItemD")
	pItemE = rs("pItemE")
	pUprice = rs("pUprice")
	rs.Close
	Set rs = Nothing
%>
	<script language="JavaScript">
	function CheckInput(){
		if(document.getElementById('pTitle').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '产品名 <%=alert04%>'});document.getElementById('pTitle').focus();return false;}
		if(document.getElementById('pUprice').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '价格 <%=alert04%>'});document.getElementById('pUprice').focus();return false;}
	}
	</script>
<style>body{overflow-y:hidden}</style>
		<form name="Save" action="GetUpdate.asp?action=Products&sType=SaveInfoEdit" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">分类</td>
								<td class="td_l_l">
									<select name="pBigClass" class="int" onchange="getTrade(this.options[this.selectedIndex].id);">
									<option value=""><%=L_Please_choose_01%></option>
									<% 
										Set rsb = Conn.Execute("select * from ProductClass where pClassFid = '0' ")
										If Not rsb.Eof then
										Do While Not rsb.Eof
										pClassid= rsb("pClassid")
										pClassname= rsb("pClassname")
									%>
										<option value="<%=pClassname%>" id="<%=pClassid%>" <%if pClassid = EasyCrm.getNewItem("ProductClass","pClassname","'"&pBigClass&"'","pClassId") then %>selected<%end if%> ><%=pClassname%></option>
									<%
										rsb.Movenext
										Loop
										End If
										rsb.Close
										Set rsb = Nothing 
									%>
									</select> 
									<span id="Stradediv"  style="margin-left:10px;padding:0;">
										<select name="Strades" class="int" onchange="getStrade(options[selectedIndex])">
											<option value=""><%=L_Please_choose_02%></option>
											<% 
											IF ""&pBigClass&""<>"" then
											Set rsb = Conn.Execute("select * from ProductClass where pClassFid='"&EasyCrm.getNewItem("ProductClass","pClassname","'"&pBigClass&"'","pClassId")&"' ")
											If Not rsb.Eof then
											Do While Not rsb.Eof
											pClassname= rsb("pClassname")
											%>
											<option value="<%=pClassname%>" <%if pClassname = EasyCrm.getNewItem("ProductClass","pClassname","'"&pSmallClass&"'","pClassname") then %>selected<%end if%> ><%=pClassname%></option>
											<%rsb.Movenext
											Loop
											End If
											rsb.Close
											Set rsb = Nothing 
											end if
											%>
										</select>
									</span>
									<input name="Strade" type="hidden" id="Strade" class="int">
									
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">产品名</td>
								<td class="td_l_l"><input name="pTitle" type="text" id="pTitle" class="int" size="30" value="<%=pTitle%>" /></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Order_Products_oProItemA%></td>
								<td class="td_l_l"><input name="pItemA" type="text" id="pItemA" class="int" size="30" value="<%=pItemA%>" /></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Order_Products_oProItemB%></td>
								<td class="td_l_l"><input name="pItemB" type="text" id="pItemB" class="int" size="30" value="<%=pItemB%>" /></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Order_Products_oProItemC%></td>
								<td class="td_l_l"><input name="pItemC" type="text" id="pItemC" class="int" size="30" value="<%=pItemC%>" /></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Order_Products_oProItemD%></td>
								<td class="td_l_l"><input name="pItemD" type="text" id="pItemD" class="int" size="30" value="<%=pItemD%>" /></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Order_Products_oProItemE%></td>
								<td class="td_l_l"><input name="pItemE" type="text" id="pItemE" class="int" size="30" value="<%=pItemE%>" /></td>
							</tr>
							<tr>
								<td class="td_l_r title">单价</td>
								<td class="td_l_l"><input name="pUprice" type="text" id="pUprice" class="int" size="10" value="<%=pUprice%>" /> 元</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input name="id" type="hidden" id="id" class="int" size="10" value="<%=id%>" />
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
		<script language="JavaScript">
		<!--
		for(var i=0;i<document.getElementById('pBigClass').options.length;i++){
			if(document.getElementById('pBigClass').options[i].value == "<% = pBigClass %>"){
			document.getElementById('pBigClass').options[i].selected = true;}}

		for(var i=0;i<document.getElementById('Strades').options.length;i++){
			if(document.getElementById('Strades').options[i].value == "<% = pSmallClass %>"){
			document.getElementById('Strades').options[i].selected = true;}}
		-->
		</script>
<%
elseif sType="SaveInfoEdit" then
		id = Request.Form("id")
		pBigClass = Request.Form("pBigClass")
		pSmallClass = Request.Form("Strade")
	if Trim(Request.Form("Strades"))<>"" then 
		pSmallClass = Trim(Request.Form("Strades"))
	else
		if Trim(Request.Form("Strade")) <> "" then 
		pSmallClass = Trim(Request.Form("Strade"))
		else
		pSmallClass = ""
		end if
	end if
		pTitle = Request.Form("pTitle")
		pItemA = Request.Form("pItemA")
		pItemB = Request.Form("pItemB")
		pItemC = Request.Form("pItemC")
		pItemD = Request.Form("pItemD")
		pItemE = Request.Form("pItemE")
		pUprice = Request.Form("pUprice")
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		'rs.Open "Select * From [Products] Where pTitle = '"&pTitle&"' and id <> "&id&" ",conn,1,1
		'If rs.RecordCount > 0 Then
		'	Response.Write("<script>location.href='GetUpdate.asp?action=Products&sType=InfoAdd&tipinfo=产品名称已存在，请重新输入！';</script>")
		'Response.End()
		'End If
		'rs.Close
		
    	rs.Open "Select Top 1 * From [Products] where id="&id&" ",conn,3,2
		rs("pBigClass") = pBigClass
		rs("pSmallClass") = pSmallClass
		rs("pTitle") = pTitle
		rs("pItemA") = pItemA
		rs("pItemB") = pItemB
		rs("pItemC") = pItemC
		rs("pItemD") = pItemD
		rs("pItemE") = pItemE
		rs("pUprice") = pUprice
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
end if
End Sub




Sub SelectData() '添加下拉框内容
	otype	=	Request.QueryString("otype")
	if otype = "" then otype = "Select_Type"
	if sType="Add" then
%>
		<form name="Save" action="?action=SelectData&sType=SaveAdd&oType=<%=otype%>" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>新增下拉框内容</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">参　数</td>
								<td class="td_l_l"><input name="otypedata" type="text" id="otypedata" class="int" size="40" /></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>

<%	
	elseif sType="SaveAdd" then
		otypedata = Trim(Request.Form("otypedata"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select  * From [SelectData] Where "&otype&" = '" & otypedata & "'",conn,3,2
		If rs.RecordCount > 0 Then
			Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
			rs.Close
			Set rs = Nothing
			Exit Sub
		Else
			rs.AddNew
			rs(""&otype&"") = otypedata
			rs.Update
			rs.Close
			Set rs = Nothing
			Response.Write("<script>$.dialog.open.origin.$('#"&otype&"').append('<option value="&otypedata&" selected>"&otypedata&"</option>');art.dialog.close();</script>")
			'Response.Redirect("?otype="&otype&"")
		End If
	end if

End Sub


Sub Setting() '页面设置

if sType="ListAll" then '客户列表配置
%>
		<form name="Save" action="?action=Setting&sType=SaveListAll" method="post"onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>列显示字段</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Client_cDate%></td>
								<td class="td_l_c"><input name="Client_cDate" type="checkbox" id="Client_cDate" value="1" <%if Client_cDate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Client_cCompany%></td>
								<td class="td_l_c"><input name="Client_cCompany" type="checkbox" id="Client_cCompany" value="1" <%if Client_cCompany = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Client_cArea%></td>
								<td class="td_l_c"><input name="Client_cArea" type="checkbox" id="Client_cArea" value="1" <%if Client_cArea = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Client_cSquare%></td>
								<td class="td_l_c"><input name="Client_cSquare" type="checkbox" id="Client_cSquare" value="1" <%if Client_cSquare = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Client_cAddress%></td>
								<td class="td_l_c"><input name="Client_cAddress" type="checkbox" id="Client_cAddress" value="1" <%if Client_cAddress = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Client_cType%></td>
								<td class="td_l_c"><input name="Client_cType" type="checkbox" id="Client_cType" value="1" <%if Client_cType = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Client_cTel%></td>
								<td class="td_l_c"><input name="Client_cTel" type="checkbox" id="Client_cTel" value="1" <%if Client_cTel = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">8.</td>
								<td class="td_l_c"> <%=L_Client_cFax%></td>
								<td class="td_l_c"><input name="Client_cFax" type="checkbox" id="Client_cFax" value="1" <%if Client_cFax = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">9.</td>
								<td class="td_l_c"> <%=L_Client_cTrade%></td>
								<td class="td_l_c"><input name="Client_cTrade" type="checkbox" id="Client_cTrade" value="1" <%if Client_cTrade = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">10.</td>
								<td class="td_l_c"> <%=L_Client_cStrade%></td>
								<td class="td_l_c"><input name="Client_cStrade" type="checkbox" id="Client_cStrade" value="1" <%if Client_cStrade = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">11.</td>
								<td class="td_l_c"> <%=L_Client_cStart%></td>
								<td class="td_l_c"><input name="Client_cStart" type="checkbox" id="Client_cStart" value="1" <%if Client_cStart = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">12.</td>
								<td class="td_l_c"> <%=L_Client_cSource%></td>
								<td class="td_l_c"><input name="Client_cSource" type="checkbox" id="Client_cSource" value="1" <%if Client_cSource = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">13.</td>
								<td class="td_l_c"> <%=L_Client_cLinkman%></td>
								<td class="td_l_c"><input name="Client_cLinkman" type="checkbox" id="Client_cLinkman" value="1" <%if Client_cLinkman = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">14.</td>
								<td class="td_l_c"> <%=L_Client_cZhiwei%></td>
								<td class="td_l_c"><input name="Client_cZhiwei" type="checkbox" id="Client_cZhiwei" value="1" <%if Client_cZhiwei = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">15.</td>
								<td class="td_l_c"> <%=L_Client_cMobile%></td>
								<td class="td_l_c"><input name="Client_cMobile" type="checkbox" id="Client_cMobile" value="1" <%if Client_cMobile = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">16.</td>
								<td class="td_l_c"> <%=L_Client_cUser%></td>
								<td class="td_l_c"><input name="Client_cUser" type="checkbox" id="Client_cUser" value="1" <%if Client_cUser = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">17.</td>
								<td class="td_l_c"> <%=L_Client_cShare%></td>
								<td class="td_l_c"><input name="Client_cShare" type="checkbox" id="Client_cShare" value="1" <%if Client_cShare = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">18.</td>
								<td class="td_l_c"> <%=L_Client_cLastUpdated%></td>
								<td class="td_l_c"><input name="Client_cLastUpdated" type="checkbox" id="Client_cLastUpdated" value="1" <%if Client_cLastUpdated = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">19.</td>
								<td class="td_l_c"> <%=L_Client_cRNextTime%></td>
								<td class="td_l_c"><input name="Client_cRNextTime" type="checkbox" id="Client_cRNextTime" value="1" <%if Client_cRNextTime = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">20.</td>
								<td class="td_l_c"> <%=L_Client_cOEDate%></td>
								<td class="td_l_c"><input name="Client_cOEDate" type="checkbox" id="Client_cOEDate" value="1" <%if Client_cOEDate = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">21.</td>
								<td class="td_l_c"> <%=L_Client_cHEdate%></td>
								<td class="td_l_c"><input name="Client_cHEdate" type="checkbox" id="Client_cHEdate" value="1" <%if Client_cHEdate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">22.</td>
								<td class="td_l_c"> <%=L_Client_cHMoney%></td>
								<td class="td_l_c"><input name="Client_cHMoney" type="checkbox" id="Client_cHMoney" value="1" <%if Client_cHMoney = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">23.</td>
								<td class="td_l_c"> <%=L_Client_cHOwed%></td>
								<td class="td_l_c"><input name="Client_cHOwed" type="checkbox" id="Client_cHOwed" value="1" <%if Client_cHOwed = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">24.</td>
								<td class="td_l_c"> <%=L_Client_cSNum%></td>
								<td class="td_l_c"><input name="Client_cSNum" type="checkbox" id="Client_cSNum" value="1" <%if Client_cSNum = 1 then %>checked<%end if%> ></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10"> 
						<div style="float:left;padding:10px 0;">
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
						</div>
					</td> 
				</tr>
			</table>
		</form>
<%
elseif sType="SaveListAll" then

		'更新列显示字段
		IF Request.Form("Client_cDate") <> "" THEN
		Client_cDate = Request.Form("Client_cDate")
		ELSE
		Client_cDate = 0
		END IF
		IF Request.Form("Client_cCompany") <> "" THEN
		Client_cCompany = Request.Form("Client_cCompany")
		ELSE
		Client_cCompany = 0
		END IF
		IF Request.Form("Client_cArea") <> "" THEN
		Client_cArea = Request.Form("Client_cArea")
		ELSE
		Client_cArea = 0
		END IF
		IF Request.Form("Client_cSquare") <> "" THEN
		Client_cSquare = Request.Form("Client_cSquare")
		ELSE
		Client_cSquare = 0
		END IF
		IF Request.Form("Client_cAddress") <> "" THEN
		Client_cAddress = Request.Form("Client_cAddress")
		ELSE
		Client_cAddress = 0
		END IF
		IF Request.Form("Client_cType") <> "" THEN
		Client_cType = Request.Form("Client_cType")
		ELSE
		Client_cType = 0
		END IF
		IF Request.Form("Client_cTel") <> "" THEN
		Client_cTel = Request.Form("Client_cTel")
		ELSE
		Client_cTel = 0
		END IF
		IF Request.Form("Client_cFax") <> "" THEN
		Client_cFax = Request.Form("Client_cFax")
		ELSE
		Client_cFax = 0
		END IF
		IF Request.Form("Client_cTrade") <> "" THEN
		Client_cTrade = Request.Form("Client_cTrade")
		ELSE
		Client_cTrade = 0
		END IF
		IF Request.Form("Client_cStrade") <> "" THEN
		Client_cStrade = Request.Form("Client_cStrade")
		ELSE
		Client_cStrade = 0
		END IF
		IF Request.Form("Client_cStart") <> "" THEN
		Client_cStart = Request.Form("Client_cStart")
		ELSE
		Client_cStart = 0
		END IF
		IF Request.Form("Client_cSource") <> "" THEN
		Client_cSource = Request.Form("Client_cSource")
		ELSE
		Client_cSource = 0
		END IF
		IF Request.Form("Client_cLinkman") <> "" THEN
		Client_cLinkman = Request.Form("Client_cLinkman")
		ELSE
		Client_cLinkman = 0
		END IF
		IF Request.Form("Client_cZhiwei") <> "" THEN
		Client_cZhiwei = Request.Form("Client_cZhiwei")
		ELSE
		Client_cZhiwei = 0
		END IF
		IF Request.Form("Client_cMobile") <> "" THEN
		Client_cMobile = Request.Form("Client_cMobile")
		ELSE
		Client_cMobile = 0
		END IF
		IF Request.Form("Client_cUser") <> "" THEN
		Client_cUser = Request.Form("Client_cUser")
		ELSE
		Client_cUser = 0
		END IF
		
		IF Request.Form("Client_cShare") <> "" THEN
		Client_cShare = Request.Form("Client_cShare")
		ELSE
		Client_cShare = 0
		END IF
		IF Request.Form("Client_cLastUpdated") <> "" THEN
		Client_cLastUpdated = Request.Form("Client_cLastUpdated")
		ELSE
		Client_cLastUpdated = 0
		END IF
		IF Request.Form("Client_cRNextTime") <> "" THEN
		Client_cRNextTime = Request.Form("Client_cRNextTime")
		ELSE
		Client_cRNextTime = 0
		END IF
		IF Request.Form("Client_cOEDate") <> "" THEN
		Client_cOEDate = Request.Form("Client_cOEDate")
		ELSE
		Client_cOEDate = 0
		END IF
		IF Request.Form("Client_cHEdate") <> "" THEN
		Client_cHEdate = Request.Form("Client_cHEdate")
		ELSE
		Client_cHEdate = 0
		END IF
		IF Request.Form("Client_cHMoney") <> "" THEN
		Client_cHMoney = Request.Form("Client_cHMoney")
		ELSE
		Client_cHMoney = 0
		END IF
		IF Request.Form("Client_cHOwed") <> "" THEN
		Client_cHOwed = Request.Form("Client_cHOwed")
		ELSE
		Client_cHOwed = 0
		END IF
		IF Request.Form("Client_cSNum") <> "" THEN
		Client_cSNum = Request.Form("Client_cSNum")
		ELSE
		Client_cSNum = 0
		END IF
		
		Dim TempStr
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "Client_cDate="& Client_cDate &"" & VbCrLf
		TempStr = TempStr & "Client_cCompany="& Client_cCompany &"" & VbCrLf
		TempStr = TempStr & "Client_cArea="& Client_cArea &"" & VbCrLf
		TempStr = TempStr & "Client_cSquare="& Client_cSquare &"" & VbCrLf
		TempStr = TempStr & "Client_cAddress="& Client_cAddress &"" & VbCrLf
		TempStr = TempStr & "Client_cType="& Client_cType &"" & VbCrLf
		TempStr = TempStr & "Client_cTel="& Client_cTel &"" & VbCrLf
		TempStr = TempStr & "Client_cFax="& Client_cFax &"" & VbCrLf
		TempStr = TempStr & "Client_cTrade="& Client_cTrade &"" & VbCrLf
		TempStr = TempStr & "Client_cStrade="& Client_cStrade &"" & VbCrLf
		TempStr = TempStr & "Client_cStart="& Client_cStart &"" & VbCrLf
		TempStr = TempStr & "Client_cSource="& Client_cSource &"" & VbCrLf
		TempStr = TempStr & "Client_cLinkman="& Client_cLinkman &"" & VbCrLf
		TempStr = TempStr & "Client_cZhiwei="& Client_cZhiwei &"" & VbCrLf
		TempStr = TempStr & "Client_cMobile="& Client_cMobile &"" & VbCrLf
		TempStr = TempStr & "Client_cUser="& Client_cUser &"" & VbCrLf
		TempStr = TempStr & "Client_cShare="& Client_cShare &"" & VbCrLf
		TempStr = TempStr & "Client_cLastUpdated="& Client_cLastUpdated &"" & VbCrLf
		TempStr = TempStr & "Client_cRNextTime="& Client_cRNextTime &"" & VbCrLf
		TempStr = TempStr & "Client_cOEDate="& Client_cOEDate &"" & VbCrLf
		TempStr = TempStr & "Client_cHEdate="& Client_cHEdate &"" & VbCrLf
		TempStr = TempStr & "Client_cHMoney="& Client_cHMoney &"" & VbCrLf
		TempStr = TempStr & "Client_cHOwed="& Client_cHOwed &"" & VbCrLf
		TempStr = TempStr & "Client_cSNum="& Client_cSNum &"" & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../data/Config/Show_Client.asp"
	
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="ClientAddMust" then '客户
%>
		<form name="Save" action="?action=Setting&sType=ClientAddSave" method="post"onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>必填字段</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Client_cCompany%></td>
								<td class="td_l_c"><input name="Must_Client_cCompany" type="checkbox" id="Must_Client_cCompany" value="1" <%if Must_Client_cCompany = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Client_cArea%></td>
								<td class="td_l_c"><input name="Must_Client_cArea" type="checkbox" id="Must_Client_cArea" value="1" <%if Must_Client_cArea = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Client_cSquare%></td>
								<td class="td_l_c"><input name="Must_Client_cSquare" type="checkbox" id="Must_Client_cSquare" value="1" <%if Must_Client_cSquare = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Client_cAddress%></td>
								<td class="td_l_c"><input name="Must_Client_cAddress" type="checkbox" id="Must_Client_cAddress" value="1" <%if Must_Client_cAddress = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Client_cZip%></td>
								<td class="td_l_c"><input name="Must_Client_cZip" type="checkbox" id="Must_Client_cZip" value="1" <%if Must_Client_cZip = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Client_cLinkman%></td>
								<td class="td_l_c"><input name="Must_Client_cLinkman" type="checkbox" id="Must_Client_cLinkman" value="1" <%if Must_Client_cLinkman = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Client_cZhiwei%></td>
								<td class="td_l_c"><input name="Must_Client_cZhiwei" type="checkbox" id="Must_Client_cZhiwei" value="1" <%if Must_Client_cZhiwei = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">8.</td>
								<td class="td_l_c"> <%=L_Client_cMobile%></td>
								<td class="td_l_c"><input name="Must_Client_cMobile" type="checkbox" id="Must_Client_cMobile" value="1" <%if Must_Client_cMobile = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">9.</td>
								<td class="td_l_c"> <%=L_Client_cTel%></td>
								<td class="td_l_c"><input name="Must_Client_cTel" type="checkbox" id="Must_Client_cTel" value="1" <%if Must_Client_cTel = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">10.</td>
								<td class="td_l_c"> <%=L_Client_cFax%></td>
								<td class="td_l_c"><input name="Must_Client_cFax" type="checkbox" id="Must_Client_cFax" value="1" <%if Must_Client_cFax = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">11.</td>
								<td class="td_l_c"> <%=L_Client_cHomepage%></td>
								<td class="td_l_c"><input name="Must_Client_cHomepage" type="checkbox" id="Must_Client_cHomepage" value="1" <%if Must_Client_cHomepage = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">12.</td>
								<td class="td_l_c"> <%=L_Client_cEmail%></td>
								<td class="td_l_c"><input name="Must_Client_cEmail" type="checkbox" id="Must_Client_cEmail" value="1" <%if Must_Client_cEmail = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">13.</td>
								<td class="td_l_c"> <%=L_Client_cTrade%></td>
								<td class="td_l_c"><input name="Must_Client_cTrade" type="checkbox" id="Must_Client_cTrade" value="1" <%if Must_Client_cTrade = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">14.</td>
								<td class="td_l_c"> <%=L_Client_cStrade%></td>
								<td class="td_l_c"><input name="Must_Client_cStrade" type="checkbox" id="Must_Client_cStrade" value="1" <%if Must_Client_cStrade = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">15.</td>
								<td class="td_l_c"> <%=L_Client_cType%></td>
								<td class="td_l_c"><input name="Must_Client_cType" type="checkbox" id="Must_Client_cType" value="1" <%if Must_Client_cType = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">16.</td>
								<td class="td_l_c"> <%=L_Client_cStart%></td>
								<td class="td_l_c"><input name="Must_Client_cStart" type="checkbox" id="Must_Client_cStart" value="1" <%if Must_Client_cStart = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">17.</td>
								<td class="td_l_c"> <%=L_Client_cSource%></td>
								<td class="td_l_c"><input name="Must_Client_cSource" type="checkbox" id="Must_Client_cSource" value="1" <%if Must_Client_cSource = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">18.</td>
								<td class="td_l_c"> <%=L_Client_cInfo%></td>
								<td class="td_l_c"><input name="Must_Client_cInfo" type="checkbox" id="Must_Client_cInfo" value="1" <%if Must_Client_cInfo = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">19.</td>
								<td class="td_l_c"> <%=L_Client_cBeizhu%></td>
								<td class="td_l_c"><input name="Must_Client_cBeizhu" type="checkbox" id="Must_Client_cBeizhu" value="1" <%if Must_Client_cBeizhu = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">20.</td>
								<td class="td_l_c"> 预留功能</td>
								<td class="td_l_c"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>隐藏字段</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Client_cTel%></td>
								<td class="td_l_c"><input name="Hidden_Client_cTel" type="checkbox" id="Hidden_Client_cTel" value="1" <%if Hidden_Client_cTel = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Client_cFax%></td>
								<td class="td_l_c"><input name="Hidden_Client_cFax" type="checkbox" id="Hidden_Client_cFax" value="1" <%if Hidden_Client_cFax = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Client_cHomepage%></td>
								<td class="td_l_c"><input name="Hidden_Client_cHomepage" type="checkbox" id="Hidden_Client_cHomepage" value="1" <%if Hidden_Client_cHomepage = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Client_cEmail%></td>
								<td class="td_l_c"><input name="Hidden_Client_cEmail" type="checkbox" id="Hidden_Client_cEmail" value="1" <%if Hidden_Client_cEmail = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Client_cStart%></td>
								<td class="td_l_c"><input name="Hidden_Client_cStart" type="checkbox" id="Hidden_Client_cStart" value="1" <%if Hidden_Client_cStart = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Client_cSource%></td>
								<td class="td_l_c"><input name="Hidden_Client_cSource" type="checkbox" id="Hidden_Client_cSource" value="1" <%if Hidden_Client_cSource = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Client_cInfo%></td>
								<td class="td_l_c"><input name="Hidden_Client_cInfo" type="checkbox" id="Hidden_Client_cInfo" value="1" <%if Hidden_Client_cInfo = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="ClientAddSave" then

		'更新必填字段		
		IF Request.Form("Must_Client_cCompany") <> "" THEN
		Must_Client_cCompany = Request.Form("Must_Client_cCompany")
		ELSE
		Must_Client_cCompany = 0
		END IF
		IF Request.Form("Must_Client_cArea") <> "" THEN
		Must_Client_cArea = Request.Form("Must_Client_cArea")
		ELSE
		Must_Client_cArea = 0
		END IF
		IF Request.Form("Must_Client_cSquare") <> "" THEN
		Must_Client_cSquare = Request.Form("Must_Client_cSquare")
		ELSE
		Must_Client_cSquare = 0
		END IF
		IF Request.Form("Must_Client_cAddress") <> "" THEN
		Must_Client_cAddress = Request.Form("Must_Client_cAddress")
		ELSE
		Must_Client_cAddress = 0
		END IF
		IF Request.Form("Must_Client_cZip") <> "" THEN
		Must_Client_cZip = Request.Form("Must_Client_cZip")
		ELSE
		Must_Client_cZip = 0
		END IF
		IF Request.Form("Must_Client_cLinkman") <> "" THEN
		Must_Client_cLinkman = Request.Form("Must_Client_cLinkman")
		ELSE
		Must_Client_cLinkman = 0
		END IF
		IF Request.Form("Must_Client_cZhiwei") <> "" THEN
		Must_Client_cZhiwei = Request.Form("Must_Client_cZhiwei")
		ELSE
		Must_Client_cZhiwei = 0
		END IF
		IF Request.Form("Must_Client_cMobile") <> "" THEN
		Must_Client_cMobile = Request.Form("Must_Client_cMobile")
		ELSE
		Must_Client_cMobile = 0
		END IF
		IF Request.Form("Must_Client_cTel") <> "" THEN
		Must_Client_cTel = Request.Form("Must_Client_cTel")
		ELSE
		Must_Client_cTel = 0
		END IF
		IF Request.Form("Must_Client_cFax") <> "" THEN
		Must_Client_cFax = Request.Form("Must_Client_cFax")
		ELSE
		Must_Client_cFax = 0
		END IF
		IF Request.Form("Must_Client_cHomepage") <> "" THEN
		Must_Client_cHomepage = Request.Form("Must_Client_cHomepage")
		ELSE
		Must_Client_cHomepage = 0
		END IF
		IF Request.Form("Must_Client_cEmail") <> "" THEN
		Must_Client_cEmail = Request.Form("Must_Client_cEmail")
		ELSE
		Must_Client_cEmail = 0
		END IF
		IF Request.Form("Must_Client_cTrade") <> "" THEN
		Must_Client_cTrade = Request.Form("Must_Client_cTrade")
		ELSE
		Must_Client_cTrade = 0
		END IF
		IF Request.Form("Must_Client_cStrade") <> "" THEN
		Must_Client_cStrade = Request.Form("Must_Client_cStrade")
		ELSE
		Must_Client_cStrade = 0
		END IF
		IF Request.Form("Must_Client_cType") <> "" THEN
		Must_Client_cType = Request.Form("Must_Client_cType")
		ELSE
		Must_Client_cType = 0
		END IF
		IF Request.Form("Must_Client_cStart") <> "" THEN
		Must_Client_cStart = Request.Form("Must_Client_cStart")
		ELSE
		Must_Client_cStart = 0
		END IF
		IF Request.Form("Must_Client_cSource") <> "" THEN
		Must_Client_cSource = Request.Form("Must_Client_cSource")
		ELSE
		Must_Client_cSource = 0
		END IF
		IF Request.Form("Must_Client_cInfo") <> "" THEN
		Must_Client_cInfo = Request.Form("Must_Client_cInfo")
		ELSE
		Must_Client_cInfo = 0
		END IF
		IF Request.Form("Must_Client_cBeizhu") <> "" THEN
		Must_Client_cBeizhu = Request.Form("Must_Client_cBeizhu")
		ELSE
		Must_Client_cBeizhu = 0
		END IF
		IF Request.Form("Must_Client_cShare") <> "" THEN
		Must_Client_cShare = Request.Form("Must_Client_cShare")
		ELSE
		Must_Client_cShare = 0
		END IF
		IF Request.Form("Hidden_Client_cTel") <> "" THEN
		Hidden_Client_cTel = Request.Form("Hidden_Client_cTel")
		ELSE
		Hidden_Client_cTel = 0
		END IF
		IF Request.Form("Hidden_Client_cFax") <> "" THEN
		Hidden_Client_cFax = Request.Form("Hidden_Client_cFax")
		ELSE
		Hidden_Client_cFax = 0
		END IF
		IF Request.Form("Hidden_Client_cHomepage") <> "" THEN
		Hidden_Client_cHomepage = Request.Form("Hidden_Client_cHomepage")
		ELSE
		Hidden_Client_cHomepage = 0
		END IF
		IF Request.Form("Hidden_Client_cEmail") <> "" THEN
		Hidden_Client_cEmail = Request.Form("Hidden_Client_cEmail")
		ELSE
		Hidden_Client_cEmail = 0
		END IF
		IF Request.Form("Hidden_Client_cStart") <> "" THEN
		Hidden_Client_cStart = Request.Form("Hidden_Client_cStart")
		ELSE
		Hidden_Client_cStart = 0
		END IF
		IF Request.Form("Hidden_Client_cSource") <> "" THEN
		Hidden_Client_cSource = Request.Form("Hidden_Client_cSource")
		ELSE
		Hidden_Client_cSource = 0
		END IF
		IF Request.Form("Hidden_Client_cInfo") <> "" THEN
		Hidden_Client_cInfo = Request.Form("Hidden_Client_cInfo")
		ELSE
		Hidden_Client_cInfo = 0
		END IF

		
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "Must_Client_cCompany="& Must_Client_cCompany &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cArea="& Must_Client_cArea &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cSquare="& Must_Client_cSquare &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cAddress="& Must_Client_cAddress &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cZip="& Must_Client_cZip &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cLinkman="& Must_Client_cLinkman &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cZhiwei="& Must_Client_cZhiwei &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cMobile="& Must_Client_cMobile &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cTel="& Must_Client_cTel &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cFax="& Must_Client_cFax &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cHomepage="& Must_Client_cHomepage &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cEmail="& Must_Client_cEmail &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cTrade="& Must_Client_cTrade &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cStrade="& Must_Client_cStrade &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cType="& Must_Client_cType &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cStart="& Must_Client_cStart &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cSource="& Must_Client_cSource &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cInfo="& Must_Client_cInfo &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cBeizhu="& Must_Client_cBeizhu &"" & VbCrLf
		TempStr = TempStr & "Must_Client_cShare="& Must_Client_cShare &"" & VbCrLf
		TempStr = TempStr & "Hidden_Client_cTel="& Hidden_Client_cTel &"" & VbCrLf
		TempStr = TempStr & "Hidden_Client_cFax="& Hidden_Client_cFax &"" & VbCrLf
		TempStr = TempStr & "Hidden_Client_cHomepage="& Hidden_Client_cHomepage &"" & VbCrLf
		TempStr = TempStr & "Hidden_Client_cEmail="& Hidden_Client_cEmail &"" & VbCrLf
		TempStr = TempStr & "Hidden_Client_cStart="& Hidden_Client_cStart &"" & VbCrLf
		TempStr = TempStr & "Hidden_Client_cSource="& Hidden_Client_cSource &"" & VbCrLf
		TempStr = TempStr & "Hidden_Client_cInfo="& Hidden_Client_cInfo &"" & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../data/Config/Must_Client.asp"
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="Linkmans" then '联系人
%>
		<form name="Save" action="?action=Setting&sType=SaveLinkmans" method="post"onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>列显示字段</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Linkmans_lName%></td>
								<td class="td_l_c"><input name="Linkmans_lName" type="checkbox" id="Linkmans_lName" value="1" <%if Linkmans_lName = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Linkmans_lSex%></td>
								<td class="td_l_c"><input name="Linkmans_lSex" type="checkbox" id="Linkmans_lSex" value="1" <%if Linkmans_lSex = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Linkmans_lZhiwei%></td>
								<td class="td_l_c"><input name="Linkmans_lZhiwei" type="checkbox" id="Linkmans_lZhiwei" value="1" <%if Linkmans_lZhiwei = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Linkmans_lMobile%></td>
								<td class="td_l_c"><input name="Linkmans_lMobile" type="checkbox" id="Linkmans_lMobile" value="1" <%if Linkmans_lMobile = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Linkmans_lTel%></td>
								<td class="td_l_c"><input name="Linkmans_lTel" type="checkbox" id="Linkmans_lTel" value="1" <%if Linkmans_lTel = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Linkmans_lEmail%></td>
								<td class="td_l_c"><input name="Linkmans_lEmail" type="checkbox" id="Linkmans_lEmail" value="1" <%if Linkmans_lEmail = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Linkmans_lQQ%></td>
								<td class="td_l_c"><input name="Linkmans_lQQ" type="checkbox" id="Linkmans_lQQ" value="1" <%if Linkmans_lQQ = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">8.</td>
								<td class="td_l_c"> <%=L_Linkmans_lMSN%></td>
								<td class="td_l_c"><input name="Linkmans_lMSN" type="checkbox" id="Linkmans_lMSN" value="1" <%if Linkmans_lMSN = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">9.</td>
								<td class="td_l_c"> <%=L_Linkmans_lALWW%></td>
								<td class="td_l_c"><input name="Linkmans_lALWW" type="checkbox" id="Linkmans_lALWW" value="1" <%if Linkmans_lALWW = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">10.</td>
								<td class="td_l_c"> <%=L_Linkmans_lBirthday%></td>
								<td class="td_l_c"><input name="Linkmans_lBirthday" type="checkbox" id="Linkmans_lBirthday" value="1" <%if Linkmans_lBirthday = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">11.</td>
								<td class="td_l_c"> <%=L_Linkmans_lContent%></td>
								<td class="td_l_c"><input name="Linkmans_lContent" type="checkbox" id="Linkmans_lContent" value="1" <%if Linkmans_lContent = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">12.</td>
								<td class="td_l_c"> <%=L_Linkmans_lTime%></td>
								<td class="td_l_c"><input name="Linkmans_lTime" type="checkbox" id="Linkmans_lTime" value="1" <%if Linkmans_lTime = 1 then %>checked<%end if%> ></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>必填项 ( 新增/修改 )</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Linkmans_lName%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lName" type="checkbox" id="Must_Linkmans_lName" value="1" <%if Must_Linkmans_lName = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Linkmans_lSex%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lSex" type="checkbox" id="Must_Linkmans_lSex" value="1" <%if Must_Linkmans_lSex = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Linkmans_lZhiwei%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lZhiwei" type="checkbox" id="Must_Linkmans_lZhiwei" value="1" <%if Must_Linkmans_lZhiwei = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Linkmans_lMobile%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lMobile" type="checkbox" id="Must_Linkmans_lMobile" value="1" <%if Must_Linkmans_lMobile = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Linkmans_lTel%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lTel" type="checkbox" id="Must_Linkmans_lTel" value="1" <%if Must_Linkmans_lTel = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Linkmans_lEmail%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lEmail" type="checkbox" id="Must_Linkmans_lEmail" value="1" <%if Must_Linkmans_lEmail = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Linkmans_lQQ%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lQQ" type="checkbox" id="Must_Linkmans_lQQ" value="1" <%if Must_Linkmans_lQQ = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">8.</td>
								<td class="td_l_c"> <%=L_Linkmans_lMSN%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lMSN" type="checkbox" id="Must_Linkmans_lMSN" value="1" <%if Must_Linkmans_lMSN = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">9.</td>
								<td class="td_l_c"> <%=L_Linkmans_lALWW%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lALWW" type="checkbox" id="Must_Linkmans_lALWW" value="1" <%if Must_Linkmans_lALWW = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">10.</td>
								<td class="td_l_c"> <%=L_Linkmans_lBirthday%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lBirthday" type="checkbox" id="Must_Linkmans_lBirthday" value="1" <%if Must_Linkmans_lBirthday = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">11.</td>
								<td class="td_l_c"> <%=L_Linkmans_lContent%></td>
								<td class="td_l_c"><input name="Must_Linkmans_lContent" type="checkbox" id="Must_Linkmans_lContent" value="1" <%if Must_Linkmans_lContent = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveLinkmans" then

		'更新列显示和必填字段
		IF Request.Form("Linkmans_lName") <> "" THEN
		Linkmans_lName = Request.Form("Linkmans_lName")
		ELSE
		Linkmans_lName = 0
		END IF
		IF Request.Form("Linkmans_lSex") <> "" THEN
		Linkmans_lSex = Request.Form("Linkmans_lSex")
		ELSE
		Linkmans_lSex = 0
		END IF
		IF Request.Form("Linkmans_lZhiwei") <> "" THEN
		Linkmans_lZhiwei = Request.Form("Linkmans_lZhiwei")
		ELSE
		Linkmans_lZhiwei = 0
		END IF
		IF Request.Form("Linkmans_lMobile") <> "" THEN
		Linkmans_lMobile = Request.Form("Linkmans_lMobile")
		ELSE
		Linkmans_lMobile = 0
		END IF
		IF Request.Form("Linkmans_lTel") <> "" THEN
		Linkmans_lTel = Request.Form("Linkmans_lTel")
		ELSE
		Linkmans_lTel = 0
		END IF
		IF Request.Form("Linkmans_lEmail") <> "" THEN
		Linkmans_lEmail = Request.Form("Linkmans_lEmail")
		ELSE
		Linkmans_lEmail = 0
		END IF
		IF Request.Form("Linkmans_lQQ") <> "" THEN
		Linkmans_lQQ = Request.Form("Linkmans_lQQ")
		ELSE
		Linkmans_lQQ = 0
		END IF
		IF Request.Form("Linkmans_lMSN") <> "" THEN
		Linkmans_lMSN = Request.Form("Linkmans_lMSN")
		ELSE
		Linkmans_lMSN = 0
		END IF
		IF Request.Form("Linkmans_lALWW") <> "" THEN
		Linkmans_lALWW = Request.Form("Linkmans_lALWW")
		ELSE
		Linkmans_lALWW = 0
		END IF
		IF Request.Form("Linkmans_lBirthday") <> "" THEN
		Linkmans_lBirthday = Request.Form("Linkmans_lBirthday")
		ELSE
		Linkmans_lBirthday = 0
		END IF
		IF Request.Form("Linkmans_lContent") <> "" THEN
		Linkmans_lContent = Request.Form("Linkmans_lContent")
		ELSE
		Linkmans_lContent = 0
		END IF
		IF Request.Form("Linkmans_lTime") <> "" THEN
		Linkmans_lTime = Request.Form("Linkmans_lTime")
		ELSE
		Linkmans_lTime = 0
		END IF
		IF Request.Form("Must_Linkmans_lName") <> "" THEN
		Must_Linkmans_lName = Request.Form("Must_Linkmans_lName")
		ELSE
		Must_Linkmans_lName = 0
		END IF
		IF Request.Form("Must_Linkmans_lSex") <> "" THEN
		Must_Linkmans_lSex = Request.Form("Must_Linkmans_lSex")
		ELSE
		Must_Linkmans_lSex = 0
		END IF
		IF Request.Form("Must_Linkmans_lZhiwei") <> "" THEN
		Must_Linkmans_lZhiwei = Request.Form("Must_Linkmans_lZhiwei")
		ELSE
		Must_Linkmans_lZhiwei = 0
		END IF
		IF Request.Form("Must_Linkmans_lMobile") <> "" THEN
		Must_Linkmans_lMobile = Request.Form("Must_Linkmans_lMobile")
		ELSE
		Must_Linkmans_lMobile = 0
		END IF
		IF Request.Form("Must_Linkmans_lTel") <> "" THEN
		Must_Linkmans_lTel = Request.Form("Must_Linkmans_lTel")
		ELSE
		Must_Linkmans_lTel = 0
		END IF
		IF Request.Form("Must_Linkmans_lEmail") <> "" THEN
		Must_Linkmans_lEmail = Request.Form("Must_Linkmans_lEmail")
		ELSE
		Must_Linkmans_lEmail = 0
		END IF
		IF Request.Form("Must_Linkmans_lQQ") <> "" THEN
		Must_Linkmans_lQQ = Request.Form("Must_Linkmans_lQQ")
		ELSE
		Must_Linkmans_lQQ = 0
		END IF
		IF Request.Form("Must_Linkmans_lMSN") <> "" THEN
		Must_Linkmans_lMSN = Request.Form("Must_Linkmans_lMSN")
		ELSE
		Must_Linkmans_lMSN = 0
		END IF
		IF Request.Form("Must_Linkmans_lALWW") <> "" THEN
		Must_Linkmans_lALWW = Request.Form("Must_Linkmans_lALWW")
		ELSE
		Must_Linkmans_lALWW = 0
		END IF
		IF Request.Form("Must_Linkmans_lBirthday") <> "" THEN
		Must_Linkmans_lBirthday = Request.Form("Must_Linkmans_lBirthday")
		ELSE
		Must_Linkmans_lBirthday = 0
		END IF
		IF Request.Form("Must_Linkmans_lContent") <> "" THEN
		Must_Linkmans_lContent = Request.Form("Must_Linkmans_lContent")
		ELSE
		Must_Linkmans_lContent = 0
		END IF

		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "Linkmans_lName="& Linkmans_lName &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lSex="& Linkmans_lSex &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lZhiwei="& Linkmans_lZhiwei &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lMobile="& Linkmans_lMobile &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lTel="& Linkmans_lTel &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lEmail="& Linkmans_lEmail &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lQQ="& Linkmans_lQQ &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lMSN="& Linkmans_lMSN &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lALWW="& Linkmans_lALWW &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lBirthday="& Linkmans_lBirthday &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lContent="& Linkmans_lContent &"" & VbCrLf
		TempStr = TempStr & "Linkmans_lTime="& Linkmans_lTime &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lName="& Must_Linkmans_lName &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lSex="& Must_Linkmans_lSex &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lZhiwei="& Must_Linkmans_lZhiwei &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lMobile="& Must_Linkmans_lMobile &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lTel="& Must_Linkmans_lTel &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lEmail="& Must_Linkmans_lEmail &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lQQ="& Must_Linkmans_lQQ &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lMSN="& Must_Linkmans_lMSN &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lALWW="& Must_Linkmans_lALWW &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lBirthday="& Must_Linkmans_lBirthday &"" & VbCrLf
		TempStr = TempStr & "Must_Linkmans_lContent="& Must_Linkmans_lContent &"" & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../data/Config/Show_Must_Linkmans.asp"
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="Records" then '跟单
%>
		<form name="Save" action="?action=Setting&sType=SaveRecords" method="post"onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>列显示字段</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Records_rType%></td>
								<td class="td_l_c"><input name="Records_rType" type="checkbox" id="Records_rType" value="1" <%if Records_rType = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Records_rState%></td>
								<td class="td_l_c"><input name="Records_rState" type="checkbox" id="Records_rState" value="1" <%if Records_rState = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Records_rLinkman%></td>
								<td class="td_l_c"><input name="Records_rLinkman" type="checkbox" id="Records_rLinkman" value="1" <%if Records_rLinkman = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Records_rNextTime%></td>
								<td class="td_l_c"><input name="Records_rNextTime" type="checkbox" id="Records_rNextTime" value="1" <%if Records_rNextTime = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Records_rContent%></td>
								<td class="td_l_c"><input name="Records_rContent" type="checkbox" id="Records_rContent" value="1" <%if Records_rContent = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Records_rUser%></td>
								<td class="td_l_c"><input name="Records_rUser" type="checkbox" id="Records_rUser" value="1" <%if Records_rUser = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Records_rTime%></td>
								<td class="td_l_c"><input name="Records_rTime" type="checkbox" id="Records_rTime" value="1" <%if Records_rTime = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"> </td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>必填项 ( 新增/修改 )</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Records_rType%></td>
								<td class="td_l_c"><input name="Must_Records_rType" type="checkbox" id="Must_Records_rType" value="1" <%if Must_Records_rType = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Records_rState%></td>
								<td class="td_l_c"><input name="Must_Records_rState" type="checkbox" id="Must_Records_rState" value="1" <%if Must_Records_rState = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Records_rLinkman%></td>
								<td class="td_l_c"><input name="Must_Records_rLinkman" type="checkbox" id="Must_Records_rLinkman" value="1" <%if Must_Records_rLinkman = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Records_rNextTime%></td>
								<td class="td_l_c"><input name="Must_Records_rNextTime" type="checkbox" id="Must_Records_rNextTime" value="1" <%if Must_Records_rNextTime = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Records_rContent%></td>
								<td class="td_l_c"><input name="Must_Records_rContent" type="checkbox" id="Must_Records_rContent" value="1" <%if Must_Records_rContent = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"> </td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"> </td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"> </td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveRecords" then

		'更新列显示和必填字段
		IF Request.Form("Records_rType") <> "" THEN
		 Records_rType= Request.Form("Records_rType")
		ELSE
		 Records_rType= 0
		END IF
		IF Request.Form("Records_rState") <> "" THEN
		 Records_rState= Request.Form("Records_rState")
		ELSE
		 Records_rState= 0
		END IF
		IF Request.Form("Records_rLinkman") <> "" THEN
		 Records_rLinkman= Request.Form("Records_rLinkman")
		ELSE
		 Records_rLinkman= 0
		END IF
		IF Request.Form("Records_rNextTime") <> "" THEN
		 Records_rNextTime= Request.Form("Records_rNextTime")
		ELSE
		 Records_rNextTime= 0
		END IF
		IF Request.Form("Records_rRemind") <> "" THEN
		 Records_rRemind= Request.Form("Records_rRemind")
		ELSE
		 Records_rRemind= 0
		END IF
		IF Request.Form("Records_rContent") <> "" THEN
		 Records_rContent= Request.Form("Records_rContent")
		ELSE
		 Records_rContent= 0
		END IF
		IF Request.Form("Records_rUser") <> "" THEN
		 Records_rUser= Request.Form("Records_rUser")
		ELSE
		 Records_rUser= 0
		END IF
		IF Request.Form("Records_rTime") <> "" THEN
		 Records_rTime= Request.Form("Records_rTime")
		ELSE
		 Records_rTime= 0
		END IF
		IF Request.Form("Must_Records_rType") <> "" THEN
		 Must_Records_rType= Request.Form("Must_Records_rType")
		ELSE
		 Must_Records_rType= 0
		END IF
		IF Request.Form("Must_Records_rState") <> "" THEN
		 Must_Records_rState= Request.Form("Must_Records_rState")
		ELSE
		 Must_Records_rState= 0
		END IF
		IF Request.Form("Must_Records_rLinkman") <> "" THEN
		 Must_Records_rLinkman= Request.Form("Must_Records_rLinkman")
		ELSE
		 Must_Records_rLinkman= 0
		END IF
		IF Request.Form("Must_Records_rNextTime") <> "" THEN
		 Must_Records_rNextTime= Request.Form("Must_Records_rNextTime")
		ELSE
		 Must_Records_rNextTime= 0
		END IF
		IF Request.Form("Must_Records_rRemind") <> "" THEN
		 Must_Records_rRemind= Request.Form("Must_Records_rRemind")
		ELSE
		 Must_Records_rRemind= 0
		END IF
		IF Request.Form("Must_Records_rContent") <> "" THEN
		 Must_Records_rContent= Request.Form("Must_Records_rContent")
		ELSE
		 Must_Records_rContent= 0
		END IF
		IF Request.Form("Must_Records_rUser") <> "" THEN
		 Must_Records_rUser= Request.Form("Must_Records_rUser")
		ELSE
		 Must_Records_rUser= 0
		END IF
		IF Request.Form("Must_Records_rTime") <> "" THEN
		 Must_Records_rTime= Request.Form("Must_Records_rTime")
		ELSE
		 Must_Records_rTime= 0
		END IF

		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "Records_rType="& Records_rType &"" & VbCrLf
		TempStr = TempStr & "Records_rState="& Records_rState &"" & VbCrLf
		TempStr = TempStr & "Records_rLinkman="& Records_rLinkman &"" & VbCrLf
		TempStr = TempStr & "Records_rNextTime="& Records_rNextTime &"" & VbCrLf
		TempStr = TempStr & "Records_rRemind="& Records_rRemind &"" & VbCrLf
		TempStr = TempStr & "Records_rContent="& Records_rContent &"" & VbCrLf
		TempStr = TempStr & "Records_rUser="& Records_rUser &"" & VbCrLf
		TempStr = TempStr & "Records_rTime="& Records_rTime &"" & VbCrLf
		TempStr = TempStr & "Must_Records_rType="& Must_Records_rType &"" & VbCrLf
		TempStr = TempStr & "Must_Records_rState="& Must_Records_rState &"" & VbCrLf
		TempStr = TempStr & "Must_Records_rLinkman="& Must_Records_rLinkman &"" & VbCrLf
		TempStr = TempStr & "Must_Records_rNextTime="& Must_Records_rNextTime &"" & VbCrLf
		TempStr = TempStr & "Must_Records_rRemind="& Must_Records_rRemind &"" & VbCrLf
		TempStr = TempStr & "Must_Records_rContent="& Must_Records_rContent &"" & VbCrLf
		TempStr = TempStr & "Must_Records_rUser="& Must_Records_rUser &"" & VbCrLf
		TempStr = TempStr & "Must_Records_rTime="& Must_Records_rTime &"" & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../data/Config/Show_Must_Records.asp"
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="Order" then '订单
%>
		<form name="Save" action="?action=Setting&sType=SaveOrder" method="post"onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>列显示字段</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Order_oLinkman%></td>
								<td class="td_l_c"><input name="Order_oLinkman" type="checkbox" id="Order_oLinkman" value="1" <%if Order_oLinkman = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Order_oSDate%></td>
								<td class="td_l_c"><input name="Order_oSDate" type="checkbox" id="Order_oSDate" value="1" <%if Order_oSDate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Order_oEDate%></td>
								<td class="td_l_c"><input name="Order_oEDate" type="checkbox" id="Order_oEDate" value="1" <%if Order_oEDate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Order_oDeposit%></td>
								<td class="td_l_c"><input name="Order_oDeposit" type="checkbox" id="Order_oDeposit" value="1" <%if Order_oDeposit = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Order_oState%></td>
								<td class="td_l_c"><input name="Order_oState" type="checkbox" id="Order_oState" value="1" <%if Order_oState = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Order_oContent%></td>
								<td class="td_l_c"><input name="Order_oContent" type="checkbox" id="Order_oContent" value="1" <%if Order_oContent = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Order_oUser%></td>
								<td class="td_l_c"><input name="Order_oUser" type="checkbox" id="Order_oUser" value="1" <%if Order_oUser = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">8.</td>
								<td class="td_l_c"> <%=L_Order_oTime%></td>
								<td class="td_l_c"><input name="Order_oTime" type="checkbox" id="Order_oTime" value="1" <%if Order_oTime = 1 then %>checked<%end if%> ></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>必填项 ( 新增/修改 )</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Order_oLinkman%></td>
								<td class="td_l_c"><input name="Must_Order_oLinkman" type="checkbox" id="Must_Order_oLinkman" value="1" <%if Must_Order_oLinkman = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Order_oSDate%></td>
								<td class="td_l_c"><input name="Must_Order_oSDate" type="checkbox" id="Must_Order_oSDate" value="1" <%if Must_Order_oSDate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Order_oEDate%></td>
								<td class="td_l_c"><input name="Must_Order_oEDate" type="checkbox" id="Must_Order_oEDate" value="1" <%if Must_Order_oEDate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Order_oDeposit%></td>
								<td class="td_l_c"><input name="Must_Order_oDeposit" type="checkbox" id="Must_Order_oDeposit" value="1" <%if Must_Order_oDeposit = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Order_oState%></td>
								<td class="td_l_c"><input name="Must_Order_oState" type="checkbox" id="Must_Order_oState" value="1" <%if Must_Order_oState = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Order_oContent%></td>
								<td class="td_l_c"><input name="Must_Order_oContent" type="checkbox" id="Must_Order_oContent" value="1" <%if Must_Order_oContent = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"> </td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"> </td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveOrder" then

		'更新列显示和必填字段
		IF Request.Form("Order_oLinkman") <> "" THEN
		 Order_oLinkman= Request.Form("Order_oLinkman")
		ELSE
		 Order_oLinkman= 0
		END IF
		IF Request.Form("Order_oSDate") <> "" THEN
		 Order_oSDate= Request.Form("Order_oSDate")
		ELSE
		 Order_oSDate= 0
		END IF
		IF Request.Form("Order_oEDate") <> "" THEN
		 Order_oEDate= Request.Form("Order_oEDate")
		ELSE
		 Order_oEDate= 0
		END IF
		IF Request.Form("Order_oDeposit") <> "" THEN
		 Order_oDeposit= Request.Form("Order_oDeposit")
		ELSE
		 Order_oDeposit= 0
		END IF
		IF Request.Form("Order_oMoney") <> "" THEN
		 Order_oMoney= Request.Form("Order_oMoney")
		ELSE
		 Order_oMoney= 0
		END IF
		IF Request.Form("Order_oState") <> "" THEN
		 Order_oState= Request.Form("Order_oState")
		ELSE
		 Order_oState= 0
		END IF
		IF Request.Form("Order_oContent") <> "" THEN
		 Order_oContent= Request.Form("Order_oContent")
		ELSE
		 Order_oContent= 0
		END IF
		IF Request.Form("Order_oUser") <> "" THEN
		 Order_oUser= Request.Form("Order_oUser")
		ELSE
		 Order_oUser= 0
		END IF
		IF Request.Form("Order_oTime") <> "" THEN
		 Order_oTime= Request.Form("Order_oTime")
		ELSE
		 Order_oTime= 0
		END IF
		IF Request.Form("Must_Order_oLinkman") <> "" THEN
		 Must_Order_oLinkman= Request.Form("Must_Order_oLinkman")
		ELSE
		 Must_Order_oLinkman= 0
		END IF
		IF Request.Form("Must_Order_oSDate") <> "" THEN
		 Must_Order_oSDate= Request.Form("Must_Order_oSDate")
		ELSE
		 Must_Order_oSDate= 0
		END IF
		IF Request.Form("Must_Order_oEDate") <> "" THEN
		 Must_Order_oEDate= Request.Form("Must_Order_oEDate")
		ELSE
		 Must_Order_oEDate= 0
		END IF
		IF Request.Form("Must_Order_oDeposit") <> "" THEN
		 Must_Order_oDeposit= Request.Form("Must_Order_oDeposit")
		ELSE
		 Must_Order_oDeposit= 0
		END IF
		IF Request.Form("Must_Order_oMoney") <> "" THEN
		 Must_Order_oMoney= Request.Form("Must_Order_oMoney")
		ELSE
		 Must_Order_oMoney= 0
		END IF
		IF Request.Form("Must_Order_oState") <> "" THEN
		 Must_Order_oState= Request.Form("Must_Order_oState")
		ELSE
		 Must_Order_oState= 0
		END IF
		IF Request.Form("Must_Order_oContent") <> "" THEN
		 Must_Order_oContent= Request.Form("Must_Order_oContent")
		ELSE
		 Must_Order_oContent= 0
		END IF

		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "Order_oLinkman="& Order_oLinkman &"" & VbCrLf
		TempStr = TempStr & "Order_oSDate="& Order_oSDate &"" & VbCrLf
		TempStr = TempStr & "Order_oEDate="& Order_oEDate &"" & VbCrLf
		TempStr = TempStr & "Order_oDeposit="& Order_oDeposit &"" & VbCrLf
		TempStr = TempStr & "Order_oMoney="& Order_oMoney &"" & VbCrLf
		TempStr = TempStr & "Order_oState="& Order_oState &"" & VbCrLf
		TempStr = TempStr & "Order_oContent="& Order_oContent &"" & VbCrLf
		TempStr = TempStr & "Order_oUser="& Order_oUser &"" & VbCrLf
		TempStr = TempStr & "Order_oTime="& Order_oTime &"" & VbCrLf
		TempStr = TempStr & "Must_Order_oLinkman="& Must_Order_oLinkman &"" & VbCrLf
		TempStr = TempStr & "Must_Order_oSDate="& Must_Order_oSDate &"" & VbCrLf
		TempStr = TempStr & "Must_Order_oEDate="& Must_Order_oEDate &"" & VbCrLf
		TempStr = TempStr & "Must_Order_oDeposit="& Must_Order_oDeposit &"" & VbCrLf
		TempStr = TempStr & "Must_Order_oMoney="& Must_Order_oMoney &"" & VbCrLf
		TempStr = TempStr & "Must_Order_oState="& Must_Order_oState &"" & VbCrLf
		TempStr = TempStr & "Must_Order_oContent="& Must_Order_oContent &"" & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../data/Config/Show_Must_Order.asp"
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="OrderProducts" then '订单产品
%>
		<form name="Save" action="?action=Setting&sType=SaveOrderProducts" method="post"onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>列显示字段</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Order_Products_oProTitle%></td>
								<td class="td_l_c"><input name="Order_Products_oProTitle" type="checkbox" id="Order_Products_oProTitle" value="1" <%if Order_Products_oProTitle = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Order_Products_oProItemA%></td>
								<td class="td_l_c"><input name="Order_Products_oProItemA" type="checkbox" id="Order_Products_oProItemA" value="1" <%if Order_Products_oProItemA = 1 then %>checked<%end if%> <%if pItemA = 0 then %>disabled<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Order_Products_oProItemB%></td>
								<td class="td_l_c"><input name="Order_Products_oProItemB" type="checkbox" id="Order_Products_oProItemB" value="1" <%if Order_Products_oProItemB = 1 then %>checked<%end if%> <%if pItemB = 0 then %>disabled<%end if%>  ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Order_Products_oProItemC%></td>
								<td class="td_l_c"><input name="Order_Products_oProItemC" type="checkbox" id="Order_Products_oProItemC" value="1" <%if Order_Products_oProItemC = 1 then %>checked<%end if%> <%if pItemC = 0 then %>disabled<%end if%>  ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Order_Products_oProItemD%></td>
								<td class="td_l_c"><input name="Order_Products_oProItemD" type="checkbox" id="Order_Products_oProItemD" value="1" <%if Order_Products_oProItemD = 1 then %>checked<%end if%> <%if pItemD = 0 then %>disabled<%end if%>  ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Order_Products_oProItemE%></td>
								<td class="td_l_c"><input name="Order_Products_oProItemE" type="checkbox" id="Order_Products_oProItemE" value="1" <%if Order_Products_oProItemE = 1 then %>checked<%end if%> <%if pItemE = 0 then %>disabled<%end if%>  ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Order_Products_oProPrice%></td>
								<td class="td_l_c"><input name="Order_Products_oProPrice" type="checkbox" id="Order_Products_oProPrice" value="1" <%if Order_Products_oProPrice = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">8.</td>
								<td class="td_l_c"> <%=L_Order_Products_oProNum%></td>
								<td class="td_l_c"><input name="Order_Products_oProNum" type="checkbox" id="Order_Products_oProNum" value="1" <%if Order_Products_oProNum = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">9.</td>
								<td class="td_l_c"> <%=L_Order_Products_oDiscount%></td>
								<td class="td_l_c"><input name="Order_Products_oDiscount" type="checkbox" id="Order_Products_oDiscount" value="1" <%if Order_Products_oDiscount = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">10.</td>
								<td class="td_l_c"> <%=L_Order_Products_oMoney%></td>
								<td class="td_l_c"><input name="Order_Products_oMoney" type="checkbox" id="Order_Products_oMoney" value="1" <%if Order_Products_oMoney = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">11.</td>
								<td class="td_l_c"> <%=L_Order_Products_oUser%></td>
								<td class="td_l_c"><input name="Order_Products_oUser" type="checkbox" id="Order_Products_oUser" value="1" <%if Order_Products_oUser = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">12.</td>
								<td class="td_l_c"> <%=L_Order_Products_oTime%></td>
								<td class="td_l_c"><input name="Order_Products_oTime" type="checkbox" id="Order_Products_oTime" value="1" <%if Order_Products_oTime = 1 then %>checked<%end if%> ></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>必填项 ( 新增/修改 )</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Order_Products_oProTitle%></td>
								<td class="td_l_c"><input name="Must_Order_Products_oProTitle" type="checkbox" id="Must_Order_Products_oProTitle" value="1" <%if Must_Order_Products_oProTitle = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Order_Products_oProNum%></td>
								<td class="td_l_c"><input name="Must_Order_Products_oProNum" type="checkbox" id="Must_Order_Products_oProNum" value="1" <%if Must_Order_Products_oProNum = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Order_Products_oDiscount%></td>
								<td class="td_l_c"><input name="Must_Order_Products_oDiscount" type="checkbox" id="Must_Order_Products_oDiscount" value="1" <%if Must_Order_Products_oDiscount = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Order_Products_oContent%></td>
								<td class="td_l_c"><input name="Must_Order_Products_oContent" type="checkbox" id="Must_Order_Products_oContent" value="1" <%if Must_Order_Products_oContent = 1 then %>checked<%end if%> ></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveOrderProducts" then

		'更新列显示和必填字段
		IF Request.Form("Order_Products_ProId") <> "" THEN
		 Order_Products_ProId= Request.Form("Order_Products_ProId")
		ELSE
		 Order_Products_ProId= 0
		END IF
		IF Request.Form("Order_Products_oProTitle") <> "" THEN
		 Order_Products_oProTitle= Request.Form("Order_Products_oProTitle")
		ELSE
		 Order_Products_oProTitle= 0
		END IF
		IF Request.Form("Order_Products_oProItemA") <> "" THEN
		 Order_Products_oProItemA= Request.Form("Order_Products_oProItemA")
		ELSE
		 Order_Products_oProItemA= 0
		END IF
		IF Request.Form("Order_Products_oProItemB") <> "" THEN
		 Order_Products_oProItemB= Request.Form("Order_Products_oProItemB")
		ELSE
		 Order_Products_oProItemB= 0
		END IF
		IF Request.Form("Order_Products_oProItemC") <> "" THEN
		 Order_Products_oProItemC= Request.Form("Order_Products_oProItemC")
		ELSE
		 Order_Products_oProItemC= 0
		END IF
		IF Request.Form("Order_Products_oProItemD") <> "" THEN
		 Order_Products_oProItemD= Request.Form("Order_Products_oProItemD")
		ELSE
		 Order_Products_oProItemD= 0
		END IF
		IF Request.Form("Order_Products_oProItemE") <> "" THEN
		 Order_Products_oProItemE= Request.Form("Order_Products_oProItemE")
		ELSE
		 Order_Products_oProItemE= 0
		END IF
		IF Request.Form("Order_Products_oProPrice") <> "" THEN
		 Order_Products_oProPrice= Request.Form("Order_Products_oProPrice")
		ELSE
		 Order_Products_oProPrice= 0
		END IF
		IF Request.Form("Order_Products_oProNum") <> "" THEN
		 Order_Products_oProNum= Request.Form("Order_Products_oProNum")
		ELSE
		 Order_Products_oProNum= 0
		END IF
		IF Request.Form("Order_Products_oProUnit") <> "" THEN
		 Order_Products_oProUnit= Request.Form("Order_Products_oProUnit")
		ELSE
		 Order_Products_oProUnit= 0
		END IF
		IF Request.Form("Order_Products_oDiscount") <> "" THEN
		 Order_Products_oDiscount= Request.Form("Order_Products_oDiscount")
		ELSE
		 Order_Products_oDiscount= 0
		END IF
		IF Request.Form("Order_Products_oMoney") <> "" THEN
		 Order_Products_oMoney= Request.Form("Order_Products_oMoney")
		ELSE
		 Order_Products_oMoney= 0
		END IF
		IF Request.Form("Order_Products_oUser") <> "" THEN
		 Order_Products_oUser= Request.Form("Order_Products_oUser")
		ELSE
		 Order_Products_oUser= 0
		END IF
		IF Request.Form("Order_Products_oTime") <> "" THEN
		 Order_Products_oTime= Request.Form("Order_Products_oTime")
		ELSE
		 Order_Products_oTime= 0
		END IF
		IF Request.Form("Must_Order_Products_oProTitle") <> "" THEN
		 Must_Order_Products_oProTitle= Request.Form("Must_Order_Products_oProTitle")
		ELSE
		 Must_Order_Products_oProTitle= 0
		END IF
		IF Request.Form("Must_Order_Products_oProNum") <> "" THEN
		 Must_Order_Products_oProNum= Request.Form("Must_Order_Products_oProNum")
		ELSE
		 Must_Order_Products_oProNum= 0
		END IF
		IF Request.Form("Must_Order_Products_oDiscount") <> "" THEN
		 Must_Order_Products_oDiscount= Request.Form("Must_Order_Products_oDiscount")
		ELSE
		 Must_Order_Products_oDiscount= 0
		END IF
		IF Request.Form("Must_Order_Products_oContent") <> "" THEN
		 Must_Order_Products_oContent= Request.Form("Must_Order_Products_oContent")
		ELSE
		 Must_Order_Products_oContent= 0
		END IF

		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "Order_Products_ProId="& Order_Products_ProId &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oProTitle="& Order_Products_oProTitle &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oProItemA="& Order_Products_oProItemA &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oProItemB="& Order_Products_oProItemB &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oProItemC="& Order_Products_oProItemC &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oProItemD="& Order_Products_oProItemD &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oProItemE="& Order_Products_oProItemE &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oProPrice="& Order_Products_oProPrice &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oProNum="& Order_Products_oProNum &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oProUnit="& Order_Products_oProUnit &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oDiscount="& Order_Products_oDiscount &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oMoney="& Order_Products_oMoney &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oUser="& Order_Products_oUser &"" & VbCrLf
		TempStr = TempStr & "Order_Products_oTime="& Order_Products_oTime &"" & VbCrLf
		TempStr = TempStr & "Must_Order_Products_oProTitle="& Must_Order_Products_oProTitle &"" & VbCrLf
		TempStr = TempStr & "Must_Order_Products_oProNum="& Must_Order_Products_oProNum &"" & VbCrLf
		TempStr = TempStr & "Must_Order_Products_oDiscount="& Must_Order_Products_oDiscount &"" & VbCrLf
		TempStr = TempStr & "Must_Order_Products_oContent="& Must_Order_Products_oContent &"" & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../data/Config/Show_Must_Order_Products.asp"
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="Hetong" then '合同
%>
		<form name="Save" action="?action=Setting&sType=SaveHetong" method="post"onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>列显示字段</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Hetong_hNum%></td>
								<td class="td_l_c"><input name="Hetong_hNum" type="checkbox" id="Hetong_hNum" value="1" <%if Hetong_hNum = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Hetong_oId%></td>
								<td class="td_l_c"><input name="Hetong_oId" type="checkbox" id="Hetong_oId" value="1" <%if Hetong_oId = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Hetong_hSdate%></td>
								<td class="td_l_c"><input name="Hetong_hSdate" type="checkbox" id="Hetong_hSdate" value="1" <%if Hetong_hSdate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Hetong_hEdate%></td>
								<td class="td_l_c"><input name="Hetong_hEdate" type="checkbox" id="Hetong_hEdate" value="1" <%if Hetong_hEdate = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Hetong_hType%></td>
								<td class="td_l_c"><input name="Hetong_hType" type="checkbox" id="Hetong_hType" value="1" <%if Hetong_hType = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Hetong_hMoney%></td>
								<td class="td_l_c"><input name="Hetong_hMoney" type="checkbox" id="Hetong_hMoney" value="1" <%if Hetong_hMoney = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Hetong_hRevenue%></td>
								<td class="td_l_c"><input name="Hetong_hRevenue" type="checkbox" id="Hetong_hRevenue" value="1" <%if Hetong_hRevenue = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">8.</td>
								<td class="td_l_c"> <%=L_Hetong_hOwed%></td>
								<td class="td_l_c"><input name="Hetong_hOwed" type="checkbox" id="Hetong_hOwed" value="1" <%if Hetong_hOwed = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">9.</td>
								<td class="td_l_c"> <%=L_Hetong_hInvoice%></td>
								<td class="td_l_c"><input name="Hetong_hInvoice" type="checkbox" id="Hetong_hInvoice" value="1" <%if Hetong_hInvoice = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">10.</td>
								<td class="td_l_c"> <%=L_Hetong_hTax%></td>
								<td class="td_l_c"><input name="Hetong_hTax" type="checkbox" id="Hetong_hTax" value="1" <%if Hetong_hTax = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">11.</td>
								<td class="td_l_c"> <%=L_Hetong_hAudit%></td>
								<td class="td_l_c"><input name="Hetong_hAudit" type="checkbox" id="Hetong_hAudit" value="1" <%if Hetong_hAudit = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">12.</td>
								<td class="td_l_c"> <%=L_Hetong_hAuditTime%></td>
								<td class="td_l_c"><input name="Hetong_hAuditTime" type="checkbox" id="Hetong_hAuditTime" value="1" <%if Hetong_hAuditTime = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">13.</td>
								<td class="td_l_c"> <%=L_Hetong_hUser%></td>
								<td class="td_l_c"><input name="Hetong_hUser" type="checkbox" id="Hetong_hUser" value="1" <%if Hetong_hUser = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">14.</td>
								<td class="td_l_c"> <%=L_Hetong_hTime%></td>
								<td class="td_l_c"><input name="Hetong_hTime" type="checkbox" id="Hetong_hTime" value="1" <%if Hetong_hTime = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c " colspan=6></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>必填项 ( 新增/修改 )</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Hetong_oId%></td>
								<td class="td_l_c"><input name="Must_Hetong_oId" type="checkbox" id="Must_Hetong_oId" value="1" <%if Must_Hetong_oId = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Hetong_hSdate%></td>
								<td class="td_l_c"><input name="Must_Hetong_hSdate" type="checkbox" id="Must_Hetong_hSdate" value="1" <%if Must_Hetong_hSdate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Hetong_hEdate%></td>
								<td class="td_l_c"><input name="Must_Hetong_hEdate" type="checkbox" id="Must_Hetong_hEdate" value="1" <%if Must_Hetong_hEdate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Hetong_hType%></td>
								<td class="td_l_c"><input name="Must_Hetong_hType" type="checkbox" id="Must_Hetong_hType" value="1" <%if Must_Hetong_hType = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Hetong_hMoney%></td>
								<td class="td_l_c"><input name="Must_Hetong_hMoney" type="checkbox" id="Must_Hetong_hMoney" value="1" <%if Must_Hetong_hMoney = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Hetong_hRevenue%></td>
								<td class="td_l_c"><input name="Must_Hetong_hRevenue" type="checkbox" id="Must_Hetong_hRevenue" value="1" <%if Must_Hetong_hRevenue = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Hetong_hInvoice%></td>
								<td class="td_l_c"><input name="Must_Hetong_hInvoice" type="checkbox" id="Must_Hetong_hInvoice" value="1" <%if Must_Hetong_hInvoice = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">8.</td>
								<td class="td_l_c"> <%=L_Hetong_hContent%></td>
								<td class="td_l_c"><input name="Must_Hetong_hContent" type="checkbox" id="Must_Hetong_hContent" value="1" <%if Must_Hetong_hContent = 1 then %>checked<%end if%> ></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveHetong" then

		'更新列显示和必填字段
		IF Request.Form("Hetong_hNum") <> "" THEN
		 Contract_hNum= Request.Form("Hetong_hNum")
		ELSE
		 Contract_hNum= 0
		END IF
		IF Request.Form("Hetong_oId") <> "" THEN
		 Contract_oId= Request.Form("Hetong_oId")
		ELSE
		 Contract_oId= 0
		END IF
		IF Request.Form("Hetong_hSdate") <> "" THEN
		 Contract_hSdate= Request.Form("Hetong_hSdate")
		ELSE
		 Contract_hSdate= 0
		END IF
		IF Request.Form("Hetong_hEdate") <> "" THEN
		 Contract_hEdate= Request.Form("Hetong_hEdate")
		ELSE
		 Contract_hEdate= 0
		END IF
		IF Request.Form("Hetong_hType") <> "" THEN
		 Contract_hType= Request.Form("Hetong_hType")
		ELSE
		 Contract_hType= 0
		END IF
		IF Request.Form("Hetong_hMoney") <> "" THEN
		 Contract_hMoney= Request.Form("Hetong_hMoney")
		ELSE
		 Contract_hMoney= 0
		END IF
		IF Request.Form("Hetong_hRevenue") <> "" THEN
		 Contract_hRevenue= Request.Form("Hetong_hRevenue")
		ELSE
		 Contract_hRevenue= 0
		END IF
		IF Request.Form("Hetong_hOwed") <> "" THEN
		 Contract_hOwed= Request.Form("Hetong_hOwed")
		ELSE
		 Contract_hOwed= 0
		END IF
		IF Request.Form("Hetong_hInvoice") <> "" THEN
		 Contract_hInvoice= Request.Form("Hetong_hInvoice")
		ELSE
		 Contract_hInvoice= 0
		END IF
		IF Request.Form("Hetong_hTax") <> "" THEN
		 Contract_hTax= Request.Form("Hetong_hTax")
		ELSE
		 Contract_hTax= 0
		END IF
		IF Request.Form("Hetong_hAudit") <> "" THEN
		 Contract_hAudit= Request.Form("Hetong_hAudit")
		ELSE
		 Contract_hAudit= 0
		END IF
		IF Request.Form("Hetong_hAuditTime") <> "" THEN
		 Contract_hAuditTime= Request.Form("Hetong_hAuditTime")
		ELSE
		 Contract_hAuditTime= 0
		END IF
		IF Request.Form("Hetong_hUser") <> "" THEN
		 Contract_hUser= Request.Form("Hetong_hUser")
		ELSE
		 Contract_hUser= 0
		END IF
		IF Request.Form("Hetong_hTime") <> "" THEN
		 Contract_hTime= Request.Form("Hetong_hTime")
		ELSE
		 Contract_hTime= 0
		END IF
		IF Request.Form("Must_Hetong_oId") <> "" THEN
		 Must_Hetong_oId= Request.Form("Must_Hetong_oId")
		ELSE
		 Must_Hetong_oId= 0
		END IF
		IF Request.Form("Must_Hetong_hSdate") <> "" THEN
		 Must_Hetong_hSdate= Request.Form("Must_Hetong_hSdate")
		ELSE
		 Must_Hetong_hSdate= 0
		END IF
		IF Request.Form("Must_Hetong_hEdate") <> "" THEN
		 Must_Hetong_hEdate= Request.Form("Must_Hetong_hEdate")
		ELSE
		 Must_Hetong_hEdate= 0
		END IF
		IF Request.Form("Must_Hetong_hType") <> "" THEN
		 Must_Hetong_hType= Request.Form("Must_Hetong_hType")
		ELSE
		 Must_Hetong_hType= 0
		END IF
		IF Request.Form("Must_Hetong_hMoney") <> "" THEN
		 Must_Hetong_hMoney= Request.Form("Must_Hetong_hMoney")
		ELSE
		 Must_Hetong_hMoney= 0
		END IF
		IF Request.Form("Must_Hetong_hRevenue") <> "" THEN
		 Must_Hetong_hRevenue= Request.Form("Must_Hetong_hRevenue")
		ELSE
		 Must_Hetong_hRevenue= 0
		END IF
		IF Request.Form("Must_Hetong_hInvoice") <> "" THEN
		 Must_Hetong_hInvoice= Request.Form("Must_Hetong_hInvoice")
		ELSE
		 Must_Hetong_hInvoice= 0
		END IF
		IF Request.Form("Must_Hetong_hContent") <> "" THEN
		 Must_Hetong_hContent= Request.Form("Must_Hetong_hContent")
		ELSE
		 Must_Hetong_hContent= 0
		END IF

		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "Hetong_hNum="& Contract_hNum &"" & VbCrLf
		TempStr = TempStr & "Hetong_oId="& Contract_oId &"" & VbCrLf
		TempStr = TempStr & "Hetong_hSdate="& Contract_hSdate &"" & VbCrLf
		TempStr = TempStr & "Hetong_hEdate="& Contract_hEdate &"" & VbCrLf
		TempStr = TempStr & "Hetong_hType="& Contract_hType &"" & VbCrLf
		TempStr = TempStr & "Hetong_hMoney="& Contract_hMoney &"" & VbCrLf
		TempStr = TempStr & "Hetong_hRevenue="& Contract_hRevenue &"" & VbCrLf
		TempStr = TempStr & "Hetong_hOwed="& Contract_hOwed &"" & VbCrLf
		TempStr = TempStr & "Hetong_hInvoice="& Contract_hInvoice &"" & VbCrLf
		TempStr = TempStr & "Hetong_hTax="& Contract_hTax &"" & VbCrLf
		TempStr = TempStr & "Hetong_hAudit="& Contract_hAudit &"" & VbCrLf
		TempStr = TempStr & "Hetong_hAuditTime="& Contract_hAuditTime &"" & VbCrLf
		TempStr = TempStr & "Hetong_hUser="& Contract_hUser &"" & VbCrLf
		TempStr = TempStr & "Hetong_hTime="& Contract_hTime &"" & VbCrLf
		
		TempStr = TempStr & "Must_Hetong_oId="& Must_Hetong_oId &"" & VbCrLf
		TempStr = TempStr & "Must_Hetong_hSdate="& Must_Hetong_hSdate &"" & VbCrLf
		TempStr = TempStr & "Must_Hetong_hEdate="& Must_Hetong_hEdate &"" & VbCrLf
		TempStr = TempStr & "Must_Hetong_hType="& Must_Hetong_hType &"" & VbCrLf
		TempStr = TempStr & "Must_Hetong_hMoney="& Must_Hetong_hMoney &"" & VbCrLf
		TempStr = TempStr & "Must_Hetong_hRevenue="& Must_Hetong_hRevenue &"" & VbCrLf
		TempStr = TempStr & "Must_Hetong_hInvoice="& Must_Hetong_hInvoice &"" & VbCrLf
		TempStr = TempStr & "Must_Hetong_hContent="& Must_Hetong_hContent &"" & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../data/Config/Show_Must_Hetong.asp"
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="Service" then '售后
%>
		<form name="Save" action="?action=Setting&sType=SaveService" method="post"onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>列显示字段</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Service_ProId%></td>
								<td class="td_l_c"><input name="Service_ProId" type="checkbox" id="Service_ProId" value="1" <%if Service_ProId = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Service_sTitle%></td>
								<td class="td_l_c"><input name="Service_sTitle" type="checkbox" id="Service_sTitle" value="1" <%if Service_sTitle = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Service_sLinkman%></td>
								<td class="td_l_c"><input name="Service_sLinkman" type="checkbox" id="Service_sLinkman" value="1" <%if Service_sLinkman = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Service_sType%></td>
								<td class="td_l_c"><input name="Service_sType" type="checkbox" id="Service_sType" value="1" <%if Service_sType = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Service_sSDate%></td>
								<td class="td_l_c"><input name="Service_sSDate" type="checkbox" id="Service_sSDate" value="1" <%if Service_sSDate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Service_sContent%></td>
								<td class="td_l_c"><input name="Service_sContent" type="checkbox" id="Service_sContent" value="1" <%if Service_sContent = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Service_sSolve%></td>
								<td class="td_l_c"><input name="Service_sSolve" type="checkbox" id="Service_sSolve" value="1" <%if Service_sSolve = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">8.</td>
								<td class="td_l_c"> <%=L_Service_sInfo%></td>
								<td class="td_l_c"><input name="Service_sInfo" type="checkbox" id="Service_sInfo" value="1" <%if Service_sInfo = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">9.</td>
								<td class="td_l_c"> <%=L_Service_sUser%></td>
								<td class="td_l_c"><input name="Service_sUser" type="checkbox" id="Service_sUser" value="1" <%if Service_sUser = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">10.</td>
								<td class="td_l_c"> <%=L_Service_sTime%></td>
								<td class="td_l_c"><input name="Service_sTime" type="checkbox" id="Service_sTime" value="1" <%if Service_sTime = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"></td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>必填项 ( 新增/修改 )</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Service_ProId%></td>
								<td class="td_l_c"><input name="Must_Service_ProId" type="checkbox" id="Must_Service_ProId" value="1" <%if Must_Service_ProId = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Service_sTitle%></td>
								<td class="td_l_c"><input name="Must_Service_sTitle" type="checkbox" id="Must_Service_sTitle" value="1" <%if Must_Service_sTitle = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Service_sType%></td>
								<td class="td_l_c"><input name="Must_Service_sType" type="checkbox" id="Must_Service_sType" value="1" <%if Must_Service_sType = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Service_sLinkman%></td>
								<td class="td_l_c"><input name="Must_Service_sLinkman" type="checkbox" id="Must_Service_sLinkman" value="1" <%if Must_Service_sLinkman = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Service_sSDate%></td>
								<td class="td_l_c"><input name="Must_Service_sSDate" type="checkbox" id="Must_Service_sSDate" value="1" <%if Must_Service_sSDate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Service_sContent%></td>
								<td class="td_l_c"><input name="Must_Service_sContent" type="checkbox" id="Must_Service_sContent" value="1" <%if Must_Service_sContent = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"></td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"> </td>
								<td class="td_l_c"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveService" then

		'更新列显示和必填字段
		IF Request.Form("Service_ProId") <> "" THEN
		 Service_ProId= Request.Form("Service_ProId")
		ELSE
		 Service_ProId= 0
		END IF
		IF Request.Form("Service_sTitle") <> "" THEN
		 Service_sTitle= Request.Form("Service_sTitle")
		ELSE
		 Service_sTitle= 0
		END IF
		IF Request.Form("Service_sLinkman") <> "" THEN
		 Service_sLinkman= Request.Form("Service_sLinkman")
		ELSE
		 Service_sLinkman= 0
		END IF
		IF Request.Form("Service_sType") <> "" THEN
		 Service_sType= Request.Form("Service_sType")
		ELSE
		 Service_sType= 0
		END IF
		IF Request.Form("Service_sSDate") <> "" THEN
		 Service_sSDate= Request.Form("Service_sSDate")
		ELSE
		 Service_sSDate= 0
		END IF
		IF Request.Form("Service_sContent") <> "" THEN
		 Service_sContent= Request.Form("Service_sContent")
		ELSE
		 Service_sContent= 0
		END IF
		IF Request.Form("Service_sSolve") <> "" THEN
		 Service_sSolve= Request.Form("Service_sSolve")
		ELSE
		 Service_sSolve= 0
		END IF
		IF Request.Form("Service_sInfo") <> "" THEN
		 Service_sInfo= Request.Form("Service_sInfo")
		ELSE
		 Service_sInfo= 0
		END IF
		IF Request.Form("Must_Service_ProId") <> "" THEN
		 Must_Service_ProId= Request.Form("Must_Service_ProId")
		ELSE
		 Must_Service_ProId= 0
		END IF
		IF Request.Form("Must_Service_sTitle") <> "" THEN
		 Must_Service_sTitle= Request.Form("Must_Service_sTitle")
		ELSE
		 Must_Service_sTitle= 0
		END IF
		IF Request.Form("Must_Service_sType") <> "" THEN
		 Must_Service_sType= Request.Form("Must_Service_sType")
		ELSE
		 Must_Service_sType= 0
		END IF
		IF Request.Form("Must_Service_sLinkman") <> "" THEN
		 Must_Service_sLinkman= Request.Form("Must_Service_sLinkman")
		ELSE
		 Must_Service_sLinkman= 0
		END IF
		IF Request.Form("Must_Service_sSDate") <> "" THEN
		 Must_Service_sSDate= Request.Form("Must_Service_sSDate")
		ELSE
		 Must_Service_sSDate= 0
		END IF
		IF Request.Form("Must_Service_sContent") <> "" THEN
		 Must_Service_sContent= Request.Form("Must_Service_sContent")
		ELSE
		 Must_Service_sContent= 0
		END IF

		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "Service_ProId="& Service_ProId &"" & VbCrLf
		TempStr = TempStr & "Service_sTitle="& Service_sTitle &"" & VbCrLf
		TempStr = TempStr & "Service_sLinkman="& Service_sLinkman &"" & VbCrLf
		TempStr = TempStr & "Service_sType="& Service_sType &"" & VbCrLf
		TempStr = TempStr & "Service_sSDate="& Service_sSDate &"" & VbCrLf
		TempStr = TempStr & "Service_sContent="& Service_sContent &"" & VbCrLf
		TempStr = TempStr & "Service_sSolve="& Service_sSolve &"" & VbCrLf
		TempStr = TempStr & "Service_sInfo="& Service_sInfo &"" & VbCrLf
		TempStr = TempStr & "Must_Service_ProId="& Must_Service_ProId &"" & VbCrLf
		TempStr = TempStr & "Must_Service_sTitle="& Must_Service_sTitle &"" & VbCrLf
		TempStr = TempStr & "Must_Service_sType="& Must_Service_sType &"" & VbCrLf
		TempStr = TempStr & "Must_Service_sLinkman="& Must_Service_sLinkman &"" & VbCrLf
		TempStr = TempStr & "Must_Service_sSDate="& Must_Service_sSDate &"" & VbCrLf
		TempStr = TempStr & "Must_Service_sContent="& Must_Service_sContent &"" & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../data/Config/Show_Must_Service.asp"

		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="Expense" then '费用
%>
		<form name="Save" action="?action=Setting&sType=SaveExpense" method="post"onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>列显示字段</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Expense_eDate%></td>
								<td class="td_l_c"><input name="Expense_eDate" type="checkbox" id="Expense_eDate" value="1" <%if Expense_eDate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Expense_eOutIn%></td>
								<td class="td_l_c"><input name="Expense_eOutIn" type="checkbox" id="Expense_eOutIn" value="1" <%if Expense_eOutIn = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Expense_eType%></td>
								<td class="td_l_c"><input name="Expense_eType" type="checkbox" id="Expense_eType" value="1" <%if Expense_eType = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Expense_eMoney%></td>
								<td class="td_l_c"><input name="Expense_eMoney" type="checkbox" id="Expense_eMoney" value="1" <%if Expense_eMoney = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_l_c"> <%=L_Expense_eContent%></td>
								<td class="td_l_c"><input name="Expense_eContent" type="checkbox" id="Expense_eContent" value="1" <%if Expense_eContent = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">6.</td>
								<td class="td_l_c"> <%=L_Expense_eUser%></td>
								<td class="td_l_c"><input name="Expense_eUser" type="checkbox" id="Expense_eUser" value="1" <%if Expense_eUser = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">7.</td>
								<td class="td_l_c"> <%=L_Expense_eTime%></td>
								<td class="td_l_c"><input name="Expense_eTime" type="checkbox" id="Expense_eTime" value="1" <%if Expense_eTime = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title"></td>
								<td class="td_l_c"></td>
								<td class="td_l_c"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<col width="50" /><col /><col width="50" /><col width="50" /><col /><col width="50" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>必填项 ( 新增/修改 )</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_c"><B>字段名称</B></td>
								<td class="td_l_c"><B>显示</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_l_c"> <%=L_Expense_eDate%></td>
								<td class="td_l_c"><input name="Must_Expense_eDate" type="checkbox" id="Must_Expense_eDate" value="1" <%if Must_Expense_eDate = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">2.</td>
								<td class="td_l_c"> <%=L_Expense_eType%></td>
								<td class="td_l_c"><input name="Must_Expense_eType" type="checkbox" id="Must_Expense_eType" value="1" <%if Must_Expense_eType = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">3.</td>
								<td class="td_l_c"> <%=L_Expense_eMoney%></td>
								<td class="td_l_c"><input name="Must_Expense_eMoney" type="checkbox" id="Must_Expense_eMoney" value="1" <%if Must_Expense_eMoney = 1 then %>checked<%end if%> ></td>
								<td class="td_l_c title">4.</td>
								<td class="td_l_c"> <%=L_Expense_eContent%></td>
								<td class="td_l_c"><input name="Must_Expense_eContent" type="checkbox" id="Must_Expense_eContent" value="1" <%if Must_Expense_eContent = 1 then %>checked<%end if%> ></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveExpense" then

		'更新列显示和必填字段
		IF Request.Form("Expense_eDate") <> "" THEN
		 Expense_eDate= Request.Form("Expense_eDate")
		ELSE
		 Expense_eDate= 0
		END IF
		IF Request.Form("Expense_eOutIn") <> "" THEN
		 Expense_eOutIn= Request.Form("Expense_eOutIn")
		ELSE
		 Expense_eOutIn= 0
		END IF
		IF Request.Form("Expense_eType") <> "" THEN
		 Expense_eType= Request.Form("Expense_eType")
		ELSE
		 Expense_eType= 0
		END IF
		IF Request.Form("Expense_eMoney") <> "" THEN
		 Expense_eMoney= Request.Form("Expense_eMoney")
		ELSE
		 Expense_eMoney= 0
		END IF
		IF Request.Form("Expense_eContent") <> "" THEN
		 Expense_eContent= Request.Form("Expense_eContent")
		ELSE
		 Expense_eContent= 0
		END IF
		IF Request.Form("Expense_eUser") <> "" THEN
		 Expense_eUser= Request.Form("Expense_eUser")
		ELSE
		 Expense_eUser= 0
		END IF
		IF Request.Form("Expense_eTime") <> "" THEN
		 Expense_eTime= Request.Form("Expense_eTime")
		ELSE
		 Expense_eTime= 0
		END IF
		IF Request.Form("Must_Expense_eDate") <> "" THEN
		 Must_Expense_eDate= Request.Form("Must_Expense_eDate")
		ELSE
		 Must_Expense_eDate= 0
		END IF
		IF Request.Form("Must_Expense_eType") <> "" THEN
		 Must_Expense_eType= Request.Form("Must_Expense_eType")
		ELSE
		 Must_Expense_eType= 0
		END IF
		IF Request.Form("Must_Expense_eMoney") <> "" THEN
		 Must_Expense_eMoney= Request.Form("Must_Expense_eMoney")
		ELSE
		 Must_Expense_eMoney= 0
		END IF
		IF Request.Form("Must_Expense_eContent") <> "" THEN
		 Must_Expense_eContent= Request.Form("Must_Expense_eContent")
		ELSE
		 Must_Expense_eContent= 0
		END IF

		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "Expense_eDate="& Expense_eDate &"" & VbCrLf
		TempStr = TempStr & "Expense_eOutIn="& Expense_eOutIn &"" & VbCrLf
		TempStr = TempStr & "Expense_eType="& Expense_eType &"" & VbCrLf
		TempStr = TempStr & "Expense_eMoney="& Expense_eMoney &"" & VbCrLf
		TempStr = TempStr & "Expense_eContent="& Expense_eContent &"" & VbCrLf
		TempStr = TempStr & "Expense_eUser="& Expense_eUser &"" & VbCrLf
		TempStr = TempStr & "Expense_eTime="& Expense_eTime &"" & VbCrLf
		TempStr = TempStr & "Must_Expense_eDate="& Must_Expense_eDate &"" & VbCrLf
		TempStr = TempStr & "Must_Expense_eType="& Must_Expense_eType &"" & VbCrLf
		TempStr = TempStr & "Must_Expense_eMoney="& Must_Expense_eMoney &"" & VbCrLf
		TempStr = TempStr & "Must_Expense_eContent="& Must_Expense_eContent &"" & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../data/Config/Show_Must_Expense.asp"
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="Products" then
%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="SaveOffer" action="?action=Setting&sType=SaveProducts" method="post"onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="50" /><col /><col width="80" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="3"><B>产品属性配置</B></td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c"><B>序号</B></td>
								<td class="td_l_l"><B>属性名称 <span class="info_help help01" onmouseover="tip.start(this)" tips="可在【系统设置】-【语言包】-【订单记录】中修改">&nbsp;</span></B></td>
								<td class="td_l_c"><B>是否使用</B></td>
							</tr>
							<tr> 
								<td class="td_l_c title">1.</td>
								<td class="td_r_l"> <%=L_Order_Products_oProItemA%></td>
								<td class="td_l_c"><input name="ShowsA" type="checkbox" id="ShowsA" value="1" <%if pItemA = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">2.</td>
								<td class="td_r_l"> <%=L_Order_Products_oProItemB%></td>
								<td class="td_l_c"><input name="ShowsB" type="checkbox" id="ShowsB" value="1" <%if pItemB = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">3.</td>
								<td class="td_r_l"> <%=L_Order_Products_oProItemC%></td>
								<td class="td_l_c"><input name="ShowsC" type="checkbox" id="ShowsC" value="1" <%if pItemC = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">4.</td>
								<td class="td_r_l"> <%=L_Order_Products_oProItemD%></td>
								<td class="td_l_c"><input name="ShowsD" type="checkbox" id="ShowsD" value="1" <%if pItemD = 1 then %>checked<%end if%> ></td>
							</tr>
							<tr> 
								<td class="td_l_c title">5.</td>
								<td class="td_r_l"> <%=L_Order_Products_oProItemE%></td>
								<td class="td_l_c"><input name="ShowsE" type="checkbox" id="ShowsE" value="1" <%if pItemE = 1 then %>checked<%end if%> ></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10"> 
						<div style="float:left;padding:10px 0;">
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
						</div>
					</td> 
				</tr>
				</form>
			</table>
<%
elseif sType="SaveProducts" then
		IF Request.Form("ShowsA") <> "" THEN
		 ShowsA= Request.Form("ShowsA")
		ELSE
		 ShowsA= 0
		END IF
		IF Request.Form("ShowsB") <> "" THEN
		 ShowsB= Request.Form("ShowsB")
		ELSE
		 ShowsB= 0
		END IF
		IF Request.Form("ShowsC") <> "" THEN
		 ShowsC= Request.Form("ShowsC")
		ELSE
		 ShowsC= 0
		END IF
		IF Request.Form("ShowsD") <> "" THEN
		 ShowsD= Request.Form("ShowsD")
		ELSE
		 ShowsD= 0
		END IF
		IF Request.Form("ShowsE") <> "" THEN
		 ShowsE= Request.Form("ShowsE")
		ELSE
		 ShowsE= 0
		END IF

		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "pItemA="& ShowsA &"" & VbCrLf
		TempStr = TempStr & "pItemB="& ShowsB &"" & VbCrLf
		TempStr = TempStr & "pItemC="& ShowsC &"" & VbCrLf
		TempStr = TempStr & "pItemD="& ShowsD &"" & VbCrLf
		TempStr = TempStr & "pItemE="& ShowsE &"" & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
		ADODB_SaveToFile TempStr,"../data/Config/Show_Must_Products.asp"
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

end if
%>


<%
End Sub

%>

<%


Sub ADODB_SaveToFile(ByVal strBody,ByVal File)
	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If
	
	If FSFlag = 1 Then
		Set WriteFile = fs.CreateTextFile(Server.MapPath(File),True)
		WriteFile.Write strBody
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请使用FTP等功能，将<font color=Red >data/config.asp</font>文件内容替换成框中内容"
			Err.Clear
			Set objStream = Noting
			Exit Sub
		End If
		With objStream
			.Type = 2
			.Open
			.Charset = "GB2312"
			.Position = objStream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	End If
End Sub
%>

<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>