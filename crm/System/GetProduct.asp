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
						<script>function Products_BClass_Add() {$.dialog.open('GetProduct.asp?action=Products&sType=BigClassAdd', {title: '新增产品大类', width: 400,height: 145, fixed: true}); };</script>
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
									<td class="td_l_r title"><input type="button" class="button_info_add" value=" " title="添加小类"  onclick='Products_SClass_Add<%=rs("pClassId")%>()' style="cursor:pointer" /><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Products_BClass_Edit<%=rs("pClassId")%>()' style="cursor:pointer" /><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onClick=window.location.href="GetProduct.asp?action=Products&sType=ProductsClassDel&pClassId=<%=rs("pClassId")%>" /></td>
								</tr>
						<script>function Products_BClass_Edit<%=rs("pClassId")%>() {$.dialog.open('GetProduct.asp?action=Products&sType=BigClassEdit&pClassId=<%=rs("pClassId")%>', {title: '编辑产品大类', width: 400,height: 145, fixed: true}); };</script>
						<script>function Products_SClass_Add<%=rs("pClassId")%>() {$.dialog.open('GetProduct.asp?action=Products&sType=SmallClassAdd&pClassFid=<%=rs("pClassId")%>', {title: '添加产品小类', width: 400,height: 180, fixed: true}); };</script>
								<%	'子分类列表
										Set rss = Server.CreateObject("ADODB.Recordset")
										rss.Open "Select * From [ProductClass] where pClassFid ='" & rs("pClassId") & "' ",conn,3,1
										Do While Not rss.BOF And Not rss.EOF
								%>
										<tr class="tr">
											<td class="td_l_l" style="padding-left:30px;">┗━━ <a  href="javascript:void(0)" onclick='Products_SClass_Edit<%=rss("pClassId")%>()' title='修改' style="cursor:pointer"><%=rss("pClassname")%></a></td>
											<td class="td_l_r"><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>" onclick='Products_SClass_Edit<%=rss("pClassId")%>()' style="cursor:pointer" /><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onClick=window.location.href="GetProduct.asp?action=Products&sType=ProductsClassDel&pClassId=<%=rss("pClassId")%>" /></td>
										</tr>
						<script>function Products_SClass_Edit<%=rss("pClassId")%>() {$.dialog.open('GetProduct.asp?action=Products&sType=SmallClassEdit&pClassId=<%=rss("pClassId")%>', {title: '编辑产品小类', width: 400,height: 180, fixed: true}); };</script>
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
		<form name="Save" action="GetProduct.asp?action=Products&sType=SaveBigClassAdd" method="post">
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
			Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=BigClassAdd&tipinfo=产品分类名不能为空';</script>")
			Exit Sub
		End If
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [Productclass] Where pClassname = '"&pClassname&"' ",conn,3,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=BigClassAdd&tipinfo=已存在！';</script>")
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
		<form name="Save" action="GetProduct.asp?action=Products&sType=SaveBigClassEdit" method="post">
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
			Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=BigClassEdit&pClassId="&pClassid&"&tipinfo=产品分类名不能为空';</script>")
			Exit Sub
		End If
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [Productclass] Where pClassname = '"&pClassname&"' And pClassid <> " & pClassid,conn,3,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=BigClassEdit&pClassId="&pClassid&"&tipinfo=已存在！';</script>")
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
		<form name="Save" action="GetProduct.asp?action=Products&sType=SaveSmallClassAdd" method="post" onSubmit="return CheckInput();">
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
        Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=SmallClassAdd&pClassFid="&pClassFid&"&tipinfo=产品大类不能为空';</script>")
		Exit Sub
	End If
	If pClassname = "" Then
        Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=SmallClassAdd&pClassFid="&pClassFid&"&tipinfo=产品小类不能为空';</script>")
		Exit Sub
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Productclass] Where pClassFid='"&pClassFid&"' and pClassname = '" & pClassname & "'",conn,3,2
	If rs.RecordCount > 0 Then
        Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=SmallClassAdd&pClassFid="&pClassFid&"&tipinfo=已存在！' ;</script>")
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
		<form name="Save" action="GetProduct.asp?action=Products&sType=SaveSmallClassEdit" method="post">
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
        Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=SmallClassEdit&pClassId="&pClassid&"&tipinfo=产品大类不能为空' ;</script>")
		Exit Sub
	End If
	If pClassname = "" Then
        Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=SmallClassEdit&pClassId="&pClassid&"&tipinfo=产品小类不能为空' ;</script>")
		Exit Sub
	End If
	
	Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [Productclass] Where pClassFid = '"&pClassFid&"' And pClassname = '"&pClassname&"' And pClassid <> "&pClassid,conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=SmallClassEdit&pClassId="&pClassid&"&tipinfo=已存在！' ;</script>")
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
        Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=ClassList&tipinfo=有子分类，禁止删除！';</script>")
	else
		Set rss = Server.CreateObject("ADODB.Recordset")
		rss.Open "Select * From [Productclass] Where pClassId = " & pClassId,conn,3,2
		If rss.RecordCount > 0 Then
			rss.Delete
			rss.Update
		End If
		rss.Close
		Set rss = Nothing
		Response.Redirect("GetProduct.asp?action=Products&sType=ClassList")
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
		<form name="Save" action="GetProduct.asp?action=Products&sType=SaveInfoAdd" method="post">
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
		'	Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=InfoAdd&tipinfo=产品名称已存在，请重新输入！';</script>")
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
		<form name="Save" action="GetProduct.asp?action=Products&sType=SaveInfoEdit" method="post">
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
		'	Response.Write("<script>location.href='GetProduct.asp?action=Products&sType=InfoAdd&tipinfo=产品名称已存在，请重新输入！';</script>")
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
%>

<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>