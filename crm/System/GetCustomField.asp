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

Sub CustomField() '自定义字段
if sType="Add" then '添加大类
%>
		<form name="Save" action="GetCustomField.asp?action=CustomField&sType=SaveAdd" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>自定义内容</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">数据表</td>
								<td class="td_l_l"><select name="cTable" id="cTable" ><option value="">请选择</option><option value="Client">客户档案</option><option value="Records">跟单记录</option><option value="Order">订单记录</option><option value="Hetong">合同记录</option><option value="Service">售后记录</option></select>　
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">显示名</td>
								<td class="td_l_l"><input name="cTitle" type="text" id="cTitle" class="int" size="20" /> 例 : 开户行</td>
							</tr>
							<tr>
								<td class="td_l_r title">字段名</td>
								<td class="td_l_l"><input name="cName" type="text" id="cName" class="int" size="20" /> 例 : BANK</td>
							</tr>
							<tr>
								<td class="td_l_r title">字段类型</td>
								<td class="td_l_l"><select name="cType" id="cType" ><option value="">请选择</option><option value="text">文本</option><option value="time">时间日期</option><option value="select">下拉框</option><option value="checkbox">多选框</option><option value="radio">单选框</option></select>　
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">表单长度</td>
								<td class="td_l_l"><input name="cWidth" type="text" id="cWidth" class="int" size="20" /> 单位 : PX</td>
							</tr>
							<tr>
								<td class="td_l_r title">备注</td>
								<td class="td_l_l" style="padding:5px 10px;"><textarea name="cContent" rows="4" id="cContent" class="int" style="height:50px;width:98%;"></textarea></td>
							</tr>
							<tr>
								<td class="td_l_r title">列显示</td>
								<td class="td_l_l"><label style="width:50px;"><input name="cList" type="radio" value="1" />&nbsp;是</label>&nbsp;&nbsp; <label><input name="cList" type="radio" value="0" checked="checked" />&nbsp;否</label></td>
							</tr>
							<tr>
								<td class="td_l_r title">启用</td>
								<td class="td_l_l"><label style="width:50px;"><input name="cYn" type="radio" value="1" checked="checked" />&nbsp;是</label>&nbsp;&nbsp; <label> <input name="cYn" type="radio" value="0" />&nbsp;否</label></td>
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
		cTable = Request.Form("cTable")
		cTitle = Request.Form("cTitle")
		cName = Request.Form("cName")
		cTypeS = Request.Form("cType")
		cWidth = Request.Form("cWidth")
		cContent = Request.Form("cContent")
		cList = Request.Form("cList")
		cYn = Request.Form("cYn")
		If cName = "" Then
			Response.Write("<script>location.href='GetCustomField.asp?action=CustomField&sType=Add&tipinfo=不能为空';</script>")
			Exit Sub
		End If
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [CustomField] ",conn,3,2
		rs.AddNew
		rs("cTable") = cTable
		rs("cTitle") = cTitle
		rs("cName") = cName
		rs("cType") = cTypeS
		rs("cWidth") = cWidth
		rs("cContent") = cContent
		rs("cList") = cList
		rs("cYn") = cYn
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

elseif sType="Edit" then '修改大类
	Id = Request("Id")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [CustomField] Where Id = " & Id,conn,1,1
	cTable = rs("cTable") 
	cTitle = rs("cTitle") 
	cName  = rs("cName") 
	cTypeS = rs("cType") 
	cWidth = rs("cWidth") 
	cContent = rs("cContent") 
	cList = rs("cList") 
	cYn = rs("cYn") 
	rs.Close
	Set rs = Nothing
%>
		<form name="Save" action="GetCustomField.asp?action=CustomField&sType=SaveEdit" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>自定义内容</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">数据表</td>
								<td class="td_l_l"><select name="cTable" id="cTable" ><option value="">请选择</option><option value="Client">客户档案</option><option value="Records">跟单记录</option><option value="Order">订单记录</option><option value="Hetong">合同记录</option><option value="Service">售后记录</option><option value="Expense">费用记录</option></select>　
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">显示名</td>
								<td class="td_l_l"><input name="cTitle" type="text" id="cTitle" class="int" size="20" value="<%=cTitle%>" /> 例 : 开户行</td>
							</tr>
							<tr>
								<td class="td_l_r title">字段名</td>
								<td class="td_l_l"><input name="cName" type="text" id="cName" class="int" size="20" value="<%=cName%>" /> 例 : BANK</td>
							</tr>
							<tr>
								<td class="td_l_r title">字段类型</td>
								<td class="td_l_l"><select name="cType" id="cType" ><option value="">请选择</option><option value="text">文本</option><option value="time">时间日期</option><option value="select">下拉框</option><option value="checkbox">多选框</option><option value="radio">单选框</option></select>　
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">表单长度</td>
								<td class="td_l_l"><input name="cWidth" type="text" id="cWidth" class="int" size="20" value="<%=cWidth%>" /> 单位 : PX</td>
							</tr>
							<tr>
								<td class="td_l_r title" style="line-height:25px;">默认值<BR>半角逗号分割</td>
								<td class="td_l_l" style="padding:5px 10px;"><textarea name="cContent" rows="4" id="cContent" class="int" style="height:50px;width:98%;"><% = cContent %></textarea></td>
							</tr>
							<tr>
								<td class="td_l_r title">列显示</td>
								<td class="td_l_l"><label style="width:50px;"><input name="cList" type="radio" value="1" <%if cList = "1" then%>checked<%end if%> />&nbsp;是</label>&nbsp;&nbsp; <label><input name="cList" type="radio" value="0" checked="checked" <%if cList = "0" then%>checked<%end if%> />&nbsp;否</label></td>
							</tr>
							<tr>
								<td class="td_l_r title">启用</td>
								<td class="td_l_l"><label style="width:50px;"><input name="cYn" type="radio" value="1" checked="checked" <%if cYn = "1" then%>checked<%end if%> />&nbsp;是</label>&nbsp;&nbsp; <label> <input name="cYn" type="radio" value="0" <%if cYn = "0" then%>checked<%end if%> />&nbsp;否</label></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input name="Id" type="hidden" id="Id" value="<% = Id %>">
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<script language="JavaScript">
<!--
for(var i=0;i<document.all.cTable.options.length;i++){
    if(document.all.cTable.options[i].value == "<% = cTable %>"){
    document.all.cTable.options[i].selected = true;}}

for(var i=0;i<document.all.cType.options.length;i++){
    if(document.all.cType.options[i].value == "<% = cTypeS %>"){
    document.all.cType.options[i].selected = true;}}
-->
</script>
<%
elseif sType="SaveEdit" then
		Id = Request.Form("Id")
		cTable = Request.Form("cTable")
		cTitle = Request.Form("cTitle")
		cName = Request.Form("cName")
		cTypeS = Request.Form("cType")
		cWidth = Request.Form("cWidth")
		cContent = Request.Form("cContent")
		cList = Request.Form("cList")
		cYn = Request.Form("cYn")
		If cName = "" Then
			Response.Write("<script>location.href='GetCustomField.asp?action=CustomField&sType=Edit&aId="&aId&"&tipinfo=不能为空';</script>")
			Exit Sub
		End If
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [CustomField] Where cName = '"&cName&"' And Id <> " & Id,conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetCustomField.asp?action=CustomField&sType=Edit&Id="&Id&"&tipinfo=已存在！';</script>")
		Response.End()
		End If
		rs.Close

    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select * From [CustomField] where Id="&Id&" ",conn,3,2
		rs("cTable") = cTable
		rs("cTitle") = cTitle
		rs("cName") = cName
		rs("cType") = cTypeS
		rs("cWidth") = cWidth
		rs("cContent") = cContent
		rs("cList") = cList
		rs("cYn") = cYn
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
end if
End Sub
%>

<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>