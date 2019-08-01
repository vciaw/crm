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

<body oncontextmenu=self.event.returnValue=false> 
<style>body{padding:35px 0 48px;}</style>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：系统管理 > 产品管理</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
			<input type="button" class="button_top_set" value=" " title="设置" onclick='Setting_Products()' style="cursor:pointer" />
        </td>
	</tr>
</table>
<script>function Setting_Products() {$.dialog.open('GetUpdate.asp?action=Setting&sType=Products', {title: '用户设置', width: 500, height: 320,fixed: true}); };</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
			<div class="MenuboxS">
				 <ul>
					<li class="hover"><span><a href="?">产品列表</a></span></li>
					<li class=""><span><a href="#" onclick='Products_ClassList()' style="cursor:pointer">分类管理</a></span></li>
					<li class=""><span><a href="#" onclick='Products_InfoAdd()' style="cursor:pointer">添加产品</a></span></li>
				</ul>
			</div>
		</td>
	</tr>
</table>
<%
action = Trim(Request("action"))
sType = Trim(Request("sType"))

Select Case action
Case "Products"
    Call Products()
Case ""
    Call Products()
End Select

Sub Products() '产品数据
	if sType="" or sType="List" then
%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10 "> 
						
<script>function Products_ClassList() {$.dialog.open('GetProduct.asp?action=Products&sType=ClassList', {title: '分类管理', width: 600,height: 450, fixed: true}); };</script>
<script>function Products_InfoAdd() {$.dialog.open('GetProduct.asp?action=Products&sType=InfoAdd', {title: '添加产品', width: 600,height: 400, fixed: true}); };</script>

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<td width="100" class="td_l_c">产品大类</td>
								<td width="100" class="td_l_c">产品小类</td>
								<td class="td_l_l">产品名称</td>
								<%if EasyCrm.getNewItem("Field_Name","Field","'pItemA'","Shows") = 1 then%>
								<td class="td_l_c"><%=L_Order_Products_oProItemA%></td>
								<%end if%>
								<%if EasyCrm.getNewItem("Field_Name","Field","'pItemB'","Shows") = 1 then%>
								<td class="td_l_c"><%=L_Order_Products_oProItemB%></td>
								<%end if%>
								<%if EasyCrm.getNewItem("Field_Name","Field","'pItemC'","Shows") = 1 then%>
								<td class="td_l_c"><%=L_Order_Products_oProItemC%></td>
								<%end if%>
								<%if EasyCrm.getNewItem("Field_Name","Field","'pItemD'","Shows") = 1 then%>
								<td class="td_l_c"><%=L_Order_Products_oProItemD%></td>
								<%end if%>
								<%if EasyCrm.getNewItem("Field_Name","Field","'pItemE'","Shows") = 1 then%>
								<td class="td_l_c"><%=L_Order_Products_oProItemE%></td>
								<%end if%>
								<td width="80" class="td_l_c">单价</td>
								<td width="90" class="td_l_c">管理</td>
							</tr>
	<%
	PN = CLng(ABS(Request("PN")))
	If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
	intPageSize = DataPageSize
	pagenum = intPageSize*(PN-1)
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	IF PN=1 THEN
	rs.Open "Select top "&intPageSize&" * From [Products] where 1=1 "&sql&" Order By Id desc ",conn,1,1 
	ELSE
	rs.Open "Select top "&intPageSize&" * From [Products] where 1=1 "&sql&" and Id < ( SELECT Min(Id) FROM ( SELECT TOP "&pagenum&" Id FROM [Products] where 1=1 "&sql&" ORDER BY Id desc ) AS T ) Order By Id desc ",conn,1,1
	END IF
	SQLstr="Select count(Id) As RecordSum From [Products] where 1=1 "&sql&""
	
	Set Rsstr=conn.Execute("Select count(Id) As RecordSum From [Products] where 1=1 "&sql&"",1,1)
	TotalRecords=Rsstr("RecordSum") 
	if Int(TotalRecords/intPageSize)=TotalRecords/intPageSize then
	TotalPages=TotalRecords/intPageSize
	else
	TotalPages=Int(TotalRecords/intPageSize)+1
	end if
	Rsstr.Close:Set Rsstr=Nothing
	If PN > TotalPages Then PN = TotalPages
	Do While Not rs.BOF And Not rs.EOF
	%>
							<tr class="tr">
								<td class="td_l_c"><%=rs("pBigClass")%></td>
								<td class="td_l_c"><%=rs("pSmallClass")%></td>
								<td class="td_l_l"><%=rs("pTitle")%></td>
								<%if EasyCrm.getNewItem("Field_Name","Field","'pItemA'","Shows") = 1 then%>
								<td class="td_l_c"><%=rs("pItemA")%></td>
								<%end if%>
								<%if EasyCrm.getNewItem("Field_Name","Field","'pItemB'","Shows") = 1 then%>
								<td class="td_l_c"><%=rs("pItemB")%></td>
								<%end if%>
								<%if EasyCrm.getNewItem("Field_Name","Field","'pItemC'","Shows") = 1 then%>
								<td class="td_l_c"><%=rs("pItemC")%></td>
								<%end if%>
								<%if EasyCrm.getNewItem("Field_Name","Field","'pItemD'","Shows") = 1 then%>
								<td class="td_l_c"><%=rs("pItemD")%></td>
								<%end if%>
								<%if EasyCrm.getNewItem("Field_Name","Field","'pItemE'","Shows") = 1 then%>
								<td class="td_l_c"><%=rs("pItemE")%></td>
								<%end if%>
								<td class="td_l_c"><%=rs("pUprice")%></td>
								<td class="td_l_c"><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Products_InfoEdit<%=rs("Id")%>()' style="cursor:pointer" /> <input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Products_InfoDel<%=rs("Id")%>()' style="cursor:pointer" /></td>
							</tr>
							<script>function Products_InfoEdit<%=rs("Id")%>() {$.dialog.open('GetProduct.asp?action=Products&sType=InfoEdit&Id=<%=rs("Id")%>', {title: '编辑产品', width: 600,height: 400, fixed: true}); };</script>
							<script>function Products_InfoDel<%=rs("Id")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () { art.dialog.open('?action=Products&sType=Del&Id=<%=rs("id")%>');return false;},cancel: true }); };</script>
	<%
	rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	%>
							
						</table>
					</td>
				</tr>	
			</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<%=EasyCrm.pagelist("?action=Products&sType=List", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<%
	elseif sType="Del" then
		id= Trim(Request("id"))
		
		set rs=conn.execute("DELETE FROM [Products] where id = "&id&" ")
		Set rs = Nothing
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	end if
end Sub
%>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>