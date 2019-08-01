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


Sub Level()
	if tipinfo<>"" then
		Response.Write("<script>art.dialog({title: 'Error',time: 1.5,icon: 'warning',content: '"&tipinfo&"'});</script>")
	end if
%>

<style>body{padding:0 0 55px 0;}</style>
	<script language="JavaScript">
	<!-- 必填项提示
	function CheckInput()
	{
		if(document.getElementById('lId').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '权限值不能为空！'});document.getElementById('lId').focus();return false;}
		if(document.getElementById('lName').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '角色名称不能为空！'});document.getElementById('lName').focus();return false;}
	}
	-->
	</script>
			<script language=javascript> 
			//全选/反选
			function selectall(id){ //用id区分  
			var tform=document.forms['Level'];  
			for(var i=0;i<tform.length;i++){  
			var e=tform.elements[i];  
			if(e.type=="checkbox" && e.id==id) e.checked=!e.checked;  } }
			</script> 
<%
if sType="Add" then
%>
		<form name="Save" id="Level" action="GetLevel.asp?action=Level&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>新增角色 </B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> 权限值</td>
								<td class="td_l_l"><input name="lId" type="text" class="int" id="lId" size="10" maxlength="2" onkeyup='this.value=this.value.replace(/\D/gi,"")' >  <span class="info_help help01">限：数字 2 - 8</span></td>
								<td class="td_l_r title"><font color="#FF0000">*</font> 角色名称</td>
								<td class="td_l_l"><input name="lName" type="text" class="int" id="lName" size="30" maxlength="16" > </td>
							</tr>
						</table>
					</td>
				</tr>
				
				
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
		
						<fieldset style="padding:10px;">
							<legend>&nbsp;<B style="font-size:14px;">全局权限</B>&nbsp;</legend>
							<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
								<col width="4%"><col width="11%"><col width="5%"> 
								<col width="4%"><col width="11%"><col width="5%"> 
								<col width="4%"><col width="11%"><col width="5%"> 
								<col width="4%"><col width="11%"><col width="5%"> 
								<col width="4%"><col width="11%"><col width="5%"> 
								<tr> 
									<td class="td_l_c fontbold">01.</td>
									<td class="td_l_r title">系统登录</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag1" value="1" <%if mid(qxflag, 1, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">02.</td>
									<td class="td_l_r title">客户管理</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag2" value="1" <%if mid(qxflag, 2, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">03.</td>
									<td class="td_l_r title">办公OA</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag3" value="1" <%if mid(qxflag, 3, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">04.</td>
									<td class="td_l_r title">功能插件</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag4" value="1" <%if mid(qxflag, 4, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">05.</td>
									<td class="td_l_r title" onmouseover="tip.start(this)" tips="超管权限！"><font color=red><B>系统设置</B></font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag5" value="1" <%if mid(qxflag, 5, 1) = "1" then Response.Write "checked"%>></td>
								</tr>
								<tr> 							
									<td class="td_l_c fontbold">06.</td>
									<td class="td_l_r title" onmouseover="tip.start(this)" tips="客户档案下拉框字段新增项目"><font color=red>新增下拉框</font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag6" value="1" <%if mid(qxflag, 6, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">07.</td>
									<td class="td_l_r title" onmouseover="tip.start(this)" tips="导出客户数据，有一定风险！"><font color=red>导出Excel</font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag7" value="1" <%if mid(qxflag, 7, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">08.</td>
									<td class="td_l_r title" onmouseover="tip.start(this)" tips="导入客户数据，有一定风险！"><font color=red>导入Excel</font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag8" value="1" <%if mid(qxflag, 8, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">09.</td>
									<td class="td_l_r title" onmouseover="tip.start(this)" tips="批量更新数据，有一定风险！"><font color=red>批量操作</font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag9" value="1" <%if mid(qxflag, 9, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">10.</td>
									<td class="td_l_r title"><font color=red>客户共享</font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag10" value="1" <%if mid(qxflag, 10, 1) = "1" then Response.Write "checked"%>></td>
								</tr>
								<tr> 
									<td class="td_l_c fontbold">11.</td>
									<td class="td_l_r title">高级搜索</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag11" value="1" <%if mid(qxflag, 11, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">12.</td>
									<td class="td_l_r title">客户转移</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag12" value="1" <%if mid(qxflag, 12, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">13.</td>
									<td class="td_l_r title">售后处理</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag13" value="1" <%if mid(qxflag, 13, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">14.</td>
									<td class="td_l_r title">合同审核</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag14" value="1" <%if mid(qxflag, 14, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">15.</td>
									<td class="td_l_r title">公海审核</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag15" value="1" <%if mid(qxflag, 15, 1) = "1" then Response.Write "checked"%>></td>
								</tr>
							</table>
						</fieldset>
						<fieldset style="margin-top:10px;padding:10px;">
							<input type="button" class="button246" onclick="javascript:selectall('levelA')" value="全选/反选" style="margin-bottom:10px;" />
							<legend>&nbsp;<B style="font-size:14px;">客户管理</B>&nbsp;</legend>
								<fieldset style="padding:10px;">
									<legend>&nbsp;客户档案&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr>
											<td class="td_l_c fontbold">16.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag16" value="1" <%if mid(qxflag, 16, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">17.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag17" value="1" <%if mid(qxflag, 17, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">18.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag18" value="1" <%if mid(qxflag, 18, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">19.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag19" value="1" <%if mid(qxflag, 19, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">20.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag20" value="1" <%if mid(qxflag, 20, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;联系人&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 		
											<td class="td_l_c fontbold">21.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag21" value="1" <%if mid(qxflag, 21, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">22.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag22" value="1" <%if mid(qxflag, 22, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">23.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag23" value="1" <%if mid(qxflag, 23, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">24.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag24" value="1" <%if mid(qxflag, 24, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">25.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag25" value="1" <%if mid(qxflag, 25, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;跟单管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
												
											<td class="td_l_c fontbold">26.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag26" value="1" <%if mid(qxflag, 26, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">27.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag27" value="1" <%if mid(qxflag, 27, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">28.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag28" value="1" <%if mid(qxflag, 28, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">29.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag29" value="1" <%if mid(qxflag, 29, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">30.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag30" value="1" <%if mid(qxflag, 30, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;订单管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
												
											<td class="td_l_c fontbold">31.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag31" value="1" <%if mid(qxflag, 31, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">32.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag32" value="1" <%if mid(qxflag, 32, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">33.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag33" value="1" <%if mid(qxflag, 33, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">34.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag34" value="1" <%if mid(qxflag, 34, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">35.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag35" value="1" <%if mid(qxflag, 35, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;合同管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%">
										<tr> 			
											<td class="td_l_c fontbold">36.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag36" value="1" <%if mid(qxflag, 36, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">37.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag37" value="1" <%if mid(qxflag, 37, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">38.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag38" value="1" <%if mid(qxflag, 38, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">39.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag39" value="1" <%if mid(qxflag, 39, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">40.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag40" value="1" <%if mid(qxflag, 40, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;售后管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
											<td class="td_l_c fontbold">41.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag41" value="1" <%if mid(qxflag, 41, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">42.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag42" value="1" <%if mid(qxflag, 42, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">43.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag43" value="1" <%if mid(qxflag, 43, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">44.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag44" value="1" <%if mid(qxflag, 44, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">45.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag45" value="1" <%if mid(qxflag, 45, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;费用管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
											<td class="td_l_c fontbold">46.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag46" value="1" <%if mid(qxflag, 46, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">47.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag47" value="1" <%if mid(qxflag, 47, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">48.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag48" value="1" <%if mid(qxflag, 48, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">49.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag49" value="1" <%if mid(qxflag, 49, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">50.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag50" value="1" <%if mid(qxflag, 50, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;附件管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
											<td class="td_l_c fontbold">51.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag51" value="1" <%if mid(qxflag, 51, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">52.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag52" value="1" <%if mid(qxflag, 52, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">53.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag53" value="1" <%if mid(qxflag, 53, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">54.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag54" value="1" <%if mid(qxflag, 54, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">55.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag55" value="1" <%if mid(qxflag, 55, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
						</fieldset>
						
						<fieldset style="margin-top:10px;padding:10px;">
						<input type="button" class="button246" onclick="javascript:selectall('levelB')" value="全选/反选" style="margin-bottom:10px;" />
							<legend>&nbsp;<B style="font-size:14px;">办公OA</B>&nbsp;</legend>
							
								<fieldset style="padding:10px;">
									<legend>&nbsp;内部公文&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 					
											<td class="td_l_c fontbold">56.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag56" value="1" <%if mid(qxflag, 56, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">57.</td>
											<td class="td_l_r title"><font color=red>新增</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag57" value="1" <%if mid(qxflag, 57, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">58.</td>
											<td class="td_l_r title"><font color=red>修改</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag58" value="1" <%if mid(qxflag, 58, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">59.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag59" value="1" <%if mid(qxflag, 59, 1) = "1" then Response.Write "checked"%>></td>	
												
											<td class="td_l_c fontbold">60.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag60" value="1" <%if mid(qxflag, 60, 1) = "1" then Response.Write "checked"%>></td>
											
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;站内短信&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 															
											<td class="td_l_c fontbold">61.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag61" value="1" <%if mid(qxflag, 61, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">62.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag62" value="1" <%if mid(qxflag, 62, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">63.</td>
											<td class="td_l_r title">回复</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag63" value="1" <%if mid(qxflag, 63, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">64.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag64" value="1" <%if mid(qxflag, 64, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">65.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag65" value="1" <%if mid(qxflag, 65, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;工作报告&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 									
											<td class="td_l_c fontbold">66.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag66" value="1" <%if mid(qxflag, 66, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">67.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag67" value="1" <%if mid(qxflag, 67, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">68.</td>
											<td class="td_l_r title"><font color=red>批注</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag68" value="1" <%if mid(qxflag, 68, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">69.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag69" value="1" <%if mid(qxflag, 69, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">70.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag70" value="1" <%if mid(qxflag, 70, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;其它&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
											<td class="td_l_c fontbold">71.</td>
											<td class="td_l_r title">文件柜</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag71" value="1" <%if mid(qxflag, 71, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">72.</td>
											<td class="td_l_r title">通讯录</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag72" value="1" <%if mid(qxflag, 72, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">73.</td>
											<td class="td_l_r title">个人日历</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag73" value="1" <%if mid(qxflag, 73, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">74.</td>
											<td class="td_l_r title">内部聊天</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag74" value="1" <%if mid(qxflag, 74, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">75.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag75" value="1" <%if mid(qxflag, 75, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;预留权限&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
											<td class="td_l_c fontbold">76.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag76" value="1" <%if mid(qxflag, 76, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">77.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag77" value="1" <%if mid(qxflag, 77, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">78.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag78" value="1" <%if mid(qxflag, 78, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">79.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag79" value="1" <%if mid(qxflag, 79, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">80.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag80" value="1" <%if mid(qxflag, 80, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
						</fieldset>
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
	lId = Trim(Request("lId"))
	lName = Trim(Request("lName"))
	qxflag = ""
	for i = 1 to 100
		if Request("qxflag" & i) = "1" then
			qxflag = qxflag & "1"
		else
			qxflag = qxflag & "0"
		end if
	next
		
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [system_Level] Where lId = " & lId & " Or lName = '" & lName & "' ",conn,3,1
	If rs.RecordCount > 0 Then
		Response.Write("<script>location.href='GetLevel.asp?action=Level&sType=Add&tipinfo=权限值或角色名称有重复';</script>")
		Response.End()
	End If
	rs.Close
	rs.Open "Select Top 1 * From [system_Level]",conn,3,2
	rs.AddNew
	rs("lId") = lId
	rs("lName") = lName
	rs("lQxfalg") = qxflag
	rs.Update
	rs.Close
	Set rs = Nothing
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	
elseif sType="Edit" then
	Id = Request("Id")
	qxflag = EasyCrm.getNewItem("system_Level","lId",""&Id&"","lQxfalg")
%>
		<form name="Save" id="Level" action="GetLevel.asp?action=Level&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>编辑角色 </B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> 权限值</td>
								<td class="td_l_l"><input name="lId" type="text" class="int" id="lId" size="10" maxlength="1" onkeyup='this.value=this.value.replace(/\D/gi,"")' value="<%=EasyCrm.getNewItem("system_Level","lId",""&Id&"","lId")%>" >  <span class="info_help help01">限：数字 2 - 8</span></td>
								<td class="td_l_r title"><font color="#FF0000">*</font> 角色名称</td>
								<td class="td_l_l"><input name="lName" type="text" class="int" id="lName" size="30" maxlength="16" value="<%=EasyCrm.getNewItem("system_Level","lId",""&Id&"","lName")%>" > </td>
							</tr>
							<tr> 
								<td class="td_l_r title" ><font color="#FF0000">*</font> 是否同步</td>
								<td class="td_l_l" colspan=3><input name="YnUpdate" type="radio" id="YnUpdate" value="1" > 是　<input name="YnUpdate" type="radio" id="YnUpdate" value="0" checked > 否  <span class="info_help help01">选中“是”，则更新当前角色内所有成员的权限</span></td>
							</tr>
						</table>
					</td>
				</tr>
				
				
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
		
						<fieldset style="padding:10px;">
							<legend>&nbsp;<B style="font-size:14px;">全局权限</B>&nbsp;</legend>
							<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
								<col width="4%"><col width="11%"><col width="5%"> 
								<col width="4%"><col width="11%"><col width="5%"> 
								<col width="4%"><col width="11%"><col width="5%"> 
								<col width="4%"><col width="11%"><col width="5%"> 
								<col width="4%"><col width="11%"><col width="5%"> 
								<tr> 
									<td class="td_l_c fontbold">01.</td>
									<td class="td_l_r title">系统登录</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag1" value="1" <%if mid(qxflag, 1, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">02.</td>
									<td class="td_l_r title">客户管理</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag2" value="1" <%if mid(qxflag, 2, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">03.</td>
									<td class="td_l_r title">办公OA</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag3" value="1" <%if mid(qxflag, 3, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">04.</td>
									<td class="td_l_r title">功能插件</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag4" value="1" <%if mid(qxflag, 4, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">05.</td>
									<td class="td_l_r title" onmouseover="tip.start(this)" tips="超管权限！"><font color=red><B>系统设置</B></font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag5" value="1" <%if mid(qxflag, 5, 1) = "1" then Response.Write "checked"%>></td>
								</tr>
								<tr> 							
									<td class="td_l_c fontbold">06.</td>
									<td class="td_l_r title" onmouseover="tip.start(this)" tips="客户档案下拉框字段新增项目"><font color=red>新增下拉框</font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag6" value="1" <%if mid(qxflag, 6, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">07.</td>
									<td class="td_l_r title" onmouseover="tip.start(this)" tips="导出客户数据，有一定风险！"><font color=red>导出Excel</font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag7" value="1" <%if mid(qxflag, 7, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">08.</td>
									<td class="td_l_r title" onmouseover="tip.start(this)" tips="导入客户数据，有一定风险！"><font color=red>导入Excel</font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag8" value="1" <%if mid(qxflag, 8, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">09.</td>
									<td class="td_l_r title" onmouseover="tip.start(this)" tips="批量更新数据，有一定风险！"><font color=red>批量操作</font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag9" value="1" <%if mid(qxflag, 9, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">10.</td>
									<td class="td_l_r title"><font color=red>客户共享</font></td>
									<td class="td_r_c"><input type="checkbox" name="qxflag10" value="1" <%if mid(qxflag, 10, 1) = "1" then Response.Write "checked"%>></td>
								</tr>
								<tr> 
									<td class="td_l_c fontbold">11.</td>
									<td class="td_l_r title">高级搜索</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag11" value="1" <%if mid(qxflag, 11, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">12.</td>
									<td class="td_l_r title">客户转移</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag12" value="1" <%if mid(qxflag, 12, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">13.</td>
									<td class="td_l_r title">售后处理</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag13" value="1" <%if mid(qxflag, 13, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">14.</td>
									<td class="td_l_r title">合同审核</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag14" value="1" <%if mid(qxflag, 14, 1) = "1" then Response.Write "checked"%>></td>
									
									<td class="td_l_c fontbold">15.</td>
									<td class="td_l_r title">公海审核</td>
									<td class="td_r_c"><input type="checkbox" name="qxflag15" value="1" <%if mid(qxflag, 15, 1) = "1" then Response.Write "checked"%>></td>
								</tr>
							</table>
						</fieldset>
						<fieldset style="margin-top:10px;padding:10px;">
							<input type="button" class="button246" onclick="javascript:selectall('levelA')" value="全选/反选" style="margin-bottom:10px;" />
							<legend>&nbsp;<B style="font-size:14px;">客户管理</B>&nbsp;</legend>
								<fieldset style="padding:10px;">
									<legend>&nbsp;客户档案&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr>
											<td class="td_l_c fontbold">16.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag16" value="1" <%if mid(qxflag, 16, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">17.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag17" value="1" <%if mid(qxflag, 17, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">18.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag18" value="1" <%if mid(qxflag, 18, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">19.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag19" value="1" <%if mid(qxflag, 19, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">20.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag20" value="1" <%if mid(qxflag, 20, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;联系人&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 		
											<td class="td_l_c fontbold">21.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag21" value="1" <%if mid(qxflag, 21, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">22.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag22" value="1" <%if mid(qxflag, 22, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">23.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag23" value="1" <%if mid(qxflag, 23, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">24.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag24" value="1" <%if mid(qxflag, 24, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">25.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag25" value="1" <%if mid(qxflag, 25, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;跟单管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
												
											<td class="td_l_c fontbold">26.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag26" value="1" <%if mid(qxflag, 26, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">27.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag27" value="1" <%if mid(qxflag, 27, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">28.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag28" value="1" <%if mid(qxflag, 28, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">29.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag29" value="1" <%if mid(qxflag, 29, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">30.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag30" value="1" <%if mid(qxflag, 30, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;订单管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
												
											<td class="td_l_c fontbold">31.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag31" value="1" <%if mid(qxflag, 31, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">32.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag32" value="1" <%if mid(qxflag, 32, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">33.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag33" value="1" <%if mid(qxflag, 33, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">34.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag34" value="1" <%if mid(qxflag, 34, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">35.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag35" value="1" <%if mid(qxflag, 35, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;合同管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%">
										<tr> 			
											<td class="td_l_c fontbold">36.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag36" value="1" <%if mid(qxflag, 36, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">37.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag37" value="1" <%if mid(qxflag, 37, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">38.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag38" value="1" <%if mid(qxflag, 38, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">39.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag39" value="1" <%if mid(qxflag, 39, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">40.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag40" value="1" <%if mid(qxflag, 40, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;售后管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
											<td class="td_l_c fontbold">41.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag41" value="1" <%if mid(qxflag, 41, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">42.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag42" value="1" <%if mid(qxflag, 42, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">43.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag43" value="1" <%if mid(qxflag, 43, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">44.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag44" value="1" <%if mid(qxflag, 44, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">45.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag45" value="1" <%if mid(qxflag, 45, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;费用管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
											<td class="td_l_c fontbold">46.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag46" value="1" <%if mid(qxflag, 46, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">47.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag47" value="1" <%if mid(qxflag, 47, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">48.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag48" value="1" <%if mid(qxflag, 48, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">49.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag49" value="1" <%if mid(qxflag, 49, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">50.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag50" value="1" <%if mid(qxflag, 50, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;附件管理&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
											<td class="td_l_c fontbold">51.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag51" value="1" <%if mid(qxflag, 51, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">52.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag52" value="1" <%if mid(qxflag, 52, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">53.</td>
											<td class="td_l_r title">修改</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag53" value="1" <%if mid(qxflag, 53, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">54.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag54" value="1" <%if mid(qxflag, 54, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">55.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelA" name="qxflag55" value="1" <%if mid(qxflag, 55, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
						</fieldset>
						
						<fieldset style="margin-top:10px;padding:10px;">
						<input type="button" class="button246" onclick="javascript:selectall('levelB')" value="全选/反选" style="margin-bottom:10px;" />
							<legend>&nbsp;<B style="font-size:14px;">办公OA</B>&nbsp;</legend>
							
								<fieldset style="padding:10px;">
									<legend>&nbsp;内部公文&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 					
											<td class="td_l_c fontbold">56.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag56" value="1" <%if mid(qxflag, 56, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">57.</td>
											<td class="td_l_r title"><font color=red>新增</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag57" value="1" <%if mid(qxflag, 57, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">58.</td>
											<td class="td_l_r title"><font color=red>修改</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag58" value="1" <%if mid(qxflag, 58, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">59.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag59" value="1" <%if mid(qxflag, 59, 1) = "1" then Response.Write "checked"%>></td>	
												
											<td class="td_l_c fontbold">60.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag60" value="1" <%if mid(qxflag, 60, 1) = "1" then Response.Write "checked"%>></td>
											
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;站内短信&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 															
											<td class="td_l_c fontbold">61.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag61" value="1" <%if mid(qxflag, 61, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">62.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag62" value="1" <%if mid(qxflag, 62, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">63.</td>
											<td class="td_l_r title">回复</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag63" value="1" <%if mid(qxflag, 63, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">64.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag64" value="1" <%if mid(qxflag, 64, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">65.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag65" value="1" <%if mid(qxflag, 65, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;工作报告&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 									
											<td class="td_l_c fontbold">66.</td>
											<td class="td_l_r title">查看</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag66" value="1" <%if mid(qxflag, 66, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">67.</td>
											<td class="td_l_r title">新增</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag67" value="1" <%if mid(qxflag, 67, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">68.</td>
											<td class="td_l_r title"><font color=red>批注</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag68" value="1" <%if mid(qxflag, 68, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">69.</td>
											<td class="td_l_r title"><font color=red>删除</font></td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag69" value="1" <%if mid(qxflag, 69, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">70.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag70" value="1" <%if mid(qxflag, 70, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;其它&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
											<td class="td_l_c fontbold">71.</td>
											<td class="td_l_r title">文件柜</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag71" value="1" <%if mid(qxflag, 71, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">72.</td>
											<td class="td_l_r title">通讯录</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag72" value="1" <%if mid(qxflag, 72, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">73.</td>
											<td class="td_l_r title">个人日历</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag73" value="1" <%if mid(qxflag, 73, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">74.</td>
											<td class="td_l_r title">内部聊天</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag74" value="1" <%if mid(qxflag, 74, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">75.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag75" value="1" <%if mid(qxflag, 75, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
								<fieldset style="margin-top:10px;">
									<legend>&nbsp;预留权限&nbsp;</legend>
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<col width="5%"><col width="10%"><col width="5%"> 
										<tr> 			
											<td class="td_l_c fontbold">76.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag76" value="1" <%if mid(qxflag, 76, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">77.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag77" value="1" <%if mid(qxflag, 77, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">78.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag78" value="1" <%if mid(qxflag, 78, 1) = "1" then Response.Write "checked"%>></td>
												
											<td class="td_l_c fontbold">79.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag79" value="1" <%if mid(qxflag, 79, 1) = "1" then Response.Write "checked"%>></td>
											
											<td class="td_l_c fontbold">80.</td>
											<td class="td_l_r title">预留</td>
											<td class="td_r_c"><input type="checkbox" id="levelB" name="qxflag80" value="1" <%if mid(qxflag, 80, 1) = "1" then Response.Write "checked"%>></td>
										</tr>
									</table>
								</fieldset>
						</fieldset>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
						<input name="lIdOld" type="hidden" id="lIdOld" value="<%=Id%>">
						<input name="lNameOld" type="hidden" id="lNameOld" value="<%=EasyCrm.getNewItem("system_Level","lId",""&Id&"","lName")%>">
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveEdit" then

	lId = Trim(Request("lId"))
	lIdOld = Trim(Request("lIdOld"))
	lName = Trim(Request("lName"))
	lNameOld = Trim(Request("lNameOld"))
	YnUpdate = Trim(Request("YnUpdate"))
		qxflag = ""
		for i = 1 to 100
			if Request("qxflag" & i) = "1" then
				qxflag = qxflag & "1"
			else
				qxflag = qxflag & "0"
			end if
		next
	
	
	if lId = lIdOld then '如果没更新权限值
		if lName <> lNameOld then
			'如果只修改角色名称，判断是否与其它角色名称重复
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open "Select * From [system_level] Where lName = '" & lName & "' ",conn,1,1
			If rs.RecordCount > 0 Then
				Response.Write("<script>location.href='GetLevel.asp?action=Level&sType=Edit&Id="&lId&"&tipinfo=角色名称有重复';</script>")
			Response.End()
			else
			conn.execute("update [system_level] set lName = '"&lName&"' where lName = '"&lNameOld&"' ")
			End If
			rs.Close
		end if
	
		conn.execute("update [system_level] set lQxfalg = '"&qxflag&"' where lId = "&lId&" ")
		if YnUpdate = "1" then '同步更新用户权限
		conn.execute("update [User] set uQxflag = '"&qxflag&"' where uLevel = "&lId&" ")
		end if
		
	else '如果更新了权限值，同步更新用户表
	
		'如果权限值与其它角色重复
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [system_level] Where lId = " & lId & " and lName <> '" & lNameOld & "' ",conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetLevel.asp?action=Level&sType=Edit&Id="&lIdOld&"&tipinfo=权限值有重复';</script>")
		Response.End()
		End If
		rs.Close
		
		if lName <> lNameOld then 
			'如果更改了角色名称，则判断角色名称是否与别的角色重复
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open "Select * From [system_level] Where lName = '" & lName & "' and lId="&lId&" ",conn,1,1
			If rs.RecordCount > 0 Then
				Response.Write("<script>location.href='GetLevel.asp?action=Level&sType=Edit&Id="&lId&"&tipinfo=角色名称有重复';</script>")
			Response.End()
			else
			conn.execute("update [system_level] set lId = '"&lId&"',lName='"&lName&"' where lId = "&lIdOld&" ")
			End If
			rs.Close
		else '如果只修改权限值，则不考虑角色名称
			conn.execute("update [system_level] set lId = '"&lId&"' where lId = "&lIdOld&" ")
		end if
	
		conn.execute("update [system_level] set lQxfalg = '"&qxflag&"' where lId = "&lIdOld&" ")
		if YnUpdate = "1" then '同步更新用户权限
		conn.execute("update [User] set uQxflag = '"&qxflag&"' where uLevel = "&lIdOld&" ")
		end if
		
	end if
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	
end if

End Sub
%>

<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>