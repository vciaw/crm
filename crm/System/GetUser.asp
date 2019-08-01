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


Sub User()
%>
	<script language="JavaScript">
	<!-- 必填项提示
	function CheckInput()
	{
		if(document.getElementById('account').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '登录账号不能为空！'});document.getElementById('account').focus();return false;}
		if(document.getElementById('name').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '真实姓名不能为空！'});document.getElementById('name').focus();return false;}
		if(document.getElementById('password').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '登录密码不能为空！'});document.getElementById('password').focus();return false;}
		if(document.getElementById('password').value != document.getElementById('confirmPWS').value){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '两次输入的密码不一样！'});document.getElementById('password').focus();return false;}
		if(document.getElementById('level').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '系统角色不能为空！'});document.getElementById('level').focus();return false;}
		if(document.getElementById('group').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '所属部门不能为空！'});document.getElementById('group').focus();return false;}
	}
	-->
	</script>
<%
if sType="Add" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10">   
			<div style="float:right;">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			</div>
            <div class="MenuboxS">
              <ul>
                <li class="hover"><span><a href="#">员工档案</a></span></li>
                <li><span><a href="#" style="text-decoration:line-through;color:#999">管理范围</a></span></li>
                <li><span><a href="#" style="text-decoration:line-through;color:#999">详细权限</a></span></li>
              </ul> 
            </div>
		</td>
	</tr>
</table>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="GetUser.asp?action=User&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="120" /><col width="260" /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="3" style="border-right:0;"><B>基本信息 </B></td>
								<td class="td_l_r"></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> 登录账号</td>
								<td class="td_l_l"><input name="account" type="text" class="int" id="account" size="11" maxlength="16"> <span class="info_help help01" onmouseover="tip.start(this)" tips="录入后不可修改">&nbsp;</span> </td>
								<td class="td_l_r title"><font color="#FF0000">*</font> 真实姓名</td>
								<td class="td_l_l"><input name="name" type="text" class="int" id="name" size="11" maxlength="16"></td>
							</tr>
							<tr>
								<td class="td_l_r title"><font color="#FF0000">*</font> 登录密码</td>
								<td class="td_l_l"><input name="password" type="password" class="int" id="password" size="11" maxlength="16"></td>
								<td class="td_l_r title"><font color="#FF0000">*</font> 重复密码</td>
								<td class="td_l_l"><input name="confirmPWS" type="password" class="int" id="confirmPWS" size="11" maxlength="16"></td>
							</tr>
							<tr>
								<td class="td_l_r title"><font color="#FF0000">*</font> 系统角色</td>
								<td class="td_l_l"><% = EasyCrm.getList(2,"system_level","lId","lName","level","") %></td>
								<td class="td_l_r title"><font color="#FF0000">*</font> 所属部门</td>
								<td class="td_l_l"><% = EasyCrm.getList(2,"system_group","gId","gName","group","") %></td>
							</tr>
							<tr>
								<td class="td_l_r title">生日</td>
								<td class="td_l_l"><input name="Birthday" type="text" id="Birthday" class="int Wdate" size="15" onFocus="WdatePicker()"></td>
								<td class="td_l_r title">入司时间</td>
								<td class="td_l_l"><input name="addtime" type="text" id="addtime" class="int Wdate" size="15" onFocus="WdatePicker()"></td>
							</tr>
							<tr>
								<td class="td_l_r title">手机</td>
								<td class="td_l_l"><input name="Mobile" type="text" class="int" id="Mobile" size="20" maxlength="16"></td>
								<td class="td_l_r title">E-mail</td>
								<td class="td_l_l"><input name="Email" type="text" class="int" id="Email" size="40"></td>
							</tr>
							<tr>
								<td class="td_l_r title">住址</td>
								<td class="td_l_l"><input name="Address" type="text" class="int" id="Address" size="40"></td>
								<td class="td_l_r title">身份证</td>
								<td class="td_l_l"><input name="card" type="text" class="int" id="card" size="40" maxlength="18"></td>
							</tr>
							<tr>
								<td class="td_l_r title">最大客户量</td>
								<td class="td_l_l" colspan=3><input name="ClientNum" type="text" id="ClientNum" class="int" size="10"  value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <span class="info_help help01" >&nbsp;０为不限制！</span></td>
							</tr>
						</table>
					</td>
				</tr>
				
				<tr>
					<td valign="top" class="td_n pdl10 pdr10"> 
						<div style="float:left;padding:10px 0 0;">
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
						</div>
					</td> 
				</tr>
				</form>
			</table>
<%
elseif sType="SaveAdd" then
	uAccount = Trim(Request("account"))
	uPassword = Lcase(Request("password"))
	uConfirmPWS = Trim(Request("confirmPWS"))
	uName = Trim(Request("name"))
	uLevel = CLng(Request("Level"))
	uGroup = CLng(Request("Group"))
	uBirthday = Trim(Request("Birthday"))
	uaddtime = Trim(Request("addtime"))
	uEmail = Trim(Request("Email"))
	uMobile = Trim(Request("Mobile"))
	ucard = Trim(Request("card"))
	uAddress = Trim(Request("Address"))
	uClientNum = Trim(Request("ClientNum"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [user] Where uName = '" & uName & "' or uAccount = '" & uAccount & "' ",conn,3,1
	If rs.RecordCount > 0 Then
		Response.Write("<script>location.href='GetUser.asp?action=User&sType=Add&tipinfo=登录名或姓名有重复';</script>")
	Response.End()
	End If
	rs.Close
	rs.Open "Select Top 1 * From [user]",conn,3,2
	rs.AddNew
	rs("uAccount") = EasyCrm.ReName(uAccount)
	rs("uPassword") = md5(uPassword,16)
	rs("uName") = uName
	rs("uLevel") = uLevel
	rs("uGroup") = uGroup
	rs("uMobile") = uMobile
	rs("uEmail") = uEmail
	rs("uAddress") = uAddress
	if uBirthday <> "" then 
	rs("uBirthday") = uBirthday
	end if
	rs("ucard") = ucard
	if uaddtime <> "" then
	rs("uaddtime") = uaddtime
	end if
	rs("uClientNum") = uClientNum
	rs("uQxflag") = EasyCrm.getNewItem("system_level","lId",""&uLevel&"","lQxfalg")
	rs.Update
	rs.Close
	Set rs = Nothing
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
elseif sType="Edit" then
	ssType = Trim(Request("ssType"))
	ID = Request("ID")
%>
<style>body{padding:45px 0 55px 0;}</style>
<div class="fixed_bg_T">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt5">   
			<div style="float:right;">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			</div>
            <div class="MenuboxS">
              <ul>
                <li <%if ssType="Main" or ssType="" then%>class="hover"<%end if%>><span><a href="?action=User&sType=Edit&ssType=Main&ID=<%=ID%>">员工档案</a></span></li>
                <li <%if ssType="Manage" then%>class="hover"<%end if%>><span><a href="?action=User&sType=Edit&ssType=Manage&ID=<%=ID%>">管理范围</a></span></li>
                <li <%if ssType="Level" then%>class="hover"<%end if%>><span><a href="?action=User&sType=Edit&ssType=Level&ID=<%=ID%>">详细权限</a></span></li>
              </ul> 
            </div>
		</td>
	</tr>
</table>
</DIV>
		<%if ssType="Main" or ssType="" then%>
		<form name="Save" action="GetUser.asp?action=User&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="120" /><col width="260" /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>基本信息 </B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> 登录账号</td>
								<td class="td_l_l"> <%=EasyCrm.getNewItem("User","uId",""&Id&"","uAccount")%> <span class="info_help help01" onmouseover="tip.start(this)" tips="不可修改">&nbsp;</span> </td>
								<td class="td_l_r title"><font color="#FF0000">*</font> 真实姓名</td>
								<td class="td_l_l"><input name="name" type="text" class="int" id="name" size="11" maxlength="16" value="<%=EasyCrm.getNewItem("User","uId",""&Id&"","uName")%>"> <span class="info_help help01" onmouseover="tip.start(this)" tips="更新所有记录，慎重！">&nbsp;</span></td>
							</tr>
							<tr>
								<td class="td_l_r title"><font color="#FF0000">*</font> 登录密码</td>
								<td class="td_l_l"><input name="password" type="password" class="int" id="password" size="11" maxlength="16"> <span class="info_help help01" onmouseover="tip.start(this)" tips="不修改请留空">&nbsp;</span></td>
								<td class="td_l_r title"><font color="#FF0000">*</font> 重复密码</td>
								<td class="td_l_l"><input name="confirmPWS" type="password" class="int" id="confirmPWS" size="11" maxlength="16"></td>
							</tr>
							<tr>
								<td class="td_l_r title"><font color="#FF0000">*</font> 系统角色</td>
								<td class="td_l_l"><% = EasyCrm.getList(2,"system_level","lId","lName","level",EasyCrm.getNewItem("system_level","lId",EasyCrm.getNewItem("User","uId",""&Id&"","uLevel"),"lName")) %></td>
								<td class="td_l_r title"><font color="#FF0000">*</font> 所属部门</td>
								<td class="td_l_l"><% = EasyCrm.getList(2,"system_group","gId","gName","group",EasyCrm.getNewItem("system_group","gId",EasyCrm.getNewItem("User","uId",""&Id&"","uGroup"),"gName")) %></td>
							</tr>
							<tr>
								<td class="td_l_r title">生日</td>
								<td class="td_l_l"><input name="Birthday" type="text" id="Birthday" class="int Wdate" size="15" onFocus="WdatePicker()" value="<%=EasyCrm.getNewItem("User","uId",""&Id&"","uBirthday")%>"></td>
								<td class="td_l_r title">入司时间</td>
								<td class="td_l_l"><input name="addtime" type="text" id="addtime" class="int Wdate" size="15" onFocus="WdatePicker()" value="<%=EasyCrm.getNewItem("User","uId",""&Id&"","uAddtime")%>"></td>
							</tr>
							<tr>
								<td class="td_l_r title">手机</td>
								<td class="td_l_l"><input name="Mobile" type="text" class="int" id="Mobile" size="20" maxlength="16" value="<%=EasyCrm.getNewItem("User","uId",""&Id&"","uMobile")%>"></td>
								<td class="td_l_r title">E-mail</td>
								<td class="td_l_l"><input name="Email" type="text" class="int" id="Email" size="40" value="<%=EasyCrm.getNewItem("User","uId",""&Id&"","uEmail")%>"></td>
							</tr>
							<tr>
								<td class="td_l_r title">住址</td>
								<td class="td_l_l"><input name="Address" type="text" class="int" id="Address" size="40" value="<%=EasyCrm.getNewItem("User","uId",""&Id&"","uAddress")%>"></td>
								<td class="td_l_r title">身份证</td>
								<td class="td_l_l"><input name="card" type="text" class="int" id="card" size="40" maxlength="18" value="<%=EasyCrm.getNewItem("User","uId",""&Id&"","uCard")%>"></td>
							</tr>
							<tr>
								<td class="td_l_r title">最大客户量</td>
								<td class="td_l_l" colspan=3><input name="ClientNum" type="text" id="ClientNum" class="int" size="10" value="<%=EasyCrm.getNewItem("User","uId",""&Id&"","uClientNum")%>" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <span class="info_help help01" >&nbsp;０为不限制！</span></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input name="uid" type="hidden" id="uid" value="<%=Id%>">
							<input name="OldName" type="hidden" id="OldName" value="<%=EasyCrm.getNewItem("User","uId",""&Id&"","uName")%>">
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
		<%elseif ssType="Manage" then%>
		<form name="Save" action="?action=User&sType=Edit&ssType=SaveManage" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>跨部门管理范围 </B></td>
							</tr>
							<%
								Set rsg = Server.CreateObject("ADODB.Recordset")
								rsg.Open "Select * From [system_group]",conn,1,1
								Do While Not rsg.BOF And Not rsg.EOF
							%>
							<tr> 
								<td class="td_l_r title"><%=rsg("gName")%></td>
								<td  class="td_l_l">
								<%
									Set rsm = Server.CreateObject("ADODB.Recordset")
									rsm.Open "Select * From [user] where uGroup="&rsg("gId")&" ",conn,1,1
									Do While Not rsm.BOF And Not rsm.EOF
								%>
								<input type="checkbox" name="umanagerange" value="<%=rsm("uName")%>" <%if rsm("uAccount")=EasyCrm.getNewItem("User","uId",""&Id&"","uAccount") then%>checked<%elseif rsm("uName")=EasyCrm.getNewItem("User","uId",""&Id&"","uName") then%>checked<%elseif rsm("uGroup") = EasyCrm.getNewItem("User","uId",""&Id&"","uGroup") and rsm("uLevel") < EasyCrm.getNewItem("User","uId",""&Id&"","uLevel") then %>checked<%else%><%if inStr(EasyCrm.getNewItem("User","uId",""&Id&"","uManagerange"),rsm("uName"))>0 then%>checked<%end if%><%end if%> > <%=rsm("uName")%>　
								<%
									rsm.MoveNext
									Loop
									rsm.Close
									Set rsm = Nothing
								%>
								</td>
							</tr> 
							<%
								rsg.MoveNext
								Loop
								rsg.Close
								Set rsg = Nothing
							%>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
						<input name="uid" type="hidden" id="uid" value="<%=Id%>">
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
		<%
		elseif ssType="SaveManage" then
		
		uId = Request("uId")
		umanagerange = Trim(Request("umanagerange"))
		conn.execute("update [User] set umanagerange = '"&umanagerange&"' where uId = "&uId&" ")
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
		elseif ssType="Level" then
		qxflag = EasyCrm.getNewItem("User","uId",""&Id&"","uQxflag")
		%>
			<script language=javascript> 
			//全选/反选
			function selectall(id){ //用id区分  
			var tform=document.forms['Level'];  
			for(var i=0;i<tform.length;i++){  
			var e=tform.elements[i];  
			if(e.type=="checkbox" && e.id==id) e.checked=!e.checked;  } }
			</script> 
		<form name="Save" id="Level" action="?action=User&sType=Edit&ssType=SaveLevel" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
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
						<input name="uid" type="hidden" id="uid" value="<%=Id%>">
						<input type="submit" name="Submit" class="button45" value="保存">　
						<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
		<%
		elseif ssType="SaveLevel" then
		
		uId = Request("uId")
		qxflag = ""
		for i = 1 to 100
			if Request("qxflag" & i) = "1" then
				qxflag = qxflag & "1"
			else
				qxflag = qxflag & "0"
			end if
		next
		conn.execute("update [User] set uQxflag = '"&qxflag&"' where uId = "&uId&" ")
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
		end if%>
		
<%
elseif sType="SaveEdit" then
	uId = Request("uId")
	uPassword = Lcase(Request("password"))
	uConfirmPWS = Trim(Request("confirmPWS"))
	uName = Trim(Request("name"))
	OldName = Trim(Request("OldName"))
	uLevel = CLng(Request("Level"))
	uGroup = CLng(Request("Group"))
	uBirthday = Trim(Request("Birthday"))
	uaddtime = Trim(Request("addtime"))
	uEmail = Trim(Request("Email"))
	uMobile = Trim(Request("Mobile"))
	ucard = Trim(Request("card"))
	uAddress = Trim(Request("Address"))
	uClientNum = Trim(Request("ClientNum"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [user] Where uName = '" & uName & "' And uId <>  " & uId,conn,1,1
	If rs.RecordCount > 0 Then
		Response.Write("<script>location.href='GetUser.asp?action=User&sType=Add&tipinfo=真实姓名重复';</script>")
	Response.End()
	End If
	rs.Close
	rs.Open "Select Top 1 * From [user] Where uId = " & uId,conn,3,2
	if uPassword <> "" then
	rs("uPassword") = md5(uPassword,16)
	end if
	if ""&OldName&"" <> ""&uName&"" then
		conn.execute("update [Client] set cUser = '"&uName&"' where cUser = '"&OldName&"' ")
		conn.execute("update [Linkmans] set lUser = '"&uName&"' where lUser = '"&OldName&"' ")
		conn.execute("update [Records] set rUser = '"&uName&"' where rUser = '"&OldName&"' ")
		conn.execute("update [Order] set oUser = '"&uName&"' where oUser = '"&OldName&"' ")
		conn.execute("update [Hetong] set hUser = '"&uName&"' where hUser = '"&OldName&"' ")
		conn.execute("update [Service] set sUser = '"&uName&"' where sUser = '"&OldName&"' ")
		conn.execute("update [Expense] set eUser = '"&uName&"' where eUser = '"&OldName&"' ")
		conn.execute("update [OA_Report] set oUser = '"&uName&"' where oUser = '"&OldName&"' ")
		conn.execute("update [OA_soft] set s_user = '"&uName&"' where s_user = '"&OldName&"' ")
		conn.execute("update [Calendar] set calendaruser = '"&uName&"' where calendaruser = '"&OldName&"' ")
	end if
	rs("uName") = uName
	rs("uLevel") = uLevel
	rs("uGroup") = uGroup
	rs("uMobile") = uMobile
	rs("uEmail") = uEmail
	rs("uAddress") = uAddress
	if uBirthday <> "" then 
	rs("uBirthday") = uBirthday
	end if
	rs("ucard") = ucard
	if uaddtime <> "" then
	rs("uaddtime") = uaddtime
	end if
	rs("uClientNum") = uClientNum
	rs("uQxflag") = EasyCrm.getNewItem("system_level","lId",""&uLevel&"","lQxfalg")
	rs.Update
	rs.Close
	Set rs = Nothing
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	
end if
%>

<%
End Sub
%>

<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>