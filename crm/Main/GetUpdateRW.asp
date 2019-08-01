<!--#include file="../Data/Conn.asp"--><!--#include file="../UpLoad/UpLoad.asp"--><!--#include file="../data/EasyCrm.asp"-->
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
<style>body {padding-bottom:55px;}</style>
<%
action = Trim(Request("action"))
sType = Trim(Request("sType"))
oType = Trim(Request("oType"))
YNUpdate = Trim(Request("YNUpdate")) '是否同步更新
cID = Trim(Request("cID"))
ID = Trim(Request("ID"))
tipinfo = Trim(Request("tipinfo"))

From_url = Cstr(Request.ServerVariables("HTTP_Referer"))
Serv_url = Cstr(Request.ServerVariables("Server_Name"))
If mid(From_url,8,len(Serv_url)) <> Serv_url Then
	Response.Write("<script>window.opener=null;window.close();</script>")
	Response.end
End If

	if tipinfo<>"" then
		Response.Write("<script>art.dialog({title: 'Error',time: 1,icon: 'warning',content: '"&tipinfo&"'});</script>")
	end if

Select Case action
Case "Linkmans"
    Call Linkmans()
Case "Records"
    Call Records()
Case "Order"
    Call Order()
Case "OrderProducts"
    Call OrderProducts()
Case "Hetong"
    Call Hetong()
Case "Service"
    Call Service()
Case "Expense"
    Call Expense()
Case "File"
    Call File()
Case "History"
    Call History()
Case "Choose"
    Call Choose()
End Select

Sub Linkmans() '联系人
%>
	<script language="JavaScript">
	<!-- 联系人必填项提示
	function CheckInput()
	{
		if (<%=Must_Linkmans_lName%>=="1"){if(document.all.lName.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lName & alert04%>'});document.all.lName.focus();return false;}}
		if (<%=Must_Linkmans_lSex%>=="1"){if(document.all.lSex.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lSex & alert04%>'});document.all.lSex.focus();return false;}}
		if (<%=Must_Linkmans_lZhiwei%>=="1"){if(document.all.lZhiwei.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lZhiwei & alert04%>'});document.all.lZhiwei.focus();return false;}}
		if (<%=Must_Linkmans_lBirthday%>=="1"){if(document.all.lBirthday.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lBirthday & alert04%>'});document.all.lBirthday.focus();return false;}}
		if (<%=Must_Linkmans_lMobile%>=="1"){if(document.all.lMobile.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lMobile & alert04%>'});document.all.lMobile.focus();return false;}}
		if (<%=Must_Linkmans_lTel%>=="1"){if(document.all.lTel.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lTel & alert04%>'});document.all.lTel.focus();return false;}}
		if (<%=Must_Linkmans_lEmail%>=="1"){if(document.all.lEmail.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lEmail & alert04%>'});document.all.lEmail.focus();return false;}}
		if (<%=Must_Linkmans_lQQ%>=="1"){if(document.all.lQQ.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lQQ & alert04%>'});document.all.lQQ.focus();return false;}}
		if (<%=Must_Linkmans_lMSN%>=="1"){if(document.all.lMSN.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lMSN & alert04%>'});document.all.lMSN.focus();return false;}}
		if (<%=Must_Linkmans_lALWW%>=="1"){if(document.all.lALWW.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lALWW & alert04%>'});document.all.lALWW.focus();return false;}}
		if (<%=Must_Linkmans_lContent%>=="1"){if(document.all.lContent.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Linkmans_lContent & alert04%>'});document.all.lContent.focus();return false;}}
	}
	-->
	</script>
	<%
	if sType="Add" then '添加
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Linkmans&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>新增联系人</B></td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Linkmans_lName = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lName%></td>
								<td class="td_r_l"> <input type="text" class="int" name="lName" id="lName" size="20" maxlength="20" > </td>
								<td class="td_l_r title"> <%if Must_Linkmans_lSex = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lSex%></td>
								<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Sex","lSex","") %></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Linkmans_lZhiwei = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lZhiwei%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_Zhiwei","lZhiwei","") %>&nbsp;<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Zhiwei_InfoAdd()' style="cursor:pointer"><script>function Select_Zhiwei_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Zhiwei', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%></td>
								<td class="td_l_r title"><%if Must_Linkmans_lBirthday = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lBirthday%></td>
								<td class="td_r_l"> <input name="lBirthday" type="text" id="lBirthday" class="Wdate" size="20" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" /></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Linkmans_lMobile = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lMobile%></td>
								<td class="td_r_l"> <input name="lMobile" type="text" class="int" id="lMobile" size="20"></td>
								<td class="td_l_r title"><%if Must_Linkmans_lTel = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lTel%></td>
								<td class="td_r_l" colspan="3"> <input name="lTel" type="text" class="int" id="lTel" size="20"></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Linkmans_lEmail = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lEmail%></td>
								<td class="td_r_l"> <input name="lEmail" type="text" class="int" id="lEmail" size="20"></td>
								<td class="td_l_r title"><%if Must_Linkmans_lQQ = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lQQ%></td>
								<td class="td_r_l" colspan="3"> <input name="lQQ" type="text" class="int" id="lQQ" size="20"></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Linkmans_lMSN = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lMSN%></td>
								<td class="td_r_l"> <input name="lMSN" type="text" class="int" id="lMSN" size="20"></td>
								<td class="td_l_r title"><%if Must_Linkmans_lALWW = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lALWW%></td>
								<td class="td_r_l" colspan="3"> <input name="lALWW" type="text" class="int" id="lALWW" size="20"></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Linkmans_lContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="lContent" rows="4" id="lContent" class="int" style="height:50px;width:98%;"></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="cID" type="hidden" value="<%=cID%>">
			<input name="lUser" type="hidden" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAdd" then '保存添加
		cID = Request.Form("cID")
		lName = Request.Form("lName")
		lSex = Request.Form("lSex")
		lZhiwei = Request.Form("lZhiwei")
		lBirthday = Request.Form("lBirthday")
		lMobile = Request.Form("lMobile")
		lTel = Request.Form("lTel")
		lEmail = Request.Form("lEmail")
		lQQ = Request.Form("lQQ")
		lMSN = Request.Form("lMSN")
		lALWW = Request.Form("lALWW")
		lContent = Request.Form("lContent")
		lUser = Request.Form("lUser")
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [Linkmans] Where lName = '"&lName&"' and cID="&cID&" ",conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='../Main/GetUpdateRW.asp?action=Linkmans&sType=Add&cID="&cID&"&tipinfo=该联系人已存在，请重新输入！';</script>")
		Response.End()
		End If
		rs.Close
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Linkmans]",conn,3,2
		rs.AddNew
		rs("cID") = cID
		rs("lName") = lName
		rs("lSex") = lSex
		rs("lZhiwei") = lZhiwei
		if lBirthday <>"" then
		rs("lBirthday") = lBirthday
		end if
		rs("lMobile") = lMobile
		rs("lTel") = lTel
		rs("lEmail") = lEmail
		rs("lQQ") = lQQ
		rs("lMSN") = lMSN
		rs("lALWW") = lALWW
		rs("lContent") = lContent
		rs("lUser") = lUser
		rs("lTime") = now()
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Linkmans&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Edit" then '修改
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Linkmans&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="12"><B>修改联系人</B></td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Linkmans_lName = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lName%></td>
								<td class="td_r_l"> <input type="text" class="int" name="lName" id="lName" size="20" maxlength="20" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lName")%>" > </td>
								<td class="td_l_r title"> <%if Must_Linkmans_lSex = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lSex%></td>
								<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Sex","lSex","'"&EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lSex")&"'") %></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Linkmans_lZhiwei = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lZhiwei%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_Zhiwei","lZhiwei",EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lZhiwei")) %>&nbsp;<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Zhiwei_InfoAdd()' style="cursor:pointer"><script>function Select_Zhiwei_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Zhiwei', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%></td>
								<td class="td_l_r title"><%if Must_Linkmans_lBirthday = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lBirthday%></td>
								<td class="td_r_l"> <input name="lBirthday" type="text" id="lBirthday" class="Wdate" size="20" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lBirthday"),2)%>"  /></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Linkmans_lMobile = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lMobile%></td>
								<td class="td_r_l"> <input name="lMobile" type="text" class="int" id="lMobile" size="20" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lMobile")%>" ></td>
								<td class="td_l_r title"><%if Must_Linkmans_lTel = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lTel%></td>
								<td class="td_r_l" colspan="3"> <input name="lTel" type="text" class="int" id="lTel" size="20" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lTel")%>" ></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Linkmans_lEmail = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lEmail%></td>
								<td class="td_r_l"> <input name="lEmail" type="text" class="int" id="lEmail" size="20" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lEmail")%>" ></td>
								<td class="td_l_r title"><%if Must_Linkmans_lQQ = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lQQ%></td>
								<td class="td_r_l" colspan="3"> <input name="lQQ" type="text" class="int" id="lQQ" size="20" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lQQ")%>" ></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Linkmans_lMSN = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lMSN%></td>
								<td class="td_r_l"> <input name="lMSN" type="text" class="int" id="lMSN" size="20" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lMSN")%>" ></td>
								<td class="td_l_r title"><%if Must_Linkmans_lALWW = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lALWW%></td>
								<td class="td_r_l" colspan="3"> <input name="lALWW" type="text" class="int" id="lALWW" size="20" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lALWW")%>" ></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Linkmans_lContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="lContent" rows="4" id="lContent" class="int" style="height:50px;width:98%;"><%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lContent")%></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="lID" type="hidden" value="<%=ID%>">
			<input name="cID" type="hidden" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","cID")%>">
			<input name="YNUpdate" type="hidden" value="<%=YNUpdate%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveEdit" then '保存修改
		cID = Request.Form("cID")
		lID = Request.Form("lID")
		lName = Request.Form("lName")
		lSex = Request.Form("lSex")
		lZhiwei = Request.Form("lZhiwei")
		lBirthday = Request.Form("lBirthday")
		lMobile = Request.Form("lMobile")
		lTel = Request.Form("lTel")
		lEmail = Request.Form("lEmail")
		lQQ = Request.Form("lQQ")
		lMSN = Request.Form("lMSN")
		lALWW = Request.Form("lALWW")
		lContent = Request.Form("lContent")
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [Linkmans] Where lName = '"&lName&"' and cID="&cID&" and lID<>"&lID&" ",conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='../Main/GetUpdateRW.asp?action=Linkmans&sType=Edit&ID="&lID&"&tipinfo=该联系人已存在，请重新输入！';</script>")
		Response.End()
		End If
		rs.Close
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Linkmans] where lID="&lID,conn,3,2
		rs("lName") = lName
		rs("lSex") = lSex
		rs("lZhiwei") = lZhiwei
		if lBirthday <>"" then
		rs("lBirthday") = lBirthday
		end if
		rs("lMobile") = lMobile
		rs("lTel") = lTel
		rs("lEmail") = lEmail
		rs("lQQ") = lQQ
		rs("lMSN") = lMSN
		rs("lALWW") = lALWW
		rs("lContent") = lContent
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		if YNUpdate="1" then
		conn.execute ("UPDATE [client] SET cLinkman='"&lName&"',cZhiwei='"&lZhiwei&"',cMobile='"&lMobile&"' Where cId ="&cId&" ")
		end if
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Linkmans&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Delete" then '删除
		If Id = "" Then Exit Sub
		cID = EasyCrm.getNewItem("Linkmans","lID",""&ID&"","cID")
		conn.execute("DELETE FROM [Linkmans] where lId = "&Id&" ")
		
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cID&"','"&L_Linkmans&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	end if
end Sub

Sub Records() '跟单记录
%>
	<script language="JavaScript">
	<!-- 跟单记录必填项提示
	function CheckInput()
	{
		if (<%=Must_Records_rType%>=="1"){if(document.all.rType.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Records_rType & alert04%>'});document.all.rType.focus();return false;}}
		if (<%=Must_Records_rState%>=="1"){if(document.all.rState.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Records_rState & alert04%>'});document.all.rState.focus();return false;}}
		if (<%=Must_Records_rLinkman%>=="1"){if(document.all.rLinkman.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Records_rLinkman & alert04%>'});document.all.rLinkman.focus();return false;}}
		if (<%=Must_Records_rNextTime%>=="1"){if(document.all.rNextTime.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Records_rNextTime & alert04%>'});document.all.rNextTime.focus();return false;}}
		if (<%=Must_Records_rRemind%>=="1"){if(document.all.rRemind.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Records_rRemind & alert04%>'});document.all.rRemind.focus();return false;}}
		if (<%=Must_Records_rContent%>=="1"){if(document.all.rContent.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Records_rContent & alert04%>'});document.all.rContent.focus();return false;}}
	}
	-->
	</script>
	<%
	if sType="Add" then '添加
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Records&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="80" /><col /><col width="80" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>新增跟单记录</B></td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Records_rType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rType%></td>
								<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Records","rType","") %>&nbsp;<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Records_InfoAdd()' style="cursor:pointer"><script>function Select_Records_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Records', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%></td>
								<td class="td_l_r title"> <%if Must_Records_rState = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rState%></td>
								<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Type","rState","") %>&nbsp;<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Type_InfoAdd()' style="cursor:pointer"><script>function Select_Type_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Type', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%> <span class="info_help help01" >&nbsp;同步客户类型</span></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Records_rLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rLinkman%></td>
								<td class="td_r_l"> 
									<% = EasyCrm.getNewSelect("linkmans","lName","rLinkman"," and cid="&cid&" ","") %>&nbsp;
									<input name="Back" type="button" id="Back" class="button222" value="新增" onclick='Linkmans_InfoAdd()' style="cursor:pointer"><script>function Linkmans_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Linkmans&sType=Add&cID=<%=cID%>', {title: '新窗口', width: 700, height: 340,fixed: true}); };</script>
								</td>
								<td class="td_l_r title"><%if Must_Records_rNextTime = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rNextTime%></td>
								<td class="td_r_l"> <input name="rNextTime" type="text" id="rNextTime" class="Wdate" size="22" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:00:00'})" /> 　提前：
									<select name="rRemind">
										<option value="1">1小时</option>
										<option value="2">2小时</option>
										<option value="3">3小时</option>
										<option value="24">1　天</option>
										<option value="48">2　天</option>
										<option value="72">3　天</option>
										<option value="168">1　周</option>
									</select> 提醒
								</td>
							</tr>
							
								<%
								Set rss = Server.CreateObject("ADODB.Recordset")
								rss.Open "Select * From [CustomField] where cTable='Records' order by Id asc ",conn,3,1
								If rss.RecordCount > 0 Then
								i=1
								Do While Not rss.BOF And Not rss.EOF
								if i mod 2 = 1 then Response.Write "<tr>"
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											next
											%>
										<%end if%>
									</td>
								<%
								i = i + 1
								if i mod 2 = 1 then Response.Write "</tr>"
								rss.MoveNext
								Loop
								end if
								rss.Close
								Set rss = Nothing
								%>
							<tr> 
								<td class="td_l_r title"><%if Must_Records_rContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rContent%></td>
								<td class="td_r_l" COLSPAN="3" style="padding:5px 10px;"> <textarea name="rContent" rows="4" id="rContent" class="int" style="height:70px;width:95%;"></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="cID" type="hidden" value="<%=cID%>">
			<input name="rUser" type="hidden" value="<%=Session("CRM_name")%>">
			<input name="cType" type="hidden" id="cType" value="<%=EasyCrm.getNewItem("Client","cId",""&cID&"","cType")%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAdd" then '保存添加
		cID = Request.Form("cID")
		rType = Request.Form("rType")
		rState = Request.Form("rState")
		rlinkman = Request.Form("rlinkman")
		rNextTime = Request.Form("rNextTime")
		rRemind = Request.Form("rRemind")
		rContent = Request.Form("rContent")
		rUser = Request.Form("rUser")
		cType = Request.Form("cType")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Records]",conn,3,2
		rs.AddNew
		rs("cID") = cID
		rs("rType") = rType
		rs("rState") = rState
		rs("rlinkman") = rlinkman
		if rNextTime <>"" then
		rs("rNextTime") = rNextTime
		end if
		if rRemind <>"" then
		rs("rRemind") = rRemind
		end if
		rs("rContent") = rContent
		rs("rUser") = rUser
		rs("rTime") = now()
    	rs.Update
    	rs.Close
    	Set rs = Nothing

	Dim rsid
	Set rsid = Server.CreateObject("ADODB.Recordset")
	if Accsql = 0 then
	rsid.Open "Select top 1 rID From [Records] order by rID desc",conn,1,1
	elseif Accsql = 1 then
	rsid.Open "Select @@IDENTITY as rID From [Records]",conn,1,1
	end if
	rID=rsid("rID")
	rsid.close
	
	'插入自定义内容
	
	cContent = ""
	Set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open "Select * From [CustomField] where cTable='Records' order by Id asc ",conn,1,1
	If rsc.RecordCount > 0 Then
	Do While Not rsc.BOF And Not rsc.EOF
	
	cContent = cContent & rsc("cName") &":"& Trim(Request(rsc("cName"))) &"|"
	
	rsc.MoveNext
	Loop
	end if
	rsc.Close
	Set rsc = Nothing
	
	conn.execute ("insert into CustomFieldContent(cID,rID,cContent) values('"&cid&"','"&rID&"','"&cContent&"')")	
	
	
		
		'同步更新客户类型和插入定时站内信
		
		if ""&rState&"" <> "" and ""&cType&"" <> ""&CRTypeEnd&"" then
		conn.execute ("UPDATE client SET cType='"&rState&"' Where cId ="&cId&" ")
		end if
		
		if rNextTime <> "" then
		RemindTime = Dateadd("h",-rRemind,rNextTime)
		conn.execute ("UPDATE client SET cRNextTime='"&rNextTime&"' Where cId ="&cId&" ")
		conn.execute ("insert into OA_mms_Receive(oReceiver,oSender,oTitle,oContent,oIsread,oAttime,oTime) values('"&Session("CRM_name")&"','系统通知','["&EasyCrm.getNewItem("Client","cid",""&cID&"","cCompany")&"] 于 ["&RemindTime&"] 需再次跟单!','',0,'"&RemindTime&"','"&now()&"')")	
		end if
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Records&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Edit" then '修改
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Records&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="80" /><col /><col width="80" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>修改跟单记录</B></td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Records_rType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rType%></td>
								<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Records","rType",""&EasyCrm.getNewItem("Records","rID",""&ID&"","rType")&"") %>&nbsp;<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Records_InfoAdd()' style="cursor:pointer"><script>function Select_Records_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Records', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%></td>
								<td class="td_l_r title"> <%if Must_Records_rState = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rState%></td>
								<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Type","rState",""&EasyCrm.getNewItem("Records","rID",""&ID&"","rState")&"") %>&nbsp;<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Type_InfoAdd()' style="cursor:pointer"><script>function Select_Type_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Type', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%> <span class="info_help help01" >&nbsp;同步客户类型</span></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Records_rLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rLinkman%></td>
								<td class="td_r_l"> 
									<% = EasyCrm.getNewSelect("Linkmans","lName","rLinkman"," and cid="&EasyCrm.getNewItem("Records","rID",""&ID&"","cID")&" ",""&EasyCrm.getNewItem("Records","rID",""&ID&"","rLinkman")&"") %>&nbsp;
									<input name="Back" type="button" id="Back" class="button222" value="新增" onclick='Linkmans_InfoAdd()' style="cursor:pointer"><script>function Linkmans_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Linkmans&sType=Add&cID=<%=EasyCrm.getNewItem("Records","rID",""&ID&"","cID")%>', {title: '新窗口', width: 700, height: 340,fixed: true}); };</script>
								</td>
								<td class="td_l_r title"><%if Must_Records_rNextTime = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rNextTime%></td>
								<td class="td_r_l"> <input name="rNextTime" type="text" id="rNextTime" class="Wdate" size="22" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:00:00'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Records","rID",""&ID&"","rNextTime"),1)%>" />　 提前：<select name="rRemind" class="int">
										<option value="1">1小时</option>
										<option value="2">2小时</option>
										<option value="3">3小时</option>
										<option value="24">1　天</option>
										<option value="48">2　天</option>
										<option value="72">3　天</option>
										<option value="168">1　周</option>
									</select> 提醒 　<span class="b red">重新提醒：<input name="RepeatRemind" type="checkbox" value="1"></span>
								</td>
							</tr>
							<%
								cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID"," "&EasyCrm.getNewItem("Records","rID",""&ID&"","cID")&" And rID = "&ID&" ","cContent")
								cContentArr = split(cContentStr,"|")								
								Set rss = Server.CreateObject("ADODB.Recordset")
								rss.Open "Select * From [CustomField] where cTable='Records' order by Id asc ",conn,1,1
								If rss.RecordCount > 0 Then
								i=1:k=0
								Do While Not rss.BOF And Not rss.EOF
								if i mod 2 = 1 then Response.Write "<tr>"
								if Ubound(cContentArr) > k then
								cContent = split(cContentArr(k),":")
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
									<%if inStr(cContentArr(k),cContent(0))>0 then%>
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="<%=cContent(1)%>">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="<%=cContent(1)%>" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											if selectstr(selectarr) = cContent(1) then
											response.Write "<option value="""&selectstr(selectarr)&""" selected>"&selectstr(selectarr)&"</option>"
											else
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											end if
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											if checkboxstr(checkboxarr) = cContent(1) then
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&""" checked> "&checkboxstr(checkboxarr)&"　"
											else
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											end if
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											if radiostr(radioarr) = cContent(1) then
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&""" checked> "&radiostr(radioarr)&"　"
											else
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											end if
											next
											%>
										<%end if%>
									<%end if%>
									</td>
								<%
								else
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											next
											%>
										<%end if%>
									</td>
								<%
								end if
								i = i + 1:k=k+1
								if i mod 2 = 1 then Response.Write "</tr>"
								rss.MoveNext
								Loop
								end if
								rss.Close
								Set rss = Nothing
							%>
							<tr> 
								<td class="td_l_r title"><%if Must_Records_rContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="rContent" rows="4" id="rContent" class="int" style="height:70px;width:95%;"><%=EasyCrm.getNewItem("Records","rID",""&ID&"","rContent")%></textarea></td>
							</tr>
							<script language="JavaScript">
							<!--
							for(var i=0;i<document.all.rRemind.options.length;i++){
								if(document.all.rRemind.options[i].value == "<%=EasyCrm.getNewItem("Records","rID",""&ID&"","rRemind")%>"){
								document.all.rRemind.options[i].selected = true;}}
							-->
							</script>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="rID" type="hidden" value="<%=ID%>">
			<input name="cType" type="hidden" id="cType" value="<%=EasyCrm.getNewItem("Client","cId",EasyCrm.getNewItem("Records","rID",""&ID&"","cID"),"cType")%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveEdit" then '保存修改
		rID = Request.Form("rID")
		rType = Request.Form("rType")
		rState = Request.Form("rState")
		rLinkman = Request.Form("rLinkman")
		rNextTime = Request.Form("rNextTime")
		rRemind = Request.Form("rRemind")
		rContent = Request.Form("rContent")
		cType = Request.Form("cType")
		cID = EasyCrm.getNewItem("Records","rID",""&rID&"","cID")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Records] where rID="&rID,conn,3,2
		rs("rType") = rType
		rs("rState") = rState
		rs("rLinkman") = rLinkman
		if rNextTime <>"" then
		rs("rNextTime") = rNextTime
		end if
		rs("rRemind") = rRemind
		rs("rContent") = rContent
    	rs.Update
    	rs.Close
    	Set rs = Nothing
	
	'更新自定义内容
	
	cContent = ""
	Set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open "Select * From [CustomField] where cTable='Records' order by Id asc ",conn,3,1
	If rsc.RecordCount > 0 Then
	Do While Not rsc.BOF And Not rsc.EOF
	'获取所有自定义字段
	cContent = cContent & rsc("cName") &":"& Trim(Request(rsc("cName"))) &"|" 
	rsc.MoveNext
	Loop
	end if
	rsc.Close
	Set rsc = Nothing
	if EasyCrm.getNewItem("CustomFieldContent","cID",""&cID&" and rID="&rID&" ","cContent")="0" then
	conn.execute ("insert into CustomFieldContent(cID,rID,cContent) values('"&cid&"','"&rID&"','"&cContent&"')")	
	else
	conn.execute ("UPDATE [CustomFieldContent] SET cContent='"&cContent&"' Where cId ="&cId&" and rID="&rID&" ")
	end if
		
		'同步更新客户类型和插入定时站内信
		
		if ""&rState&"" <> "" and ""&cType&"" <> ""&CRTypeEnd&"" then
		conn.execute ("UPDATE client SET cType='"&rState&"' Where cId ="&EasyCrm.getNewItem("Records","rID",""&rID&"","cID")&" ")
		end if
		
		if rNextTime <> "" then
		if RepeatRemind = 1 then
		RemindTime = Dateadd("h",-rRemind,rNextTime)
		conn.execute ("insert into OA_mms_Receive(oReceiver,oSender,oTitle,oContent,oIsread,oAttime,oTime) values('"&Session("CRM_name")&"','系统通知','["&EasyCrm.getNewItem("Client","cid",EasyCrm.getNewItem("Records","rID",""&rID&"","cID"),"cCompany")&"] 于 ["&RemindTime&"] 需再次跟单!','',0,'"&RemindTime&"','"&now()&"')")	
		end if
		conn.execute ("UPDATE client SET cRNextTime='"&rNextTime&"' Where cId ="&cId&" ")
		end if
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Records&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="DelReason" then '删除客户填写操作原因
	%>	
	<script language="JavaScript">
	<!-- 跟单记录必填项提示
	function CheckInput()
	{
		if(document.all.Reason.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=alert04%>'});document.all.Reason.focus();return false;}
	}
	-->
	</script>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="?action=Records&sType=Delete&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<textarea name="Reason" rows="4" id="Reason" class="int" style="height:80px;width:98%;"></textarea>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="Id" type="hidden" id="cId" value="<%=ID%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="Delete" then '删除
		If Id = "" Then Exit Sub
		Reason = Trim(Request("Reason"))
		cID = EasyCrm.getNewItem("Records","rID",""&ID&"","cID")
		conn.execute("DELETE FROM [Records] where rId = "&Id&" ")
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lReason,lUser,lTime) values('"&cID&"','"&L_Records&"','"&L_insert_action_03&"','"&Reason&"','"&Session("CRM_name")&"','"&now()&"')")
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	end if
end Sub

Sub Order() '订单记录
%>
	<script language="JavaScript">
	<!-- 跟单记录必填项提示
	function CheckInput()
	{
		if (<%=Must_Order_oLinkman%>=="1"){if(document.all.oLinkman.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Order_oLinkman & alert04%>'});document.all.oLinkman.focus();return false;}}
		if (<%=Must_Order_oSDate%>=="1"){if(document.all.oSDate.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Order_oSDate & alert04%>'});document.all.oSDate.focus();return false;}}
		if (<%=Must_Order_oEDate%>=="1"){if(document.all.oEDate.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Order_oEDate & alert04%>'});document.all.oEDate.focus();return false;}}
		if (<%=Must_Order_oDeposit%>=="1"){if(document.all.oDeposit.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Order_oDeposit & alert04%>'});document.all.oDeposit.focus();return false;}}
		if (<%=Must_Order_oContent%>=="1"){if(document.all.oContent.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Order_oContent & alert04%>'});document.all.oContent.focus();return false;}}
	}
	-->
	</script>
	<%
	if sType="Add" then '添加
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Order&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>新增订单记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Order_oCode%></td>
								<td class="td_r_l" COLSPAN="3"> <input name="oCode" type="text" id="oCode" size="23" class="int" <%if YnDDNum=1 then%> value="DD<%=EasyCrm.FormatDate(now(),14)&right(FormatNumber(timer(),3),3)%>" readonly style="border:0;" <%end if%> /> <span class="info_help help01" >&nbsp;手动模式请按固定规则编号</span> </td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Order_oLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_oLinkman%></td>
								<td class="td_r_l"> 
									<% = EasyCrm.getNewSelect("linkmans","lName","oLinkman"," and cid="&cid&" ","") %>&nbsp;
									<input name="Back" type="button" id="Back" class="button222" value="新增" onclick='Linkmans_InfoAdd()' style="cursor:pointer"><script>function Linkmans_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Linkmans&sType=Add&cID=<%=cID%>', {title: '新窗口', width: 700, height: 340,fixed: true}); };</script>
								</td>
								<td class="td_l_r title"><%if Must_Order_oDeposit = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_oDeposit%></td>
								<td class="td_r_l"> <input name="oDeposit" type="text" id="oDeposit" class="int" size="10" value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%></td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Order_oSDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_oSDate%></td>
								<td class="td_r_l"> <input name="oSDate" type="text" id="oSDate" class="Wdate" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(date(),2)%>" /></td>
								<td class="td_l_r title"> <%if Must_Order_oEDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_oEDate%></td>
								<td class="td_r_l"> <input name="oEDate" type="text" id="oEDate" class="Wdate" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" /></td>
							</tr>
							
								<%
								Set rss = Server.CreateObject("ADODB.Recordset")
								rss.Open "Select * From [CustomField] where cTable='Order' order by Id asc ",conn,3,1
								If rss.RecordCount > 0 Then
								i=1
								Do While Not rss.BOF And Not rss.EOF
								if i mod 2 = 1 then Response.Write "<tr>"
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											next
											%>
										<%end if%>
									</td>
								<%
								i = i + 1
								if i mod 2 = 1 then Response.Write "</tr>"
								rss.MoveNext
								Loop
								end if
								rss.Close
								Set rss = Nothing
								%>
							<tr> 
								<td class="td_l_r title"><%if Must_Order_oContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_oContent%></td>
								<td class="td_r_l" COLSPAN="3" style="padding:5px 10px;"> <textarea name="oContent" rows="4" id="oContent" class="int" style="height:50px;width:98%;"></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="cID" type="hidden" value="<%=cID%>">
			<input name="oUser" type="hidden" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAdd" then '保存添加
		cID = Request.Form("cID")
		oCode = Request.Form("oCode")
		oLinkman = Request.Form("oLinkman")
		oSDate = Request.Form("oSDate")
		oEDate = Request.Form("oEDate")
		oDeposit = Request.Form("oDeposit")
		oContent = Request.Form("oContent")
		oUser = Request.Form("oUser")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Order]",conn,3,2
		rs.AddNew
		rs("cID") = cID
		rs("oCode") = oCode
		rs("oLinkman") = oLinkman
		if oSDate <> "" then
		rs("oSDate") = oSDate
		end if
		if oEDate <> "" then
		rs("oEDate") = oEDate
		end if
		if oDeposit <> "" then
		rs("oDeposit") = oDeposit
		end if
		rs("oContent") = oContent
		rs("oMoney") = 0
		rs("oState") = 0
		rs("oUser") = oUser
		rs("oTime") = now()
    	rs.Update
    	rs.Close
    	Set rs = Nothing

	Dim rsid
	Set rsid = Server.CreateObject("ADODB.Recordset")
	if Accsql = 0 then
	rsid.Open "Select top 1 oID From [Order] order by oID desc",conn,1,1
	elseif Accsql = 1 then
	rsid.Open "Select @@IDENTITY as oID From [Order]",conn,1,1
	end if
	oID=rsid("oID")
	rsid.close
	
	'插入自定义内容
	
	cContent = ""
	Set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open "Select * From [CustomField] where cTable='Order' order by Id asc ",conn,1,1
	If rsc.RecordCount > 0 Then
	Do While Not rsc.BOF And Not rsc.EOF
	
	cContent = cContent & rsc("cName") &":"& Trim(Request(rsc("cName"))) &"|"
	
	rsc.MoveNext
	Loop
	end if
	rsc.Close
	Set rsc = Nothing
	
	conn.execute ("insert into CustomFieldContent(cID,oID,cContent) values('"&cid&"','"&oID&"','"&cContent&"')")	
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		if oEDate <> "" then
		conn.execute("update [Client] set cOEDate='"&oEDate&"' where cId = "&cId&" ")
		end if
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Order&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Edit" then '修改
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Order&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  /><col width="120" /><col  /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>修改订单记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Order_oCode%></td>
								<td class="td_r_l" COLSPAN="3"> <%=EasyCrm.getNewItem("Order","oID",""&ID&"","oCode")%> <span class="info_help help01" >&nbsp;不可修改</span></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Order_oLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_oLinkman%></td>
								<td class="td_r_l"> 
									<% = EasyCrm.getNewSelect("Linkmans","lName","oLinkman"," and cid="&EasyCrm.getNewItem("Order","oID",""&ID&"","cID")&" ",EasyCrm.getNewItem("Order","oID",""&ID&"","oLinkman")) %>&nbsp;
									<input name="Back" type="button" id="Back" class="button222" value="新增" onclick='Linkmans_InfoAdd()' style="cursor:pointer"><script>function Linkmans_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Linkmans&sType=Add&cID=<%=EasyCrm.getNewItem("Order","oID",""&ID&"","cID")%>', {title: '新窗口', width: 700, height: 340,fixed: true}); };</script>
								</td>
								<td class="td_l_r title"><%if Must_Order_oDeposit = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_oDeposit%></td>
								<td class="td_r_l"> <input name="oDeposit" type="text" id="oDeposit" class="int" size="10" value="<%=EasyCrm.getNewItem("Order","oID",""&ID&"","oDeposit")%>" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%></td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Order_oSDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_oSDate%></td>
								<td class="td_r_l"> <input name="oSDate" type="text" id="oSDate" class="Wdate" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Order","oID",""&ID&"","oSDate"),2)%>" /></td>
								<td class="td_l_r title"> <%if Must_Order_oEDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_oEDate%></td>
								<td class="td_r_l"> <input name="oEDate" type="text" id="oEDate" class="Wdate" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Order","oID",""&ID&"","oEDate"),2)%>" /></td>
							</tr>
							<%
								cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID"," "&EasyCrm.getNewItem("Order","oID",""&ID&"","cID")&" And oID = "&ID&" ","cContent")
								cContentArr = split(cContentStr,"|")								
								Set rss = Server.CreateObject("ADODB.Recordset")
								rss.Open "Select * From [CustomField] where cTable='Order' order by Id asc ",conn,1,1
								If rss.RecordCount > 0 Then
								i=1:k=0
								Do While Not rss.BOF And Not rss.EOF
								if i mod 2 = 1 then Response.Write "<tr>"
								if Ubound(cContentArr) > k then
								cContent = split(cContentArr(k),":")
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
									<%if inStr(cContentArr(k),cContent(0))>0 then%>
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="<%=cContent(1)%>">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="<%=cContent(1)%>" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											if selectstr(selectarr) = cContent(1) then
											response.Write "<option value="""&selectstr(selectarr)&""" selected>"&selectstr(selectarr)&"</option>"
											else
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											end if
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											if checkboxstr(checkboxarr) = cContent(1) then
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&""" checked> "&checkboxstr(checkboxarr)&"　"
											else
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											end if
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											if radiostr(radioarr) = cContent(1) then
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&""" checked> "&radiostr(radioarr)&"　"
											else
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											end if
											next
											%>
										<%end if%>
									<%end if%>
									</td>
								<%
								else
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											next
											%>
										<%end if%>
									</td>
								<%
								end if
								i = i + 1:k=k+1
								if i mod 2 = 1 then Response.Write "</tr>"
								rss.MoveNext
								Loop
								end if
								rss.Close
								Set rss = Nothing
							%>
							<tr> 
								<td class="td_l_r title"><%if Must_Order_oContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_oContent%></td>
								<td class="td_r_l" COLSPAN="3" style="padding:5px 10px;"> <textarea name="oContent" rows="4" id="oContent" class="int" style="height:50px;width:98%;"><%=EasyCrm.getNewItem("Order","oID",""&ID&"","oContent")%></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="oID" type="hidden" value="<%=ID%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveEdit" then '保存修改
		oID = Request.Form("oID")
		oLinkman = Request.Form("oLinkman")
		oSDate = Request.Form("oSDate")
		oEDate = Request.Form("oEDate")
		oDeposit = Request.Form("oDeposit")
		oContent = Request.Form("oContent")
		cID = EasyCrm.getNewItem("Order","oId",""&oID&"","cID")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Order] where oid="&oid,conn,3,2
		rs("oLinkman") = oLinkman
		if oSDate <> "" then
		rs("oSDate") = oSDate
		end if
		if oEDate <> "" then
		rs("oEDate") = oEDate
		end if
		if oDeposit <> "" then
		rs("oDeposit") = oDeposit
		end if
		if oMoney <> "" then
		end if
		rs("oContent") = oContent
    	rs.Update
    	rs.Close
    	Set rs = Nothing
	
	'更新自定义内容
	
	cContent = ""
	Set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open "Select * From [CustomField] where cTable='Order' order by Id asc ",conn,3,1
	If rsc.RecordCount > 0 Then
	Do While Not rsc.BOF And Not rsc.EOF
	'获取所有自定义字段
	cContent = cContent & rsc("cName") &":"& Trim(Request(rsc("cName"))) &"|" 
	rsc.MoveNext
	Loop
	end if
	rsc.Close
	Set rsc = Nothing
	if EasyCrm.getNewItem("CustomFieldContent","cID",""&cID&" and oID="&oID&" ","cContent")="0" then
	conn.execute ("insert into CustomFieldContent(cID,oID,cContent) values('"&cid&"','"&oID&"','"&cContent&"')")	
	else
	conn.execute ("UPDATE [CustomFieldContent] SET cContent='"&cContent&"' Where cId ="&cId&" and oID="&oID&" ")
	end if
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		if oEDate <> "" then
		conn.execute("update [Client] set cOEDate='"&oEDate&"' where cId = "&cId&" ")
		end if
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Order&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

		
	elseif sType="Audit" then '填写审核原因
	%>	
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="?action=Order&sType=SaveAudit&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						状态：
						<input name="oState" type="radio" value="2" <%if EasyCrm.getNewItem("Order","oID",""&ID&"","oState") = "2" then%> checked<%end if%>> 已完成　 
						<input name="oState" type="radio" value="1" <%if EasyCrm.getNewItem("Order","oID",""&ID&"","oState") = "1" then%> checked<%end if%>> 处理中 　 
						<input name="oState" type="radio" value="0" <%if EasyCrm.getNewItem("Order","oID",""&ID&"","oState") = "0" then%> checked<%end if%>> 转未处理 
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<textarea name="oAuditReasons" rows="4" id="oAuditReasons" class="int" style="height:80px;width:98%;"><%=EasyCrm.getNewItem("Order","oID",""&ID&"","oAuditReasons")%></textarea>
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
	elseif sType="SaveAudit" then '保存
		If Id = "" Then Exit Sub
		oState = Request.Form("oState")
		oAuditReasons = Request.Form("oAuditReasons")
		conn.execute("update [Order] set oState = '"&oState&"',oAuditReasons = '"&oAuditReasons&"' where oId = "&ID&" ")
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

	elseif sType="DelReason" then '删除填写操作原因
	%>	
	<script language="JavaScript">
	<!-- 
	function CheckInput()
	{
		if(document.all.Reason.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=alert04%>'});document.all.Reason.focus();return false;}
	}
	-->
	</script>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="?action=Order&sType=Delete&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<textarea name="Reason" rows="4" id="Reason" class="int" style="height:80px;width:98%;"></textarea>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="Id" type="hidden" id="Id" value="<%=ID%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="Delete" then '删除
		If Id = "" Then Exit Sub
		Reason = Trim(Request("Reason"))
		cID = EasyCrm.getNewItem("Order","oId",""&ID&"","cID")
		conn.execute("DELETE FROM [Order] where oId = "&Id&" ")
		'删除订单同步删除订单详情的产品信息
		conn.execute("DELETE FROM [Order_Products] where oId = "&Id&" ")
		'写入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lReason,lUser,lTime) values('"&cID&"','"&L_Order&"','"&L_insert_action_03&"','"&Reason&"','"&Session("CRM_name")&"','"&now()&"')")
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	end if
end Sub

Sub OrderProducts() '订单产品明细
%>
<style>body {padding-bottom:55px;}</style>
	<script language="JavaScript">
	<!-- 产品必填项提示
	function CheckInput()
	{
		if (<%=Must_Order_Products_oProTitle%>=="1"){if(document.all.ProTitle.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Order_Products_oProTitle & alert04%>'});document.all.ProTitle.focus();return false;}}
		if (<%=Must_Order_Products_oProNum%>=="1"){if(document.all.oProNum.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Order_Products_oProNum & alert04%>'});document.all.oProNum.focus();return false;}}
		if (<%=Must_Order_Products_oDiscount%>=="1"){if(document.all.oDiscount.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Order_Products_oDiscount & alert04%>'});document.all.oDiscount.focus();return false;}}
		if (<%=Must_Order_Products_oContent%>=="1"){if(document.all.oContent.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Order_Products_oContent & alert04%>'});document.all.oContent.focus();return false;}}
	}
	-->
	</script>
	<%
	if sType="" or sType = "List" then '列表
	%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">订单编号: <%=EasyCrm.getNewItem("Order","oID",""&ID&"","oCode")%></td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_set" value=" " title="设置" onclick='Setting_OrderProducts()' style="cursor:pointer" />
        </td>
	</tr>
</table>
<script>function Setting_OrderProducts() {$.dialog.open('../system/GetUpdate.asp?action=Setting&sType=OrderProducts', {title: '自定义设置', width: 900, height: 480,fixed: true}); };</script>
		<style>body{padding-top:35px;}</style>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<%if Order_Products_ProId = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_ProId%></td>
								<%end if%>
								<%if Order_Products_oProTitle = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oProTitle%></td>
								<%end if%>
								<%if pItemA = 1 then %>
								<%if Order_Products_oProItemA = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oProItemA%></td>
								<%end if%>
								<%end if%>
								<%if pItemB = 1 then %>
								<%if Order_Products_oProItemB = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oProItemB%></td>
								<%end if%>
								<%end if%>
								<%if pItemC = 1 then %>
								<%if Order_Products_oProItemC = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oProItemC%></td>
								<%end if%>
								<%end if%>
								<%if pItemD = 1 then %>
								<%if Order_Products_oProItemD = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oProItemD%></td>
								<%end if%>
								<%end if%>
								<%if pItemE = 1 then %>
								<%if Order_Products_oProItemE = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oProItemE%></td>
								<%end if%>
								<%end if%>
								<%if Order_Products_oProPrice = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oProPrice%></td>
								<%end if%>
								<%if Order_Products_oProNum = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oProNum%></td>
								<%end if%>
								<%if Order_Products_oProUnit = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oProUnit%></td>
								<%end if%>
								<%if Order_Products_oDiscount = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oDiscount%></td>
								<%end if%>
								<%if Order_Products_oMoney = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oMoney%></td>
								<%end if%>
								<%if Order_Products_oUser = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oUser%></td>
								<%end if%>
								<%if Order_Products_oTime = 1 then %>
								<td class="td_l_c"><%=L_Order_Products_oTime%></td>
								<%end if%>
								<td width="90" class="td_l_c">管理</td>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Order_Products] where oId = "&Id&" Order By oId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
							<tr class="tr">
								<%if Order_Products_ProId = 1 then %>
								<td class="td_l_c"><%=rs("ProId")%></td>
								<%end if%>
								<%if Order_Products_oProTitle = 1 then %>
								<td class="td_l_c"><%=EasyCrm.getNewItem("Products","id",rs("ProId"),"pTitle")%></td>
								<%end if%>
								<%if pItemA = 1 then %>
								<%if Order_Products_oProItemA = 1 then %>
								<td class="td_l_c"><%=EasyCrm.getNewItem("Products","id",rs("ProId"),"pItemA")%></td>
								<%end if%>
								<%end if%>
								<%if pItemB = 1 then %>
								<%if Order_Products_oProItemB = 1 then %>
								<td class="td_l_c"><%=EasyCrm.getNewItem("Products","id",rs("ProId"),"pItemB")%></td>
								<%end if%>
								<%end if%>
								<%if pItemC = 1 then %>
								<%if Order_Products_oProItemC = 1 then %>
								<td class="td_l_c"><%=EasyCrm.getNewItem("Products","id",rs("ProId"),"pItemC")%></td>
								<%end if%>
								<%end if%>
								<%if pItemD = 1 then %>
								<%if Order_Products_oProItemD = 1 then %>
								<td class="td_l_c"><%=EasyCrm.getNewItem("Products","id",rs("ProId"),"pItemD")%></td>
								<%end if%>
								<%end if%>
								<%if pItemE = 1 then %>
								<%if Order_Products_oProItemE = 1 then %>
								<td class="td_l_c"><%=EasyCrm.getNewItem("Products","id",rs("ProId"),"pItemE")%></td>
								<%end if%>
								<%end if%>
								<%if Order_Products_oProPrice = 1 then %>
								<td class="td_l_c"><%=EasyCrm.getNewItem("Products","id",rs("ProId"),"pUprice")%></td>
								<%end if%>
								<%if Order_Products_oProNum = 1 then %>
								<td class="td_l_c"><%=rs("oProNum")%></td>
								<%end if%>
								<%if Order_Products_oProUnit = 1 then %>
								<td class="td_l_c"><%=rs("oUser")%></td>
								<%end if%>
								<%if Order_Products_oDiscount = 1 then %>
								<td class="td_l_c"><%=rs("oDiscount")%></td>
								<%end if%>
								<%if Order_Products_oMoney = 1 then %>
								<td class="td_l_c"><%if rs("oMoney")<1 and rs("oMoney")>0 then%>0<%end if%><%=rs("oMoney")%></td>
								<%end if%>
								<%if Order_Products_oUser = 1 then %>
								<td class="td_l_c"><%=rs("oUser")%></td>
								<%end if%>
								<%if Order_Products_oTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("oTime"),2)%></td>
								<%end if%>
								<td class="td_l_c"><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Order_Products_InfoEdit<%=rs("osId")%>()' style="cursor:pointer" /> <input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Order_Products_InfoDel<%=rs("osId")%>()' style="cursor:pointer" /></td>
							</tr>
							<script>function Order_Products_InfoEdit<%=rs("osId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=OrderProducts&sType=Edit&Id=<%=rs("osId")%>', {title: '编辑', width: 700,height: 400, fixed: true}); };</script>
							<script>function Order_Products_InfoDel<%=rs("osId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () { art.dialog.open('../Main/GetUpdateRW.asp?action=OrderProducts&sType=Delete&Id=<%=rs("osId")%>');return false;},cancel: true }); };</script>
						<%
						rs.MoveNext
						Loop
							else
							%>
							<tr><td class="td_l_l" colspan=13><%=L_Notfound%></td></tr>
							<%
							end if
						rs.Close
						Set rs = Nothing
						%>
							
						</table>
					</td> 
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="Back" type="button" id="Back" class="button45" value="新增" onclick='Order_Products_InfoAdd()' style="cursor:pointer">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();art.dialog.open.origin.location.reload();">
			
		</td>
	</tr>
</table>
</div>
<script>function Order_Products_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=OrderProducts&sType=Add&ID=<%=ID%>', {title: '新窗口', width: 700, height: 400,fixed: true}); };</script>
	<%
	elseif sType="Add" then '添加
	%>
		<form name="Save" action="?action=OrderProducts&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  /><col width="120" /><col  /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>新增订单产品</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Order_Products_oProTitle = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_Products_oProTitle%></td>
								<td class="td_r_l" COLSPAN="3"> 
									<input name="ProId" type="hidden" id="ProId" size="23" value="" readonly />
									<input name="ProTitle" type="text" id="ProTitle" size="23" value="" readonly onclick='Choose_Products()' style="cursor:pointer" /> 
									<input name="Back" type="button" id="Back" class="button221" value="…" title="请选择" onclick='Choose_Products()' style="cursor:pointer"><script>function Choose_Products() {$.dialog.open('../Main/GetUpdateRW.asp?action=Choose&sType=Products&oType=Order', {title: '新窗口', width: 900, height: 480,fixed: true}); };</script>
								</td>
							</tr>
							<tr> 
								<%if pItemA = 1 then %>
								<td class="td_l_r title"><%=L_Order_Products_oProItemA%></td>
								<td class="td_r_l" <%if pItemB = 0 then %> colspan=3 <%end if%>> <input name="oProItemA" type="text" id="oProItemA" value="" readonly /> </td>
								<%end if%>
								<%if pItemB = 1 then %>
								<td class="td_l_r title"><%=L_Order_Products_oProItemB%></td>
								<td class="td_r_l" <%if pItemA = 0 then %> colspan=3 <%end if%>> <input name="oProItemB" type="text" id="oProItemB" value="" readonly /> </td>
								<%end if%>
							</tr>
							<tr> 
								<%if pItemC = 1 then %>
								<td class="td_l_r title"><%=L_Order_Products_oProItemC%></td>
								<td class="td_r_l" <%if pItemD = 0 then %> colspan=3 <%end if%>> <input name="oProItemC" type="text" id="oProItemC" value="" readonly /> </td>
								<%end if%>
								<%if pItemD = 1 then %>
								<td class="td_l_r title"><%=L_Order_Products_oProItemD%></td>
								<td class="td_r_l" <%if pItemC = 0 then %> colspan=3 <%end if%>> <input name="oProItemD" type="text" id="oProItemD" value="" readonly /> </td>
								<%end if%>
							</tr>
							<tr> 
								<%if pItemE = 1 then %>
								<td class="td_l_r title"><%=L_Order_Products_oProItemE%></td>
								<td class="td_r_l"> <input name="oProItemE" type="text" id="oProItemE" value="" readonly /> </td>
								<%end if%>
								<td class="td_l_r title"><%=L_Order_Products_oProPrice%></td>
								<td class="td_r_l" <%if pItemE = 0 then %> colspan=3 <%end if%>> <input name="oProPrice" type="text" id="oProPrice" size="10" onchange="oMoney.value=oProPrice.value*parseInt(oProNum.value)-oDiscount.value" /> <%=L_Yuan%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Order_Products_oProNum = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_Products_oProNum%></td>
								<td class="td_r_l"> <input name="oProNum" type="text" id="oProNum" size="10" onchange="oMoney.value=oProPrice.value*parseInt(oProNum.value)-oDiscount.value" value="1" onfocus="if (value =='1'){value ='1'}"onblur="if (value ==''){value='1'}" /> </td>
								<td class="td_l_r title"><%if Must_Order_Products_oDiscount = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_Products_oDiscount%></td>
								<td class="td_r_l"> <input name="oDiscount" type="text" id="oDiscount" size="10" onchange="oMoney.value=oProPrice.value*parseInt(oProNum.value)-oDiscount.value" value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%> </td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Order_Products_oMoney%></td>
								<td class="td_r_l" COLSPAN="3"> <input name="oMoney" type="text" id="oMoney" size="10" style="font-weight:bold;color:Red;" class="" value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" readonly /> <%=L_Yuan%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Order_Products_oContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_Products_oContent%></td>
								<td class="td_r_l" COLSPAN="3" style="padding:5px 10px;"> <textarea name="oContent" rows="4" id="oContent" class="int" style="height:50px;width:98%;"></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input name="oId" type="hidden" value="<%=ID%>">
							<input name="oUser" type="hidden" value="<%=Session("CRM_name")%>">
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
	<%
	elseif sType="SaveAdd" then '保存添加
		oID = Request.Form("oID")
		ProId = Request.Form("ProId")
		oProNum = Request.Form("oProNum")
		oDiscount = Request.Form("oDiscount")
		oMoney = Request.Form("oMoney")
		oContent = Request.Form("oContent")
		oUser = Request.Form("oUser")
		cID = EasyCrm.getNewItem("Order","oID",""&oID&"","cID")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Order_Products]",conn,3,2
		rs.AddNew
		rs("oID") = oID
		rs("cID") = cID
		rs("ProId") = ProId
		rs("oProNum") = oProNum
		rs("oDiscount") = oDiscount
		rs("oMoney") = oMoney
		rs("oContent") = oContent
		rs("oUser") = oUser
		rs("oTime") = now()
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		'添加订单产品同时更新订单状态为处理中，和更新订单总金额
		conn.execute("update [Order] set oState = '1',oMoney = oMoney+'"&oMoney&"' where oId = "&oID&" ")
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','订单产品','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Edit" then '修改
	%>
		<form name="Save" action="?action=OrderProducts&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  /><col width="120" /><col  /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>修改订单产品</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Order_Products_oProTitle = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_Products_oProTitle%></td>
								<td class="td_r_l" COLSPAN="3"> 
									<input name="ProTitle" type="text" id="ProTitle" size="23" value="<%=EasyCrm.getNewItem("Products","ID",EasyCrm.getNewItem("Order_Products","osID",""&ID&"","ProId"),"pTitle")%>" readonly /> 
								</td>
							</tr>
							<tr> 
								<%if pItemA = 1 then %>
								<td class="td_l_r title"><%=L_Order_Products_oProItemA%></td>
								<td class="td_r_l" <%if pItemB = 0 then %> colspan=3 <%end if%>> <input name="oProItemA" type="text" id="oProItemA" value="<%=EasyCrm.getNewItem("Products","ID",EasyCrm.getNewItem("Order_Products","osID",""&ID&"","ProId"),"pItemA")%>" readonly /> </td>
								<%end if%>
								<%if pItemB = 1 then %>
								<td class="td_l_r title"><%=L_Order_Products_oProItemB%></td>
								<td class="td_r_l" <%if pItemA = 0 then %> colspan=3 <%end if%>> <input name="oProItemB" type="text" id="oProItemB" value="<%=EasyCrm.getNewItem("Products","ID",EasyCrm.getNewItem("Order_Products","osID",""&ID&"","ProId"),"pItemB")%>" readonly /> </td>
								<%end if%>
							</tr>
							<tr> 
								<%if pItemC = 1 then %>
								<td class="td_l_r title"><%=L_Order_Products_oProItemC%></td>
								<td class="td_r_l" <%if pItemD = 0 then %> colspan=3 <%end if%>> <input name="oProItemC" type="text" id="oProItemC" value="<%=EasyCrm.getNewItem("Products","ID",EasyCrm.getNewItem("Order_Products","osID",""&ID&"","ProId"),"pItemC")%>" readonly /> </td>
								<%end if%>
								<%if pItemD = 1 then %>
								<td class="td_l_r title"><%=L_Order_Products_oProItemD%></td>
								<td class="td_r_l" <%if pItemC = 0 then %> colspan=3 <%end if%>> <input name="oProItemD" type="text" id="oProItemD" value="<%=EasyCrm.getNewItem("Products","ID",EasyCrm.getNewItem("Order_Products","osID",""&ID&"","ProId"),"pItemD")%>" readonly /> </td>
								<%end if%>
							</tr>
							<tr> 
								<%if pItemE = 1 then %>
								<td class="td_l_r title"><%=L_Order_Products_oProItemE%></td>
								<td class="td_r_l"> <input name="oProItemE" type="text" id="oProItemE" value="<%=EasyCrm.getNewItem("Products","ID",EasyCrm.getNewItem("Order_Products","osID",""&ID&"","ProId"),"pItemE")%>" readonly /> </td>
								<%end if%>
								<td class="td_l_r title"><%=L_Order_Products_oProPrice%></td>
								<td class="td_r_l" <%if pItemE = 0 then %> colspan=3 <%end if%>> <input name="oProPrice" type="text" id="oProPrice" size="10" onchange="oMoney.value=parseInt(oProPrice.value)*parseInt(oProNum.value)-parseInt(oDiscount.value)" value="<%=EasyCrm.getNewItem("Products","ID",EasyCrm.getNewItem("Order_Products","osID",""&ID&"","ProId"),"pUprice")%>" readonly /> <%=L_Yuan%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Order_Products_oProNum = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_Products_oProNum%></td>
								<td class="td_r_l"> <input name="oProNum" type="text" id="oProNum" size="10" onchange="oMoney.value=parseInt(oProPrice.value)*parseInt(oProNum.value)-parseInt(oDiscount.value)" value="<%=EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oProNum")%>" onfocus="if (value =='1'){value ='1'}"onblur="if (value ==''){value='<%=EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oProNum")%>'}" /> </td>
								<td class="td_l_r title"><%if Must_Order_Products_oDiscount = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_Products_oDiscount%></td>
								<td class="td_r_l"> <input name="oDiscount" type="text" id="oDiscount" size="10" onchange="oMoney.value=parseInt(oProPrice.value)*parseInt(oProNum.value)-parseInt(oDiscount.value)" value="<%=EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oDiscount")%>" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='<%=EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oDiscount")%>'}" /> <%=L_Yuan%> </td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Order_Products_oMoney%></td>
								<td class="td_r_l" COLSPAN="3"> <input name="oMoney" type="text" id="oMoney" size="10" style="font-weight:bold;color:Red;" class="" value="<%=EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oMoney")%>" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='<%=EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oMoney")%>'}" readonly /> <%=L_Yuan%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Order_Products_oContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Order_Products_oContent%></td>
								<td class="td_r_l" COLSPAN="3" style="padding:5px 10px;"> <textarea name="oContent" rows="4" id="oContent" class="int" style="height:50px;width:98%;"><%=EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oContent")%></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input name="osId" type="hidden" value="<%=ID%>">
							<input name="oMoneyOld" type="hidden" value="<%=EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oMoney")%>">
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
	<%
	elseif sType="SaveEdit" then '保存修改
		osID = Request.Form("osID")
		oProNum = Request.Form("oProNum")
		oDiscount = Request.Form("oDiscount")
		oMoney = Request.Form("oMoney")
		oMoneyOld = Request.Form("oMoneyOld")
		oContent = Request.Form("oContent")
		oID = EasyCrm.getNewItem("Order_Products","osID",""&osID&"","oID")
		cID = EasyCrm.getNewItem("Order","oID",""&oID&"","cID")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Order_Products] where osid="&osid,conn,3,2
		rs("oProNum") = oProNum
		rs("oDiscount") = oDiscount
		rs("oMoney") = oMoney
		rs("oContent") = oContent
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		'修改订单产品同时更新订单总金额
		conn.execute("update [Order] set oMoney = oMoney+'"&oMoney&"'-'"&oMoneyOld&"' where oId = "&oID&" ")
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','订单产品','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Delete" then '删除
		If Id = "" Then Exit Sub
		oID = EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oId")
		cID = EasyCrm.getNewItem("Order","oID",""&oID&"","cID")
		'先减去金额，再删除数据
		conn.execute("update [Order] set oMoney = oMoney-'"&EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oMoney")&"' where oId = "&EasyCrm.getNewItem("Order_Products","osID",""&ID&"","oId")&" ")
		conn.execute("DELETE FROM [Order_Products] where osId = "&Id&" ")
		'删除后判断是否存在产品，如果没有，则更新订单状态
		if EasyCrm.getCountItem("Order_Products","osID","osID"," and oId = "&oID&" ")=0 then
		conn.execute("update [Order] set oState = 0 where oId = "&oID&" ")
		end if
		
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','订单产品','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	end if
end Sub

Sub Hetong() '合同记录
%>
	<script language="JavaScript">
	<!-- 合同记录必填项提示
	function CheckInput()
	{
		if (<%=YnHTNum%>=="0"){if(document.all.hNum.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_hNum & alert04%>'});document.all.hNum.focus();return false;}}
		if (<%=Must_Hetong_oId%>=="1"){if(document.all.oId.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_oId & alert04%>'});document.all.oId.focus();return false;}}
		if (<%=Must_Hetong_hSdate%>=="1"){if(document.all.hSdate.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_hSdate & alert04%>'});document.all.hSdate.focus();return false;}}
		if (<%=Must_Hetong_hEdate%>=="1"){if(document.all.hEdate.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_hEdate & alert04%>'});document.all.hEdate.focus();return false;}}
		if (<%=Must_Hetong_hType%>=="1"){if(document.all.hType.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_hType & alert04%>'});document.all.hType.focus();return false;}}
		if (<%=Must_Hetong_hMoney%>=="1"){if(document.all.hMoney.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_hMoney & alert04%>'});document.all.hMoney.focus();return false;}}
		if (<%=Must_Hetong_hRevenue%>=="1"){if(document.all.hRevenue.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_hRevenue & alert04%>'});document.all.hRevenue.focus();return false;}}
		if (<%=Must_Hetong_hInvoice%>=="1"){if(document.all.hInvoice.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_hInvoice & alert04%>'});document.all.hInvoice.focus();return false;}}
		if (<%=Must_Hetong_hInvoice%>=="1"){if(document.all.hTax.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_hTax & alert04%>'});document.all.hTax.focus();return false;}}
		if (<%=Must_Hetong_hContent%>=="1"){if(document.all.hContent.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_hContent & alert04%>'});document.all.hContent.focus();return false;}}
	}
	-->
	</script>
	
	<%
	if sType="Choose" then '订单列表
	%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">订单列表 （单击选中）</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
        </td>
	</tr>
</table>
		<style>body{padding-top:35px;}</style>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<%if Order_oCode = 1 then %>
								<td class="td_l_c">订单编号</td>
								<%end if%>
								<%if Order_oLinkman = 1 then %>
								<td class="td_l_c">联系人</td>
								<%end if%>
								<%if Order_oSDate = 1 then %>
								<td class="td_l_c">下单日期</td>
								<%end if%>
								<%if Order_oEDate = 1 then %>
								<td class="td_l_c">交单日期</td>
								<%end if%>
								<%if Order_oDeposit = 1 then %>
								<td class="td_l_c">预付款</td>
								<%end if%>
								<td class="td_l_c">订单金额</td>
								<%if Order_oState = 1 then %>
								<td class="td_l_c">订单状态</td>
								<%end if%>
								<%if Order_oUser = 1 then %>
								<td class="td_l_c">业务员</td>
								<%end if%>
								<%if Order_oTime = 1 then %>
								<td class="td_l_c">录入时间</td>
								<%end if%>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Order] where cId = "&cId&" Order By oId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
							<tr class="tr" <%if rs("oState") = 0 or rs("oState") = 3 then%> onClick="art.dialog.tips('该订单不含【产品详情】或【已作废】，不可选择！', 3);"<%else%> onClick=window.location.href="javascript:$.dialog.open.origin.$('#oId').val('<%=rs("oid")%>');$.dialog.open.origin.$('#oCode').val('<%=rs("oCode")%>');$.dialog.open.origin.$('#hMoney').val('<%=rs("oMoney")%>');$.dialog.open.origin.$('#hRevenue').val('<%=rs("oDeposit")%>');$.dialog.close();"<%end if%> style="cursor:pointer;" >
								<%if Order_oCode = 1 then %>
								<td class="td_l_c"><%=rs("oCode")%></td>
								<%end if%>
								<%if Order_oLinkman = 1 then %>
								<td class="td_l_c"><%=rs("oLinkman")%></td>
								<%end if%>
								<%if Order_oSDate = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("oSDate"),2)%></td>
								<%end if%>
								<%if Order_oEDate = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("oEDate"),2)%></td>
								<%end if%>
								<%if Order_oDeposit = 1 then %>
								<td class="td_l_c"><%=rs("oDeposit")%></td>
								<%end if%>
								<td class="td_l_c"><%=rs("oMoney")%></td>
								<%if Order_oState = 1 then %>
								<td class="td_l_c"><%if rs("oState") = 0 then%>未处理<%elseif rs("oState") = 1 then%>处理中<%elseif rs("oState") = 2 then%>已完成<%elseif rs("oState") = 3 then%>已取消<%end if%></td>
								<%end if%>
								<%if Order_oUser = 1 then %>
								<td class="td_l_c"><%=rs("oUser")%></td>
								<%end if%>
								<%if Order_oTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("oTime"),2)%></td>
								<%end if%>
							</tr>
						<%
						rs.MoveNext
						Loop
							else
							%>
							<tr><td class="td_l_l" colspan=12><%=L_Notfound%></td></tr>
							<%
							end if
						rs.Close
						Set rs = Nothing
						%>
							
						</table>
					</td>
				</tr>	
			</table>

	<%
	elseif sType="Add" then '添加
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Hetong&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="100" /><col  /><col width="100" /><col  />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>新增合同记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Hetong_hNum%></td>
								<td class="td_r_l" colspan=3> <input name="hNum" type="text" id="hNum" size="23" class="int" <%if YnHTNum=1 then%> value="HT<%=EasyCrm.FormatDate(now(),14)&right(FormatNumber(timer(),3),3)%>" readonly style="border:0;" <%end if%> /> <span class="info_help help01" >&nbsp;手动模式请按固定规则编号</span></td>
							</tr>
							<tr> 
								<td class="td_l_r title"> <%if Must_Hetong_hType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hType%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_Hetong","hType","") %>&nbsp;<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Hetong_InfoAdd()' style="cursor:pointer"><script>function Select_Hetong_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Hetong', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%></td>
								<td class="td_l_r title"><%if Must_Hetong_oId = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_oId%></td>
								<td class="td_r_l"> 
									<input name="oId" type="hidden" id="oId" size="23" value="" readonly />
									<input name="oCode" type="text" id="oCode" size="23" class="int" value="" /> 
									<input name="Back" type="button" id="Back" class="button221" value="…" title="请选择" onclick='Choose_Order()' style="cursor:pointer"><script>function Choose_Order() {$.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=Choose&cid=<%=cID%>', {title: '新窗口', width: 900, height: 480,fixed: true}); };</script>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Hetong_hSdate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hSdate%></td>
								<td class="td_r_l"> <input name="hSdate" type="text" id="hSdate" class="Wdate" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(date(),2)%>" /></td>
								<td class="td_l_r title"> <%if Must_Hetong_hEdate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hEdate%></td>
								<td class="td_r_l"> <input name="hEdate" type="text" id="hEdate" class="Wdate" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="" /></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Hetong_hRevenue = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hRevenue%></td>
								<td class="td_r_l"> <input name="hRevenue" type="text" id="hRevenue" class="int" size="10" value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%></td>
								<td class="td_l_r title"><%if Must_Hetong_hMoney = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hMoney%></td>
								<td class="td_r_l"> <input name="hMoney" type="text" id="hMoney" class="int" size="10" value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"> <%if Must_Hetong_hInvoice = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hInvoice%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_YN","hInvoice","") %></td>
								<td class="td_l_r title"> <%if Must_Hetong_hInvoice = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hTax%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_YN","hTax","") %></td>
							</tr> 
							
								<%
								Set rss = Server.CreateObject("ADODB.Recordset")
								rss.Open "Select * From [CustomField] where cTable='Hetong' order by Id asc ",conn,3,1
								If rss.RecordCount > 0 Then
								i=1
								Do While Not rss.BOF And Not rss.EOF
								if i mod 2 = 1 then Response.Write "<tr>"
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											next
											%>
										<%end if%>
									</td>
								<%
								i = i + 1
								if i mod 2 = 1 then Response.Write "</tr>"
								rss.MoveNext
								Loop
								end if
								rss.Close
								Set rss = Nothing
								%>
							<tr> 
								<td class="td_l_r title"><%if Must_Hetong_hContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="hContent" rows="4" id="hContent" class="int" style="height:50px;width:95%;"></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="cID" type="hidden" value="<%=cID%>">
			<input name="hUser" type="hidden" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAdd" then '保存添加
		cID = Request.Form("cID")
		oId = Request.Form("oId")
		hNum = Request.Form("hNum")
		hType = Request.Form("hType")
		hSdate = Request.Form("hSdate")
		hEdate = Request.Form("hEdate")
		hRevenue = Request.Form("hRevenue")
		hMoney = Request.Form("hMoney")
		hInvoice = Request.Form("hInvoice")
		hTax = Request.Form("hTax")
		hContent = Request.Form("hContent")
		hUser = Request.Form("hUser")
		cMoney = Request.Form("hMoney")
		cRevenue = Request.Form("hRevenue")
		chContent = Request.Form("hContent")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
		
		'判断合同编号是否重复
		rs.Open "Select * From [Hetong] Where hNum = '" & hNum & "' ",conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='../Main/GetUpdateRW.asp?action=Hetong&sType=Add&tipinfo="&L_Hetong_hNum&alert02&"';</script>")
		Response.End()
		End If
		rs.Close
		
    	rs.Open "Select Top 1 * From [Hetong]",conn,3,2
		rs.AddNew
		rs("cID") = cID
		if oId <> "" then
		rs("oId") = oId
		end if
		rs("hNum") = hNum
		rs("hType") = hType
		if hSdate <> "" then
		rs("hSdate") = hSdate
		end if
		if hEdate <> "" then
		rs("hEdate") = hEdate
		end if
		if hRevenue <> "" then
		rs("hRevenue") = hRevenue
		end if
		if hMoney <> "" then
		rs("hMoney") = hMoney
		end if
		rs("hOwed") = hMoney - hRevenue
		rs("hInvoice") = hInvoice
		rs("hTax") = hTax
		rs("hContent") = hContent
		rs("hState") = "待审"
		rs("hUser") = hUser
		rs("hTime") = now()
    	rs.Update
    	rs.Close
    	Set rs = Nothing

	Dim rsid
	Set rsid = Server.CreateObject("ADODB.Recordset")
	if Accsql = 0 then
	rsid.Open "Select top 1 hID From [Hetong] order by hID desc",conn,1,1
	elseif Accsql = 1 then
	rsid.Open "Select @@IDENTITY as hID From [Hetong]",conn,1,1
	end if
	hID=rsid("hID")
	rsid.close
	
	'插入自定义内容
	
	cContent = ""
	Set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open "Select * From [CustomField] where cTable='Hetong' order by Id asc ",conn,1,1
	If rsc.RecordCount > 0 Then
	Do While Not rsc.BOF And Not rsc.EOF
	
	cContent = cContent & rsc("cName") &":"& Trim(Request(rsc("cName"))) &"|"
	
	rsc.MoveNext
	Loop
	end if
	rsc.Close
	Set rsc = Nothing
	
	conn.execute ("insert into CustomFieldContent(cID,hID,cContent) values('"&cid&"','"&hID&"','"&cContent&"')")	
		
		'同步更新订单为已完成（不可编辑）状态
		if oId <> "" then
		conn.execute("update [Order] set oState = 2 where oId = "&oId&" ")
		end if
		
		if Request.Form("hEdate") <> "" then
		conn.execute("update [Client] set cHEdate='"&Request.Form("hEdate")&"' where cId = "&cId&" ")
		end if
		
		'插入一条费用记录
		conn.execute ("insert into Expense(cId,eDate,eOutIn,eType,eMoney,eContent,eUser,eTime) values('"&cId&"','"&hSdate&"',1,'合同款','"&cRevenue&"','"&chContent&"','"&hUser&"','"&now()&"') ")	
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Hetong&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Edit" then '修改
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Hetong&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="100" /><col  /><col width="100" /><col  />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>修改合同记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Hetong_hNum%></td>
								<td class="td_r_l" COLSPAN="3"> <%=EasyCrm.getNewItem("Hetong","hID",""&ID&"","hNum")%> <span class="info_help help01" >&nbsp;不可修改</span></td>
							</tr>
							<tr> 
								<td class="td_l_r title"> <%if Must_Hetong_hType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hType%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_Hetong","hType",EasyCrm.getNewItem("Hetong","hID",""&ID&"","hType")) %>&nbsp;<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Hetong_InfoAdd()' style="cursor:pointer"><script>function Select_Hetong_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Hetong', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%></td>
								<td class="td_l_r title"><%if Must_Hetong_oId = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_oId%> </td>
								<td class="td_r_l"> 
									<input name="oId" type="hidden" id="oId" size="23" value="<%=EasyCrm.getNewItem("Hetong","hID",""&ID&"","oId")%>" readonly />
									<%if EasyCrm.getNewItem("Hetong","hID",""&ID&"","oId") <> "" then%>
									<input name="oCode" type="text" id="oCode" size="23" class="int" value="<%=EasyCrm.getNewItem("Order","oID",""&EasyCrm.getNewItem("Hetong","hID",""&ID&"","oId")&"","oCode")%>" readonly onclick='Choose_Order()' style="cursor:pointer" /> 
									<%else%>
									<input name="oCode" type="text" id="oCode" size="23" class="int" value="" readonly onclick='Choose_Order()' style="cursor:pointer" /> 
									<%end if%>
									<input name="Back" type="button" id="Back" class="button221" value="…" title="请选择" onclick='Choose_Order()' style="cursor:pointer"><script>function Choose_Order() {$.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=Choose&cid=<%=EasyCrm.getNewItem("Hetong","hID",""&ID&"","cId")%>', {title: '新窗口', width: 900, height: 480,fixed: true}); };</script>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Hetong_hSdate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hSdate%></td>
								<td class="td_r_l"> <input name="hSdate" type="text" id="hSdate" class="Wdate" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Hetong","hID",""&ID&"","hSdate"),2)%>" /></td>
								<td class="td_l_r title"> <%if Must_Hetong_hEdate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hEdate%></td>
								<td class="td_r_l"> <input name="hEdate" type="text" id="hEdate" class="Wdate" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Hetong","hID",""&ID&"","hEdate"),2)%>" /></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Hetong_hRevenue = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hRevenue%></td>
								<td class="td_r_l"> <%=EasyCrm.getNewItem("Hetong","hID",""&ID&"","hRevenue")%> <%=L_Yuan%>　
									<input name="Back" type="button" id="Back" class="button222" value="新增" onclick='Hetong_Revenue_InfoAdd()' style="cursor:pointer"><script>function Hetong_Revenue_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=AddRevenue&ID=<%=ID%>', {title: '新窗口', width: 400, height: 200,fixed: true}); };</script></td>
								<td class="td_l_r title"><%if Must_Hetong_hMoney = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hMoney%></td>
								<td class="td_r_l"> <input name="hMoney" type="text" id="hMoney" class="int" size="10" value="<%=EasyCrm.getNewItem("Hetong","hID",""&ID&"","hMoney")%>" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"> <%if Must_Hetong_hInvoice = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hInvoice%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_YN","hInvoice",EasyCrm.getNewItem("Hetong","hID",""&ID&"","hInvoice")) %></td>
								<td class="td_l_r title"> <%if Must_Hetong_hInvoice = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hTax%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_YN","hTax","'"&EasyCrm.getNewItem("Hetong","hID",""&ID&"","hTax")&"'") %></td>
							</tr> 
							<%
								cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID"," "&EasyCrm.getNewItem("Hetong","hID",""&ID&"","cID")&" And hID = "&ID&" ","cContent")
								cContentArr = split(cContentStr,"|")								
								Set rss = Server.CreateObject("ADODB.Recordset")
								rss.Open "Select * From [CustomField] where cTable='Hetong' order by Id asc ",conn,1,1
								If rss.RecordCount > 0 Then
								i=1:k=0
								Do While Not rss.BOF And Not rss.EOF
								if i mod 2 = 1 then Response.Write "<tr>"
								if Ubound(cContentArr) > k then
								cContent = split(cContentArr(k),":")
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
									<%if inStr(cContentArr(k),cContent(0))>0 then%>
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="<%=cContent(1)%>">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="<%=cContent(1)%>" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											if selectstr(selectarr) = cContent(1) then
											response.Write "<option value="""&selectstr(selectarr)&""" selected>"&selectstr(selectarr)&"</option>"
											else
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											end if
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											if checkboxstr(checkboxarr) = cContent(1) then
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&""" checked> "&checkboxstr(checkboxarr)&"　"
											else
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											end if
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											if radiostr(radioarr) = cContent(1) then
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&""" checked> "&radiostr(radioarr)&"　"
											else
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											end if
											next
											%>
										<%end if%>
									<%end if%>
									</td>
								<%
								else
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											next
											%>
										<%end if%>
									</td>
								<%
								end if
								i = i + 1:k=k+1
								if i mod 2 = 1 then Response.Write "</tr>"
								rss.MoveNext
								Loop
								end if
								rss.Close
								Set rss = Nothing
							%>
							<tr> 
								<td class="td_l_r title"><%if Must_Hetong_hContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Hetong_hContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="hContent" rows="4" id="hContent" class="int" style="height:50px;width:98%;"><%=EasyCrm.getNewItem("Hetong","hID",""&ID&"","hContent")%></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="hID" type="hidden" value="<%=ID%>">
			<input type="submit" name="Submit" class="button45" value="保存" <%if EasyCrm.getNewItem("Hetong","hID",""&ID&"","hState") = ""&L_Hetong_hState_2&"" then%> disabled<%end if%> >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveEdit" then '保存修改
		hID = Request.Form("hID")
		oId = Request.Form("oId")
		'hNum = Request.Form("hNum")
		hType = Request.Form("hType")
		hSdate = Request.Form("hSdate")
		hEdate = Request.Form("hEdate")
		hRevenue = EasyCrm.getNewItem("Hetong","hID",""&hID&"","hRevenue")
		hMoney = Request.Form("hMoney")
		hInvoice = Request.Form("hInvoice")
		hTax = Request.Form("hTax")
		hContent = Request.Form("hContent")
		cID = EasyCrm.getNewItem("Hetong","hID",""&hID&"","cID")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Hetong] where hid="&hid,conn,3,2
		if oId <> "" then
		rs("oId") = oId
		end if
		'rs("hNum") = hNum
		rs("hType") = hType
		if hSdate <> "" then
		rs("hSdate") = hSdate
		end if
		if hEdate <> "" then
		rs("hEdate") = hEdate
		end if
		if hRevenue <> "" then
		rs("hRevenue") = hRevenue
		end if
		if hMoney <> "" then
		rs("hMoney") = hMoney
		end if
		rs("hOwed") = hMoney - hRevenue
		rs("hInvoice") = hInvoice
		rs("hTax") = hTax
		rs("hContent") = hContent
    	rs.Update
    	rs.Close
    	Set rs = Nothing
	
	'更新自定义内容
	
	cContent = ""
	Set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open "Select * From [CustomField] where cTable='Hetong' order by Id asc ",conn,3,1
	If rsc.RecordCount > 0 Then
	Do While Not rsc.BOF And Not rsc.EOF
	'获取所有自定义字段
	cContent = cContent & rsc("cName") &":"& Trim(Request(rsc("cName"))) &"|" 
	rsc.MoveNext
	Loop
	end if
	rsc.Close
	Set rsc = Nothing
	if EasyCrm.getNewItem("CustomFieldContent","cID",""&cID&" and hID="&hID&" ","cContent")="0" then
	conn.execute ("insert into CustomFieldContent(cID,hID,cContent) values('"&cid&"','"&hID&"','"&cContent&"')")	
	else
	conn.execute ("UPDATE [CustomFieldContent] SET cContent='"&cContent&"' Where cId ="&cId&" and hID="&hID&" ")
	end if
		
		'同步更新订单为已完成（不可编辑）状态
		if oId <> "" then
		conn.execute("update [Order] set oState = 2 where oId = "&oId&" ")
		end if
		
		if Request.Form("hEdate") <> "" then
		conn.execute("update [Client] set cHEdate='"&Request.Form("hEdate")&"' where cId = "&cId&" ")
		end if
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Hetong&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>
	
	<%
	elseif sType="AddRevenue" then '新增合同到款
	%>
	<script language="JavaScript">
	<!-- 合同到款记录必填项提示
	function CheckInput()
	{
		if(document.all.rEdate.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_Renew_rEdate & alert04%>'});document.all.rEdate.focus();return false;}
	}
	-->
	</script>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Hetong&sType=SaveAddRevenue" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="100" /><col  />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>新增到款</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> 到款金额</td>
								<td class="td_r_l"> <input name="hRevenue" type="text" id="hRevenue" class="int" size="10" value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Hetong_hContent%></td>
								<td class="td_r_l" style="padding:5px 10px;"> <textarea name="eContent" rows="4" id="eContent" class="int" style="height:50px;width:95%;"></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="hID" type="hidden" value="<%=ID%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAddRevenue" then '保存添加
		hID = Request.Form("hID")
		cRevenue = Request.Form("hRevenue")
		eContent = Request.Form("eContent")
		cId = EasyCrm.getNewItem("Hetong","hID",""&hID&"","cId")
		
		'同步更新合同
		conn.execute("update [Hetong] set hRevenue = hRevenue+"&cRevenue&" where hId = "&hId&" ")
		conn.execute("update [Hetong] set hOwed=hMoney-hRevenue where hId = "&hId&" ")
		'同步更新费用记录
		conn.execute ("insert into Expense(cId,eDate,eOutIn,eType,eMoney,eContent,eUser,eTime) values('"&cId&"','"&now()&"',1,'合同款','"&cRevenue&"','"&eContent&"','"&Session("CRM_name")&"','"&now()&"') ")
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','合同到款','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	elseif sType="AddRenew" then '添加
	%>
	<script language="JavaScript">
	<!-- 合同续费记录必填项提示
	function CheckInput()
	{
		if(document.all.rEdate.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_Renew_rEdate & alert04%>'});document.all.rEdate.focus();return false;}
		if(document.all.rRevenue.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_Renew_rRevenue & alert04%>'});document.all.rRevenue.focus();return false;}
		if(document.all.rMoney.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_Renew_rMoney & alert04%>'});document.all.rMoney.focus();return false;}
	}
	-->
	</script>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Hetong&sType=SaveAddRenew" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  /><col width="120" /><col  /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>新增合同续费记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Hetong_hNum%></td>
								<td class="td_r_l" colspan=3> <%=EasyCrm.getNewItem("Hetong","hID",""&ID&"","hNum")%> </td>
							</tr>
							<tr> 
								<td class="td_l_r title"> 起始时间</td>
								<td class="td_r_l"> <%=EasyCrm.FormatDate(EasyCrm.getNewItem("Hetong","hID",""&ID&"","hSdate"),2)%></td>
								<td class="td_l_r title"> 到期时间</td>
								<td class="td_r_l"> <%=EasyCrm.FormatDate(EasyCrm.getNewItem("Hetong","hID",""&ID&"","hEdate"),2)%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"> 已收款</td>
								<td class="td_r_l"> <%=EasyCrm.getNewItem("Hetong","hID",""&ID&"","hRevenue")%></td>
								<td class="td_l_r title"> 总金额</td>
								<td class="td_r_l"> <%=EasyCrm.getNewItem("Hetong","hID",""&ID&"","hMoney")%></td>
							</tr>
							<tr>
								<td class="td_l_r title"><font color="#FF0000">*</font> 新到期时间</td>
								<td class="td_r_l" colspan=3> <input name="rEdate" type="text" id="rEdate" class="Wdate" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="" /> <span class="info_help help01">更新合同到期时间</span></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> <%=L_Hetong_Renew_rRevenue%></td>
								<td class="td_r_l"> <input name="rRevenue" type="text" id="rRevenue" class="int" size="10" value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%></td>
								<td class="td_l_r title"><font color="#FF0000">*</font> <%=L_Hetong_Renew_rMoney%></td>
								<td class="td_r_l"> <input name="rMoney" type="text" id="rMoney" class="int" size="10" value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Hetong_Renew_rContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="rContent" rows="4" id="rContent" class="int" style="height:50px;width:98%;"></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="hID" type="hidden" value="<%=ID%>">
			<input name="rUser" type="hidden" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAddRenew" then '保存添加
		hID = Request.Form("hID")
		rEdate = Request.Form("rEdate")
		rRevenue = Request.Form("rRevenue")
		rMoney = Request.Form("rMoney")
		rContent = Request.Form("rContent")
		rUser = Request.Form("rUser")
		cId = EasyCrm.getNewItem("Hetong","hID",""&hID&"","cId")
		
		'插入原始数据
		if EasyCrm.getCountItem("Hetong_Renew","rid","ridstr"," and hID="&hID&" ") = 0 then
		conn.execute ("insert into Hetong_Renew(hID,rEdate,rRevenue,rMoney,rState,rUser,rTime) values('"&hID&"','"&EasyCrm.getNewItem("Hetong","hID",""&hID&"","hEdate")&"','"&EasyCrm.getNewItem("Hetong","hID",""&hID&"","hRevenue")&"','"&EasyCrm.getNewItem("Hetong","hID",""&hID&"","hMoney")&"','"&EasyCrm.getNewItem("Hetong","hID",""&hID&"","hState")&"','"&EasyCrm.getNewItem("Hetong","hID",""&hID&"","hUser")&"','"&EasyCrm.getNewItem("Hetong","hID",""&hID&"","hTime")&"')")	
		end if
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
		
    	rs.Open "Select Top 1 * From [Hetong_Renew]",conn,3,2
		rs.AddNew
		rs("hID") = hID
		if rEdate <> "" then
		rs("rEdate") = rEdate
		end if
		if rRevenue <> "" then
		rs("rRevenue") = rRevenue
		end if
		if rMoney <> "" then
		rs("rMoney") = rMoney
		end if
		rs("rContent") = rContent
		rs("rState") = "待审"
		rs("rUser") = rUser
		rs("rTime") = now()
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		'同步更新合同
		conn.execute("update [Hetong] set hMoney = hMoney+"&rMoney&",hRevenue = hRevenue+"&rRevenue&", hEdate='"&rEdate&"' where hId = "&hId&" ")
		conn.execute("update [Hetong] set hOwed=hMoney-hRevenue where hId = "&hId&" ")
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','合同续费','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	elseif sType="EditRenew" then '修改
	%>
	<script language="JavaScript">
	<!-- 合同续费记录必填项提示
	function CheckInput()
	{
		if(document.all.rEdate.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_Renew_rEdate & alert04%>'});document.all.rEdate.focus();return false;}
		if(document.all.rRevenue.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_Renew_rRevenue & alert04%>'});document.all.rRevenue.focus();return false;}
		if(document.all.rMoney.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Hetong_Renew_rMoney & alert04%>'});document.all.rMoney.focus();return false;}
	}
	-->
	</script>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Save" action="?action=Hetong&sType=SaveEditRenew" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  /><col width="120" /><col  /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>修改合同续费记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Hetong_hNum%></td>
								<td class="td_r_l" colspan=3> <%=EasyCrm.getNewItem("Hetong","hID",EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","hID"),"hNum")%> </td>
							</tr>
							<tr> 
								<td class="td_l_r title"> 起始时间</td>
								<td class="td_r_l"> <%=EasyCrm.FormatDate(EasyCrm.getNewItem("Hetong","hID",EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","hID"),"hSdate"),2)%></td>
								<td class="td_l_r title"> 到期时间</td>
								<td class="td_r_l"> <%=EasyCrm.FormatDate(EasyCrm.getNewItem("Hetong","hID",EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","hID"),"hEdate"),2)%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"> 已收款</td>
								<td class="td_r_l"> <%=EasyCrm.getNewItem("Hetong","hID",EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","hID"),"hRevenue")%></td>
								<td class="td_l_r title"> 总金额</td>
								<td class="td_r_l"> <%=EasyCrm.getNewItem("Hetong","hID",EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","hID"),"hMoney")%></td>
							</tr>
							<tr>
								<td class="td_l_r title"><font color="#FF0000">*</font> 新到期时间</td>
								<td class="td_r_l" colspan=3> <input name="rEdate" type="text" id="rEdate" class="Wdate" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rEdate"),2)%>" /> <span class="info_help help01">更新合同到期时间</span></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> <%=L_Hetong_Renew_rRevenue%></td>
								<td class="td_r_l"> <input name="rRevenue" type="text" id="rRevenue" class="int" size="10" value="<%=EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rRevenue")%>" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%></td>
								<td class="td_l_r title"><font color="#FF0000">*</font> <%=L_Hetong_Renew_rMoney%></td>
								<td class="td_r_l"> <input name="rMoney" type="text" id="rMoney" class="int" size="10" value="<%=EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rMoney")%>" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Hetong_Renew_rContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="rContent" rows="4" id="rContent" class="int" style="height:50px;width:98%;"><%=EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rContent")%></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="rID" type="hidden" value="<%=ID%>">
			<input name="rRevenueOld" type="hidden" value="<%=EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rRevenue")%>">
			<input name="rMoneyOld" type="hidden" value="<%=EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rMoney")%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveEditRenew" then '保存添加
		rID = Request.Form("rID")
		rEdate = Request.Form("rEdate")
		rRevenue = Request.Form("rRevenue")
		rRevenueOld = Request.Form("rRevenueOld")
		rMoney = Request.Form("rMoney")
		rMoneyOld = Request.Form("rMoneyOld")
		rContent = Request.Form("rContent")
		hId = EasyCrm.getNewItem("Hetong_Renew","rID",""&rID&"","hID")
		cId = EasyCrm.getNewItem("Hetong","hID",""&hID&"","cId")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
		
    	rs.Open "Select Top 1 * From [Hetong_Renew] where rID= "&rID&" ",conn,3,2
		if rEdate <> "" then
		rs("rEdate") = rEdate
		end if
		if rRevenue <> "" then
		rs("rRevenue") = rRevenue
		end if
		if rMoney <> "" then
		rs("rMoney") = rMoney
		end if
		rs("rContent") = rContent
		rs("rState") = "待审"
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		'同步更新合同
		conn.execute("update [Hetong] set hMoney = hMoney+"&rMoney&"-"&rMoneyOld&",hRevenue = hRevenue+"&rRevenue&"-"&rRevenueOld&", hEdate='"&rEdate&"' where hId = "&hId&" ")
		conn.execute("update [Hetong] set hOwed=hMoney-hRevenue where hId = "&hId&" ")
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','合同续费','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
	elseif sType="RenewList" then '续费记录列表
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<td class="td_l_c">续费记录</td>
								<td class="td_l_c">到期时间</td>
								<td class="td_l_c">预付款</td>
								<td class="td_l_c">总金额</td>
								<td class="td_l_c">详情备注</td>
								<td class="td_l_c">续费状态</td>
							<%If mid(Session("CRM_qx"), 14, 1) = "1" Then%>
								<td class="td_l_c">审核</td>
							<%end if%>
								<td class="td_l_c">业务员</td>
								<td class="td_l_c">录入时间</td>
								<td class="td_l_c" width="90">管理</td>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select top 1 * From [Hetong_Renew] where hId = "&ID&" Order By rId asc ",conn,1,1
						Do While Not rs.BOF And Not rs.EOF
						%>
							<tr class="tr" >
								<td class="td_l_c">合同初始信息</td>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("rEdate"),2)%></td>
								<td class="td_l_c"><%=rs("rRevenue")%></td>
								<td class="td_l_c"><%=rs("rMoney")%></td>
								<td class="td_l_c"> - </td>
								<td class="td_l_c"> - </td>
								<td class="td_l_c"> - </td>
								<td class="td_l_c"> - </td>
								<td class="td_l_c"> - </td>
								<td class="td_l_c"> - </td>
							</tr>
						
						<%
						
						rs.MoveNext
						Loop
						rs.Close
						rs.Open "Select * From [Hetong_Renew] where hId = "&ID&" and rid not in ( select top 1 rid From [Hetong_Renew] where hId = "&ID&" Order By rId asc ) Order By rId asc ",conn,1,1
						i=0
						Do While Not rs.BOF And Not rs.EOF
						i=i+1
						%>
							<tr class="tr" >
								<td class="td_l_c">第 <%=i%> 次 续费</td>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("rEdate"),2)%></td>
								<td class="td_l_c"><%=rs("rRevenue")%></td>
								<td class="td_l_c"><%=rs("rMoney")%></td>
								<td class="td_l_c"><%if rs("rContent")<>"" then%><input type="button" class="button226" value="查看" onclick='Hetong_Renew_InfoView<%=rs("rId")%>()' style="cursor:pointer" /><%end if%></td>
								<td class="td_l_c"><%=rs("rState")%></td>
							<%If mid(Session("CRM_qx"), 14, 1) = "1" Then%>
								<td class="td_l_c">
									<input type="button" class="button222" value="审核"  onclick='Hetong_Renew_InfoAudit<%=rs("rId")%>()' style="cursor:pointer" />
								</td>
							<%end if%>
								<td class="td_l_c"><%=rs("rUser")%></td>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("rTime"),2)%></td>
								
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 38, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Hetong_Renew_InfoEdit<%=rs("rId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 39, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Hetong_Renew_InfoDel<%=rs("rId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Hetong_Renew_InfoEdit<%=rs("rId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=EditRenew&Id=<%=rs("rId")%>', {title: '编辑', width: 600,height: 340, fixed: true}); };</script>
							<script>function Hetong_Renew_InfoAudit<%=rs("rId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=AuditRenew&Id=<%=rs("rId")%>', {title: '审核', width: 400,height: 180, fixed: true}); };</script>
							<script>function Hetong_Renew_InfoDel<%=rs("rId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () { art.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=DeleteRenew&Id=<%=rs("rId")%>');return false;},cancel: true }); };</script>
							<script>function Hetong_Renew_InfoView<%=rs("rId")%>() {art.dialog({ title: '详情备注', content: '<%=EasyCrm.clearWord(""&rs("rContent")&"")%>',drag: false,resize: false}); };</script>
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
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="Back" type="button" id="Back" class="button45" value="新增" onclick='Hetong_Renew_InfoAdd()' style="cursor:pointer">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();art.dialog.open.origin.location.reload();">
		</td>
	</tr>
</table>
</div>
<script>function Hetong_Renew_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=AddRenew&Id=<%=Id%>', {title: '续费', width: 600,height: 340, fixed: true}); };</script>

	<%
	elseif sType="Audit" then '填写审核原因
	%>	
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="?action=Hetong&sType=SaveAudit&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						状态：
						<input name="hState" type="radio" value="<%=L_Hetong_hState_2%>" <%if EasyCrm.getNewItem("Hetong","hID",""&ID&"","hState") = ""&L_Hetong_hState_2&"" then%> checked<%end if%>> <%=L_Hetong_hState_2%>　 
						<input name="hState" type="radio" value="<%=L_Hetong_hState_3%>" <%if EasyCrm.getNewItem("Hetong","hID",""&ID&"","hState") = ""&L_Hetong_hState_3&"" then%> checked<%end if%>> <%=L_Hetong_hState_3%> 　 
						<input name="hState" type="radio" value="<%=L_Hetong_hState_1%>" <%if EasyCrm.getNewItem("Hetong","hID",""&ID&"","hState") = ""&L_Hetong_hState_1&"" then%> checked<%end if%>> 转<%=L_Hetong_hState_1%> 
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<textarea name="hAuditReasons" rows="4" id="hAuditReasons" class="int" style="height:80px;width:98%;"><%=EasyCrm.getNewItem("Hetong","hID",""&ID&"","hAuditReasons")%></textarea>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="hAudit" type="hidden" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAudit" then '保存
		If Id = "" Then Exit Sub
		hState = Request.Form("hState")
		cAuditReasons = Request.Form("hAuditReasons")
		cAudit = Request.Form("hAudit")
		conn.execute("update [Hetong] set hState = '"&hState&"',hAuditReasons = '"&cAuditReasons&"',hAudit = '"&cAudit&"',hAuditTime = '"&now()&"' where hId = "&ID&" ")
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
	elseif sType="AuditRenew" then '填写审核原因
	%>	
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="?action=Hetong&sType=SaveAuditRenew&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						状态：
						<input name="rState" type="radio" value="续费有效" <%if EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rState") = "续费有效" then%> checked<%end if%>> 续费有效　 
						<input name="rState" type="radio" value="续费无效" <%if EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rState") = "续费无效" then%> checked<%end if%>> 续费无效 　 
						<input name="rState" type="radio" value="待审" <%if EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rState") = "待审" then%> checked<%end if%>> 转待审 
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<textarea name="rAuditReasons" rows="4" id="rAuditReasons" class="int" style="height:80px;width:98%;"><%=EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rAuditReasons")%></textarea>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="rAudit" type="hidden" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAuditRenew" then '保存
		If Id = "" Then Exit Sub
		rState = Request.Form("rState")
		rAuditReasons = Request.Form("rAuditReasons")
		rAudit = Request.Form("rAudit")
		hId = EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","hId")
		'如果第一次续费，则更新两条记录
		if EasyCrm.getCountItem("Hetong_Renew","rId","IDstr"," and hID = "&hId&" ") =2 then
		conn.execute("update [Hetong_Renew] set rState = '"&rState&"',rAuditReasons = '"&rAuditReasons&"',rAudit = '"&rAudit&"',rAuditTime = '"&now()&"' where hId = "&hId&" ")
		else
		conn.execute("update [Hetong_Renew] set rState = '"&rState&"',rAuditReasons = '"&rAuditReasons&"',rAudit = '"&rAudit&"',rAuditTime = '"&now()&"' where rId = "&ID&" ")
		end if
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	
	elseif sType="DeleteRenew" then '删除
		If Id = "" Then Exit Sub
		hId = EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","hId")
		cId = EasyCrm.getNewItem("Hetong","hID",""&hID&"","cId")
		'同步更新合同
		conn.execute("update [Hetong] set hMoney = hMoney-"&EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rMoney")&",hRevenue = hRevenue-"&EasyCrm.getNewItem("Hetong_Renew","rID",""&ID&"","rRevenue")&" where hId = "&hId&" ")
		conn.execute("update [Hetong] set hOwed=hMoney-hRevenue where hId = "&hId&" ")
		
		conn.execute("DELETE FROM [Hetong_Renew] where rId = "&Id&" ")
		
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','合同续费','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
		
	elseif sType="DelReason" then '删除客户填写操作原因
	%>	
	<script language="JavaScript">
	<!-- 
	function CheckInput()
	{
		if(document.all.Reason.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=alert04%>'});document.all.Reason.focus();return false;}
	}
	-->
	</script>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="?action=Hetong&sType=Delete&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<textarea name="Reason" rows="4" id="Reason" class="int" style="height:80px;width:98%;"></textarea>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="Id" type="hidden" id="cId" value="<%=ID%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
		
	elseif sType="Delete" then '删除
		If Id = "" Then Exit Sub
		Reason = Trim(Request("Reason"))
		cId = EasyCrm.getNewItem("Hetong","hID",""&ID&"","cId")
		conn.execute("DELETE FROM [Hetong] where hId = "&Id&" ")
		
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lReason,lUser,lTime) values('"&cID&"','"&L_Hetong&"','"&L_insert_action_03&"','"&Reason&"','"&Session("CRM_name")&"','"&now()&"')")
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	end if
end Sub

Sub Service() '售后记录
%>
	<script language="JavaScript">
	<!-- 售后记录必填项提示
	function CheckInput()
	{
		if (<%=Must_Service_ProId%>=="1"){if(document.all.ProId.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Service_ProId & alert04%>'});document.all.ProId.focus();return false;}}
		if (<%=Must_Service_sTitle%>=="1"){if(document.all.sTitle.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Service_sTitle & alert04%>'});document.all.sTitle.focus();return false;}}
		if (<%=Must_Service_sType%>=="1"){if(document.all.sType.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Service_sType & alert04%>'});document.all.sType.focus();return false;}}
		if (<%=Must_Service_sLinkman%>=="1"){if(document.all.sLinkman.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Service_sLinkman & alert04%>'});document.all.sLinkman.focus();return false;}}
		if (<%=Must_Service_sSDate%>=="1"){if(document.all.sSDate.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Service_sSDate & alert04%>'});document.all.sSDate.focus();return false;}}
		if (<%=Must_Service_sContent%>=="1"){if(document.all.sContent.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Service_sContent & alert04%>'});document.all.sContent.focus();return false;}}
	}
	-->
	</script>
	<script>
	function Setdisabled(evt)
	{
		var evt=evt || window.event;   
		var e =evt.srcElement || evt.target;
		 if(e.value=="1")
		 {
			document.all.sInfo.disabled = false; document.all.sInfo.readOnly = false;
			document.all.sEDate.disabled = false; document.all.sEDate.readOnly = false;
			document.all.sEDate.classname = "Wdate int";document.all.sEDate.value = "<%=EasyCrm.FormatDate(date(),2)%>";
		 }
		 else
		 {
			document.all.sInfo.disabled = true; document.all.sInfo.readOnly = true;
			document.all.sEDate.disabled = true; document.all.sEDate.readOnly = true;
			document.all.sEDate.classname = "";document.all.sEDate.value = "";
		 }
	}
	</script>
	<%
	if sType="Add" then '添加
	%>
		<form name="Save" action="?action=Service&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>新增售后记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Service_ProId = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_ProId%></td>
								<td class="td_r_l" colspan=3> 
									<input name="ProId" type="hidden" id="ProId" size="23" value="" readonly />
									<input name="ProTitle" type="text" id="ProTitle" class="int" size="23" value="" readonly onclick='Choose_Products()' style="cursor:pointer" /> 
									<input name="Back" type="button" id="Back" class="button221" value="…" title="请选择" onclick='Choose_Products()' style="cursor:pointer"><script>function Choose_Products() {$.dialog.open('../Main/GetUpdateRW.asp?action=Choose&sType=Products&oType=Service&cID=<%=cID%>', {title: '新窗口', width: 900, height: 480,fixed: true}); };</script>
								</td>
							</tr>
							<tr> 
								<td class="td_l_r title"> <%if Must_Service_sTitle = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sTitle%></td>
								<td class="td_r_l"> <input name="sTitle" type="text" id="sTitle" class="int" size="30" value="" /> </td>
								<td class="td_l_r title"><%if Must_Service_sType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sType%></td>
								<td class="td_r_l" colspan=3> <% = EasyCrm.getSelect("SelectData","Select_Service","sType","") %>&nbsp;
									 <% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Service_InfoAdd()' style="cursor:pointer"><script>function Select_Service_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Service', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Service_sLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sLinkman%></td>
								<td class="td_r_l"> 
									<% = EasyCrm.getNewSelect("Linkmans","lName","sLinkman"," and cid="&cID&" ","") %>&nbsp;
									<input name="Back" type="button" id="Back" class="button222" value="新增" onclick='Linkmans_InfoAdd()' style="cursor:pointer"><script>function Linkmans_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Linkmans&sType=Add&cID=<%=cID%>', {title: '新窗口', width: 700, height: 340,fixed: true}); };</script>
								</td>
								<td class="td_l_r title"> <%if Must_Service_sSDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sSDate%></td>
								<td class="td_r_l"> <input name="sSDate" type="text" id="sSDate" class="Wdate int" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(date(),2)%>" /></td>
							</tr>
							
								<%
								Set rss = Server.CreateObject("ADODB.Recordset")
								rss.Open "Select * From [CustomField] where cTable='Service' order by Id asc ",conn,3,1
								If rss.RecordCount > 0 Then
								i=1
								Do While Not rss.BOF And Not rss.EOF
								if i mod 2 = 1 then Response.Write "<tr>"
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											next
											%>
										<%end if%>
									</td>
								<%
								i = i + 1
								if i mod 2 = 1 then Response.Write "</tr>"
								rss.MoveNext
								Loop
								end if
								rss.Close
								Set rss = Nothing
								%>
							<tr> 
								<td class="td_l_r title"><%if Must_Service_sContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="sContent" rows="4" id="sContent" class="int" style="height:50px;width:98%;"></textarea></td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Service_sSolve = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sSolve%></td>
								<td class="td_r_l"> <input name="sSolve" type="radio" value="0" checked onclick="Setdisabled()"> <%=L_Service_sSolve_0%>　 <input name="sSolve" type="radio" value="1" onclick="Setdisabled()"> <%=L_Service_sSolve_1%>
								</td>
								<td class="td_l_r title"> <%if Must_Service_sEDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sEDate%></td>
								<td class="td_r_l"> <input name="sEDate" type="text" id="sEDate" class="int" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="" disabled readOnly /></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Service_sInfo = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sInfo%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="sInfo" rows="4" id="sInfo" class="int" style="height:50px;width:98%;" disabled readOnly></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="cID" type="hidden" value="<%=cID%>">
			<input name="sUser" type="hidden" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAdd" then '保存添加
		cID = Request.Form("cID")
		ProId = Request.Form("ProId")
		sTitle = Request.Form("sTitle")
		sType = Request.Form("sType")
		sLinkman = Request.Form("sLinkman")
		sSDate = Request.Form("sSDate")
		sEDate = Request.Form("sEDate")
		sContent = Request.Form("sContent")
		sSolve = Request.Form("sSolve")
		sInfo = Request.Form("sInfo")
		sUser = Request.Form("sUser")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Service]",conn,3,2
		rs.AddNew
		rs("cID") = cID
		if ProId <> "" then
		rs("ProId") = ProId
		end if
		rs("sTitle") = sTitle
		rs("sType") = sType
		rs("sLinkman") = sLinkman
		if sSDate<>"" then
		rs("sSDate") = sSDate
		end if
		if sEDate<>"" then
		rs("sEDate") = sEDate
		end if
		rs("sContent") = sContent
		rs("sSolve") = sSolve
		rs("sInfo") = sInfo
		rs("sUser") = sUser
		rs("sTime") = now()
    	rs.Update
    	rs.Close
    	Set rs = Nothing

	Dim rsid
	Set rsid = Server.CreateObject("ADODB.Recordset")
	if Accsql = 0 then
	rsid.Open "Select top 1 sID From [Service] order by sID desc",conn,1,1
	elseif Accsql = 1 then
	rsid.Open "Select @@IDENTITY as sID From [Service]",conn,1,1
	end if
	sID=rsid("sID")
	rsid.close
	
	'插入自定义内容
	
	cContent = ""
	Set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open "Select * From [CustomField] where cTable='Service' order by Id asc ",conn,1,1
	If rsc.RecordCount > 0 Then
	Do While Not rsc.BOF And Not rsc.EOF
	
	cContent = cContent & rsc("cName") &":"& Trim(Request(rsc("cName"))) &"|"
	
	rsc.MoveNext
	Loop
	end if
	rsc.Close
	Set rsc = Nothing
	
	conn.execute ("insert into CustomFieldContent(cID,sID,cContent) values('"&cid&"','"&sID&"','"&cContent&"')")	
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Service&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Edit" then '修改
	%>
		<form name="Save" action="?action=Service&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>修改售后记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Service_ProId = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_ProId%></td>
								<td class="td_r_l" COLSPAN="3"> 
									<input name="ProId" type="hidden" id="ProId" size="23" value="<%=EasyCrm.getNewItem("Service","sID",""&ID&"","ProId")%>" readonly />
									<%if EasyCrm.getNewItem("Service","sID",""&ID&"","ProId") <>"" then%>
									<input name="ProTitle" type="text" id="ProTitle" class="int" size="23" value="<%=EasyCrm.getNewItem("Products","ID",EasyCrm.getNewItem("Service","sID",""&ID&"","ProId"),"pTitle")%>" /> 
									<%else%>
									<input name="ProTitle" type="text" id="ProTitle" class="int" size="23" value="" /> 
									<%end if%>
									<input name="Back" type="button" id="Back" class="button221" value="…" title="请选择" onclick='Choose_Products()' style="cursor:pointer"><script>function Choose_Products() {$.dialog.open('../Main/GetUpdateRW.asp?action=Choose&sType=Products&oType=Service&cID=<%=EasyCrm.getNewItem("Service","sID",""&ID&"","cID")%>', {title: '新窗口', width: 900, height: 480,fixed: true}); };</script>
								</td>
							</tr>
							<tr> 
								<td class="td_l_r title"> <%if Must_Service_sTitle = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sTitle%></td>
								<td class="td_r_l"> <input name="sTitle" type="text" id="sTitle" class="int" size="30" value="<%=EasyCrm.getNewItem("Service","sID",""&ID&"","sTitle")%>" /> </td>
								<td class="td_l_r title"><%if Must_Service_sType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sType%></td>
								<td class="td_r_l" colspan=3> <% = EasyCrm.getSelect("SelectData","Select_Service","sType",EasyCrm.getNewItem("Service","sID",""&ID&"","sType")) %>
									 <% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Service_InfoAdd()' style="cursor:pointer"><script>function Select_Service_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Service', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Service_sLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sLinkman%></td>
								<td class="td_r_l"> 
									<% = EasyCrm.getNewSelect("Linkmans","lName","sLinkman"," and cid="&EasyCrm.getNewItem("Service","sID",""&ID&"","cID")&" ",EasyCrm.getNewItem("Service","sID",""&ID&"","sLinkman")) %>&nbsp;
									<input name="Back" type="button" id="Back" class="button222" value="新增" onclick='Linkmans_InfoAdd()' style="cursor:pointer"><script>function Linkmans_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Linkmans&sType=Add&cID=<%=EasyCrm.getNewItem("Service","sID",""&ID&"","cID")%>', {title: '新窗口', width: 700, height: 340,fixed: true}); };</script>
								</td>
								<td class="td_l_r title"> <%if Must_Service_sSDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sSDate%></td>
								<td class="td_r_l"> <input name="sSDate" type="text" id="sSDate" class="Wdate int" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Service","sID",""&ID&"","sSDate"),2)%>" /></td>
							</tr>
							<%
								cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID"," "&EasyCrm.getNewItem("Service","sID",""&ID&"","cID")&" And sID = "&ID&" ","cContent")
								cContentArr = split(cContentStr,"|")								
								Set rss = Server.CreateObject("ADODB.Recordset")
								rss.Open "Select * From [CustomField] where cTable='Service' order by Id asc ",conn,1,1
								If rss.RecordCount > 0 Then
								i=1:k=0
								Do While Not rss.BOF And Not rss.EOF
								if i mod 2 = 1 then Response.Write "<tr>"
								if Ubound(cContentArr) > k then
								cContent = split(cContentArr(k),":")
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
									<%if inStr(cContentArr(k),cContent(0))>0 then%>
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="<%=cContent(1)%>">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="<%=cContent(1)%>" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											if selectstr(selectarr) = cContent(1) then
											response.Write "<option value="""&selectstr(selectarr)&""" selected>"&selectstr(selectarr)&"</option>"
											else
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											end if
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											if checkboxstr(checkboxarr) = cContent(1) then
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&""" checked> "&checkboxstr(checkboxarr)&"　"
											else
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											end if
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											if radiostr(radioarr) = cContent(1) then
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&""" checked> "&radiostr(radioarr)&"　"
											else
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											end if
											next
											%>
										<%end if%>
									<%end if%>
									</td>
								<%
								else
								%>
									<td class="td_l_r title"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											next
											%>
										<%end if%>
									</td>
								<%
								end if
								i = i + 1:k=k+1
								if i mod 2 = 1 then Response.Write "</tr>"
								rss.MoveNext
								Loop
								end if
								rss.Close
								Set rss = Nothing
							%>
							<tr> 
								<td class="td_l_r title"><%if Must_Service_sContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="sContent" rows="4" id="sContent" class="int" style="height:50px;width:98%;"><%=EasyCrm.getNewItem("Service","sID",""&ID&"","sContent")%></textarea></td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Service_sSolve = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sSolve%></td>
								<td class="td_r_l"> <input name="sSolve" type="radio" value="0" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=0 then%>checked <%end if%> onclick="Setdisabled()"> <%=L_Service_sSolve_0%>　 <input name="sSolve" type="radio" value="1" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=1 then%>checked <%end if%> onclick="Setdisabled()"> <%=L_Service_sSolve_1%>
								</td>
								<td class="td_l_r title"> <%if Must_Service_sEDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sEDate%></td>
								<td class="td_r_l"> <input name="sEDate" type="text" id="sEDate" class="int" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Service","sID",""&ID&"","sEDate"),2)%>" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=0 then%>disabled readOnly<%end if%> /></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Service_sInfo = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sInfo%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="sInfo" rows="4" id="sInfo" class="int" style="height:50px;width:98%;" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=0 then%>disabled readOnly<%end if%> ><%=EasyCrm.getNewItem("Service","sID",""&ID&"","sInfo")%></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="sID" type="hidden" value="<%=ID%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveEdit" then '保存修改
		sID = Request.Form("sID")
		ProId = Request.Form("ProId")
		sTitle = Request.Form("sTitle")
		sType = Request.Form("sType")
		sLinkman = Request.Form("sLinkman")
		sSDate = Request.Form("sSDate")
		sEDate = Request.Form("sEDate")
		sSolve = Request.Form("sSolve")
		sContent = Request.Form("sContent")
		sInfo = Request.Form("sInfo")
		cId = EasyCrm.getNewItem("Service","sID",""&sID&"","cId")

    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Service] where sID="&sID,conn,3,2
		if ProId <> "" then
		rs("ProId") = ProId
		end if
		rs("sTitle") = sTitle
		rs("sType") = sType
		rs("sLinkman") = sLinkman
		rs("sSDate") = sSDate
		if sEDate <> "" then
		rs("sEDate") = sEDate
		end if
		rs("sSolve") = sSolve
		rs("sContent") = sContent
		rs("sInfo") = sInfo
    	rs.Update
    	rs.Close
    	Set rs = Nothing
	
	'更新自定义内容
	
	cContent = ""
	Set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open "Select * From [CustomField] where cTable='Service' order by Id asc ",conn,3,1
	If rsc.RecordCount > 0 Then
	Do While Not rsc.BOF And Not rsc.EOF
	'获取所有自定义字段
	cContent = cContent & rsc("cName") &":"& Trim(Request(rsc("cName"))) &"|" 
	rsc.MoveNext
	Loop
	end if
	rsc.Close
	Set rsc = Nothing
	if EasyCrm.getNewItem("CustomFieldContent","cID",""&cID&" and sID="&sID&" ","cContent")="0" then
	conn.execute ("insert into CustomFieldContent(cID,sID,cContent) values('"&cid&"','"&sID&"','"&cContent&"')")	
	else
	conn.execute ("UPDATE [CustomFieldContent] SET cContent='"&cContent&"' Where cId ="&cId&" and sID="&sID&" ")
	end if
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Service&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Audit" then '填写审核原因
	%>	
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="?action=Service&sType=SaveAudit&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col  /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>售后处理结果</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Service_ProId%></td>
								<td class="td_r_l" COLSPAN="3"> <%if EasyCrm.getNewItem("Service","sID",""&ID&"","ProId")<>"" then%><%=EasyCrm.getNewItem("Products","ID",EasyCrm.getNewItem("Service","sID",""&ID&"","ProId"),"pTitle")%><%end if%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Service_sTitle%></td>
								<td class="td_r_l"><%=EasyCrm.getNewItem("Service","sID",""&ID&"","sTitle")%></td>
								<td class="td_l_r title"><%=L_Service_sType%></td>
								<td class="td_r_l"><%=EasyCrm.getNewItem("Service","sID",""&ID&"","sType") %></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Service_sLinkman%></td>
								<td class="td_r_l"><% = EasyCrm.getNewItem("Service","sID",""&ID&"","sLinkman") %></td>
								<td class="td_l_r title"><%=L_Service_sSDate%></td>
								<td class="td_r_l"><%=EasyCrm.FormatDate(EasyCrm.getNewItem("Service","sID",""&ID&"","sSDate"),2)%></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Service_sContent%></td>
								<td class="td_r_l" colspan=3 style="height:65px;"> <%=EasyCrm.getNewItem("Service","sID",""&ID&"","sContent")%></td>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Service_sSolve = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sSolve%></td>
								<td class="td_r_l"> <input name="sSolve" type="radio" value="0" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=0 then%>checked <%end if%> onclick="Setdisabled()"> <%=L_Service_sSolve_0%>　 <input name="sSolve" type="radio" value="1" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=1 then%>checked <%end if%> onclick="Setdisabled()"> <%=L_Service_sSolve_1%>
								</td>
								<td class="td_l_r title"> <%if Must_Service_sEDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sEDate%></td>
								<td class="td_r_l"> <input name="sEDate" type="text" id="sEDate" class="int" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Service","sID",""&ID&"","sEDate"),2)%>" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=0 then%>disabled readOnly<%end if%> /></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Service_sInfo = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sInfo%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="sInfo" rows="4" id="sInfo" class="int" style="height:50px;width:98%;" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=0 then%>disabled readOnly<%end if%> ><%=EasyCrm.getNewItem("Service","sID",""&ID&"","sInfo")%></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="sUser" type="hidden" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAudit" then '保存
		If Id = "" Then Exit Sub
		sSolve = Request.Form("sSolve")
		sEDate = Request.Form("sEDate")
		sInfo = Request.Form("sInfo")
		sUser = Request.Form("sUser")
		conn.execute("update [Service] set sSolve = '"&sSolve&"',sEDate = '"&sEDate&"',sInfo = '"&sInfo&"',sUser = '"&sUser&"' where sId = "&ID&" ")
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
	elseif sType="DelReason" then '删除客户填写操作原因
	%>	
	<script language="JavaScript">
	<!-- 
	function CheckInput()
	{
		if(document.all.Reason.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=alert04%>'});document.all.Reason.focus();return false;}
	}
	-->
	</script>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="?action=Service&sType=Delete&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<textarea name="Reason" rows="4" id="Reason" class="int" style="height:80px;width:98%;"></textarea>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="Id" type="hidden" id="cId" value="<%=ID%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
		
	elseif sType="Delete" then '删除
		If Id = "" Then Exit Sub
		Reason = Trim(Request("Reason"))
		cId = EasyCrm.getNewItem("Service","sID",""&ID&"","cId")
		conn.execute("DELETE FROM [Service] where sID = "&Id&" ")
		
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lReason,lUser,lTime) values('"&cID&"','"&L_Service&"','"&L_insert_action_03&"','"&Reason&"','"&Session("CRM_name")&"','"&now()&"')")
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	end if
end Sub

Sub Expense() '费用记录
	eOutIn = Trim(Request("eOutIn"))
%>
	<script language="JavaScript">
	<!-- 费用记录必填项提示
	function CheckInput()
	{
		if (<%=Must_Expense_eDate%>=="1"){if(document.all.eDate.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Expense_eDate & alert04%>'});document.all.eDate.focus();return false;}}
		if (<%=Must_Expense_eType%>=="1"){if(document.all.eType.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Expense_eType & alert04%>'});document.all.eType.focus();return false;}}
		if (<%=Must_Expense_eMoney%>=="1"){if(document.all.eMoney.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Expense_eMoney & alert04%>'});document.all.eMoney.focus();return false;}}
		if (<%=Must_Expense_eContent%>=="1"){if(document.all.eContent.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Expense_eContent & alert04%>'});document.all.eContent.focus();return false;}}
	}
	-->
	</script>
	<%
	if sType="Add" then '添加
	%>
		<form name="Save" action="?action=Expense&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>新增<%if eOutIn = 1 then %>收入<%else%>支出<%end if%>记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Expense_eDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eDate%></td>
								<td class="td_r_l"> <input name="eDate" type="text" id="eDate" class="Wdate int" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(date(),2)%>" /></td>
							</tr>
							<tr> 
								<td class="td_l_r title"> <%if Must_Expense_eType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eType%></td>
								<%if eOutIn = 1 then%>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_ExpenseIN","eType","") %>&nbsp;
									<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_ExpenseIN_InfoAdd()' style="cursor:pointer"><script>function Select_ExpenseIN_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_ExpenseIN', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%>
								</td>
								<%else%>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_ExpenseOUT","eType","") %>&nbsp;
									<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_ExpenseOUT_InfoAdd()' style="cursor:pointer"><script>function Select_ExpenseOUT_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_ExpenseOUT', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%>
								</td>
								<%end if%>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Expense_eMoney = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eMoney%></td>
								<td class="td_r_l"> <input name="eMoney" type="text" id="eMoney" class="int" size="10" value="0" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%> </td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Expense_eContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="eContent" rows="4" id="eContent" class="int" style="height:50px;width:98%;"></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="cID" type="hidden" value="<%=cID%>">
			<input name="eOutIn" type="hidden" value="<%=eOutIn%>" />
			<input name="eUser" type="hidden" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAdd" then '保存添加
		cID = Request.Form("cID")
		eDate = Request.Form("eDate")
		eOutIn = Request.Form("eOutIn")
		eType = Request.Form("eType")
		eMoney = Request.Form("eMoney")
		eContent = Request.Form("eContent")
		eUser = Request.Form("eUser")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Expense]",conn,3,2
		rs.AddNew
		rs("cID") = cID
		if eDate<>"" then
		rs("eDate") = eDate
		end if
		rs("eOutIn") = eOutIn
		rs("eType") = eType
		rs("eMoney") = eMoney
		rs("eContent") = eContent
		rs("eUser") = eUser
		rs("eTime") = now()
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Expense&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Edit" then '修改
	%>
		<form name="Save" action="?action=Expense&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="4"><B>修改<%if eOutIn = 1 then %>收入<%else%>支出<%end if%>记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Expense_eDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eDate%></td>
								<td class="td_r_l"> <input name="eDate" type="text" id="eDate" class="Wdate int" size="15" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Expense","eID",""&ID&"","eDate"),2)%>" /></td>
							</tr>
							<tr> 
								<td class="td_l_r title"> <%if Must_Expense_eType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eType%></td>
								<%if eOutIn = 1 then%>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_ExpenseIN","eType",EasyCrm.getNewItem("Expense","eID",""&ID&"","eType")) %>&nbsp;
									<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_ExpenseIN_InfoAdd()' style="cursor:pointer"><script>function Select_ExpenseIN_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_ExpenseIN', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%>
								</td>
								<%else%>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_ExpenseOUT","eType",EasyCrm.getNewItem("Expense","eID",""&ID&"","eType")) %>&nbsp;
									<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_ExpenseOUT_InfoAdd()' style="cursor:pointer"><script>function Select_ExpenseOUT_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_ExpenseOUT', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%>
								</td>
								<%end if%>
							</tr>
							<tr>
								<td class="td_l_r title"> <%if Must_Expense_eMoney = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eMoney%></td>
								<td class="td_r_l"> <input name="eMoney" type="text" id="eMoney" class="int" size="10" value="<%=EasyCrm.getNewItem("Expense","eID",""&ID&"","eMoney")%>" onfocus="if (value =='0'){value =''}"onblur="if (value ==''){value='0'}" /> <%=L_Yuan%> </td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%if Must_Expense_eContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eContent%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> <textarea name="eContent" rows="4" id="eContent" class="int" style="height:50px;width:98%;"><%=EasyCrm.getNewItem("Expense","eID",""&ID&"","eContent")%></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="eID" type="hidden" value="<%=ID%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveEdit" then '保存修改
		eID = Request.Form("eID")
		eDate = Request.Form("eDate")
		eType = Request.Form("eType")
		eMoney = Request.Form("eMoney")
		eContent = Request.Form("eContent")
		cId = EasyCrm.getNewItem("Expense","eID",""&eID&"","cId")

    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [Expense] where eID="&eID,conn,3,2
		if eDate<>"" then
		rs("eDate") = eDate
		end if
		rs("eType") = eType
		rs("eMoney") = eMoney
		rs("eContent") = eContent
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Expense&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	elseif sType="DelReason" then '删除客户填写操作原因
	%>	
	<script language="JavaScript">
	<!-- 
	function CheckInput()
	{
		if(document.all.Reason.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=alert04%>'});document.all.Reason.focus();return false;}
	}
	-->
	</script>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="?action=Expense&sType=Delete&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<textarea name="Reason" rows="4" id="Reason" class="int" style="height:80px;width:98%;"></textarea>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="Id" type="hidden" id="cId" value="<%=ID%>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
		
	elseif sType="Delete" then '删除
		If Id = "" Then Exit Sub
		Reason = Trim(Request("Reason"))
		cId = EasyCrm.getNewItem("Expense","eID",""&ID&"","cId")
		conn.execute("DELETE FROM [Expense] where eID = "&Id&" ")
		
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lReason,lUser,lTime) values('"&cID&"','"&L_Expense&"','"&L_insert_action_03&"','"&Reason&"','"&Session("CRM_name")&"','"&now()&"')")
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	end if
end Sub

Sub File() '附件记录
%>
	<script language="JavaScript">
	<!-- 附件记录必填项提示
	function CheckInput()
	{
		if(document.all.fFile.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '未选择要上传的附件！'});document.all.fFile.focus();return false;}
	}
	-->
	</script>
	<%
	if sType="Add" then '添加
	%>
	        <%
				filefolder = Server.MapPath("../upload/Client_File/")
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				if not fso.FolderExists(filefolder) then
				fso.CreateFolder(filefolder) 
				end if
            %>
		<form name="Save" action="?action=File&sType=SaveAdd" method="post" enctype="multipart/form-data" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>新增附件记录</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> <%=L_Select%></td>
								<td class="td_r_l" style="padding:5px 10px;"> <input name="fFile" type="file" id="fFile" value="" class="int"></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_File_fContent%></td>
								<td class="td_r_l" style="padding:5px 10px;"> <textarea name="fContent" rows="4" id="fContent" class="int" style="height:50px;width:98%;"></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="cID" type="hidden" value="<%=cID%>">
			<input name="fTitle" type="hidden" id="fTitle" value="">
			<input name="fUser" type="hidden" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" onClick="fTitle.value=/[^\\]+\.\w+$/.exec(fFile.value)[0]" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveAdd" then '保存添加
	
		dim nTime : nTime = Timer()
		dim request,lngUpSize
		Set request=new UpLoadClass
		request.TotalSize= 104857600
		request.MaxSize  = 100000*1024
		request.FileType = ""&uploadtype&""
		request.Savepath = "../upload/Client_File/"
		lngUpSize = request.Open()
		    
		cID = Request.Form("cID")
		fTitle = Request.Form("fTitle")
		fUser = Request.Form("fUser")
		fFile = request.Savepath & Request.Form("fFile")
		if fFile = request.Savepath then fFile=""
		fContent = Request.Form("fContent")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [File] ",conn,3,2
		rs.AddNew
		rs("cId") = cId
		rs("fTitle") = fTitle
		rs("fUser") = fUser
		rs("fFile") = fFile
		rs("fContent") = fContent
		rs("fTime") = Now()
    	rs.Update
    	rs.Close
    	Set rs = Nothing 
		
		conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_File&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	elseif sType="Delete" then '删除
		If Id = "" Then Exit Sub
		cId = EasyCrm.getNewItem("File","fId",""&ID&"","cId")
		FileInfo = EasyCrm.getNewItem("File","fId",""&ID&"","fFile")
		conn.execute("DELETE FROM [File] where fId = "&Id&" ")
		
		If FileInfo <> "" Then
			Set fso = CreateObject("Scripting.FileSystemObject")
			IF fso.FileExists(server.MapPath(FileInfo)) Then
			fso.DeleteFile(server.MapPath(FileInfo))
			End IF
		End If
		
		'插入操作记录
		conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_File&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")	
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	%>

	<%
	end if
end Sub

Sub Choose() '列表选择

	PFid = Trim(Request("PFid"))
	PSid = Trim(Request("PSid"))
	
	sql=""
	if PFid <> "" then
	sql = sql & " and pBigClass='"& EasyCrm.getNewItem("ProductClass","pClassid",""&PFid&"","pClassname") &"' "
	end if
	if PSid <> "" then
	sql = sql & " and pSmallClass='"& EasyCrm.getNewItem("ProductClass","pClassid",""&PSid&"","pClassname") &"' "
	end if
	
	'售后选择产品仅限订单中包含的产品（是否已完成: 【 and oid in ( select oid from [Order] where cID="&cID&" and oState = 2 ) 】 ） 
	if oType = "Service" then
	sql = sql & " and Id IN ( Select ProID from [Order_Products] where cID="&cID&" ) "
	end if
%>
	<%
	if sType="Products" then '添加
	%>
	<script>function changese(obj){  window.location.href=obj.value }</script>

	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
		<tr>
			<td class="top_left td_t_n td_r_n">产品列表 （单击选中）
				　　按分类筛选：
				<select name="BigClass" onchange="changese(this)">
				<option value="?action=Choose&sType=Products">请选择</option>
				<% 
					Set rsa = Conn.Execute("select * from ProductClass where pClassFid = '0' ")
					If Not rsa.Eof then
					Do While Not rsa.Eof
				%>
					<option value="?action=Choose&sType=Products&oType=<%=oType%>&cID=<%=cID%>&PFid=<%=rsa("pClassid")%>" <%if PFid=""&rsa("pClassid")&"" then%>selected<%end if%>><%=rsa("pClassname")%></option>
				<%
					rsa.Movenext
					Loop
					End If
					rsa.Close
					Set rsa = Nothing 
				%>
				</select> 
				<%if PFid<>"" then%>
				<select name="SmallClass" onchange="changese(this)">
				<option value="?action=Choose&sType=Products&oType=<%=oType%>&cID=<%=cID%>&PFid=<%=PFid%>">请选择</option>
				<% 
				Set rsb = Conn.Execute("select * from ProductClass where pClassFid='"&PFid&"' ")
				If Not rsb.Eof then
				Do While Not rsb.Eof
				%>
				<option value="?action=Choose&sType=Products&oType=<%=oType%>&cID=<%=cID%>&PFid=<%=PFid%>&PSid=<%=rsb("pClassid")%>" <%if PSid=""&rsb("pClassid")&"" then%>selected<%end if%>><%=rsb("pClassname")%></option>
				<%rsb.Movenext
				Loop
				End If
				rsb.Close
				Set rsb = Nothing 
				%>
				</select> 
				<%end if%>
			
			</td>
			<td class="top_right td_t_n td_r_n">
				<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			</td>
		</tr>
	</table>
		<style>body{padding-top:35px;}</style>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<td width="100" class="td_l_c">产品大类</td>
								<td width="100" class="td_l_c">产品小类</td>
								<td class="td_l_l">产品名称</td>
								<%if pItemA = 1 then%>
								<td class="td_l_c"><%=L_Order_Products_oProItemA%></td>
								<%end if%>
								<%if pItemB = 1 then%>
								<td class="td_l_c"><%=L_Order_Products_oProItemB%></td>
								<%end if%>
								<%if pItemC = 1 then%>
								<td class="td_l_c"><%=L_Order_Products_oProItemC%></td>
								<%end if%>
								<%if pItemD = 1 then%>
								<td class="td_l_c"><%=L_Order_Products_oProItemD%></td>
								<%end if%>
								<%if pItemE = 1 then%>
								<td class="td_l_c"><%=L_Order_Products_oProItemE%></td>
								<%end if%>
								<td width="80" class="td_l_c">单价</td>
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
							<tr class="tr" onClick=window.location.href="javascript:$.dialog.open.origin.$('#ProId').val('<%=rs("id")%>');$.dialog.open.origin.$('#ProTitle').val('<%=rs("pTitle")%>');<%if oType="Order" then%>$.dialog.open.origin.$('#oProItemA').val('<%=rs("pItemA")%>');$.dialog.open.origin.$('#oProItemB').val('<%=rs("pItemB")%>');$.dialog.open.origin.$('#oProItemC').val('<%=rs("pItemC")%>');$.dialog.open.origin.$('#oProItemD').val('<%=rs("pItemD")%>');$.dialog.open.origin.$('#oProItemE').val('<%=rs("pItemE")%>');$.dialog.open.origin.$('#oProPrice').val('<%=rs("pUprice")%>');$.dialog.open.origin.$('#oMoney').val('<%=rs("pUprice")%>');<%end if%>$.dialog.close();" style="cursor:pointer;" >
								<td class="td_l_c"><%=rs("pBigClass")%></td>
								<td class="td_l_c"><%=rs("pSmallClass")%></td>
								<td class="td_l_l"><%=rs("pTitle")%></td>
								<%if pItemA = 1 then%>
								<td class="td_l_c"><%=rs("pItemA")%></td>
								<%end if%>
								<%if pItemB = 1 then%>
								<td class="td_l_c"><%=rs("pItemB")%></td>
								<%end if%>
								<%if pItemC = 1 then%>
								<td class="td_l_c"><%=rs("pItemC")%></td>
								<%end if%>
								<%if pItemD = 1 then%>
								<td class="td_l_c"><%=rs("pItemD")%></td>
								<%end if%>
								<%if pItemE = 1 then%>
								<td class="td_l_c"><%=rs("pItemE")%></td>
								<%end if%>
								<td class="td_l_c"><%=rs("pUprice")%></td>
							</tr>
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
				<%=EasyCrm.pagelist("?action=Choose&sType=Products", PN,TotalPages,TotalRecords)%>
			</td>
		</tr>
	</table>
	</div>

	<%
	end if
end Sub
%>
<script src="../data/calendar/WdatePicker.js"></script>
</body>
</html>
<% Set EasyCrm = nothing %>