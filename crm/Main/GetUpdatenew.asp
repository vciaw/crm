<!--#include file="../Data/Conn.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" oncontextmenu=self.event.returnValue=false><head>
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
cID = Trim(Request("cID"))
ID = Trim(Request("ID"))
tipinfo = Trim(Request("tipinfo"))
YNRange = Trim(Request("YNRange"))

'禁止直接打开子窗口
From_url = Cstr(Request.ServerVariables("HTTP_Referer"))
Serv_url = Cstr(Request.ServerVariables("Server_Name"))
If mid(From_url,8,len(Serv_url)) <> Serv_url Then
	Response.Write("<script>window.opener=null;window.close();</script>")
	Response.end
End If

Select Case action
Case "Customer"
    Call Customer()
End Select


Sub Customer() '客户档案
	cid = Trim(Request("cid"))
	if tipinfo<>"" then
		Response.Write("<script>art.dialog({title: 'Error',time: 1,icon: 'warning',content: '"&tipinfo&"'});</script>")
	end if
%>
<%
if sType="Add" then '添加
%>
<% If mid(Session("CRM_qx"), 17, 1) = 1 Then %>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Check.min.js"></script>
<style>body {padding-top:35px;padding-bottom:55px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Top_View_Company%>  (<font color="#FF0000">*</font>)</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
        </td>
	</tr>
</table>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Add" action="?action=Customer&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col width="300" /><col width="120" />
							<tr>
								<td class="td_l_r title"><%=L_Customer_cCompany%></td>
								<td class="td_r_l" colspan=3> 
								<input type="text" class="int" name="Company" id="Company" size="50" maxlength="50" autocomplete="off" onChange="checkcompany(this.value);" onkeyup="searchSuggest();"> <span id="check1"> <span class="info_warn help01"><%=L_Tip_Info_01%></span></span><div id="search_suggest" style="display:none"></div></td>
							</tr>
							
							
							<tr> 
								<td class="td_l_r title"><%=L_Customer_cTel%></td>
								<td class="td_r_l" > <input name="Tel" type="text" class="int" id="Tel" size="30"></td>
							    <td class="td_l_r title"><%=L_Customer_cType%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_Mtype","Type","") %>
								<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %><input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Type_InfoAdd()' style="cursor:pointer"><script>function Select_Type_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Type', {title: '新窗口', width: 400, height: 140,fixed: true}); };</script><%end if%></td>
							
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Customer_cHomepage%></td>
								<td class="td_r_l"> 
								<input name="Homepage" type="text" class="int" id="Homepage" size="35" > 
								</td>
							    <td class="td_l_r title"><%=L_Customer_cFax%></td>
								<td class="td_r_l"><input name="Fax" type="text" class="int" id="Fax" size="30"></td>
							
							</tr>

							<tr> 
								<td class="td_l_r title"><%=L_Customer_cInfo%></td>
								<td class="td_r_l" colspan=3>
								<input name="Info" type="text" class="int" id="Info" size="30">
								</td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Customer_cBeizhu%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> 
								<textarea name="Beizhu" id="Beizhu" class="int" style="height:50px;width:98%;">
								</textarea>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="User" type="hidden" value="<%=Session("CRM_name")%>">
			<input name="Group" type="hidden" value="<%=Session("CRM_group")%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
<%end if%>
<%
elseif sType="SaveAdd" then
	
	
	cCompany = Trim(Request("Company"))
	cArea = Trim(Request("Area"))
	cSquare = Trim(Request("Square"))
	cAddress = Trim(Request("Address"))
	cZip = Trim(Request("Zip"))
	cLinkman = Trim(Request("Linkman"))
	cZhiwei = Trim(Request("Zhiwei"))
	cMobile = Trim(Request("Mobile"))
	cTel = Trim(Request("Tel"))
	cFax = Trim(Request("Fax"))
	cHomepage = Trim(Request("Homepage"))
	cEmail = Trim(Request("Email"))  
	cTrade = Trim(Request("Trade"))
	cStrade = Trim(Request("Strade"))
	cType = Trim(Request("Type"))
	cStart = Trim(Request("Start"))
	cSource = Trim(Request("Source"))    
	cInfo = Trim(Request("Info"))
	cBeizhu = Trim(Request("Beizhu"))
	cUser = Trim(Request("User"))
	cGroup = Trim(Request("Group"))
	
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select Top 1 * From Customer",conn,3,2
	rs.AddNew
	rs("cCompany") = cCompany
	rs("cArea") = cArea
	rs("cSquare") = cSquare
	rs("cAddress") = cAddress
	rs("cZip") = cZip
	rs("cLinkman") = cLinkman
	rs("cZhiwei") = cZhiwei
	rs("cMobile") = cMobile
	rs("cTel") = cTel
	rs("cFax") = cFax
	rs("cHomepage") = cHomepage
	rs("cEmail") = cEmail
	rs("cTrade") = cTrade
	rs("cStrade") = cStrade
	rs("cType") = cType
	rs("cStart") = cStart
	rs("cSource") = cSource
	rs("cInfo") = cInfo
	rs("cBeizhu") = cBeizhu
	rs("cUser") = cUser
	rs("cGroup") = cGroup
	rs("cLastUpdated") = now()
	
	'写入默认值
	rs("cDate") = Date()
	rs("cYn") = 1
	rs("cShare") = 0

	rs.Update
	rs.Close
	Set rs = Nothing

	Dim rsid
	Set rsid = Server.CreateObject("ADODB.Recordset")
	if Accsql = 0 then
	rsid.Open "Select top 1 cid From Customer order by cid desc",conn,1,1
	elseif Accsql = 1 then
	rsid.Open "Select @@IDENTITY as cid From Customer",conn,1,1
	end if
	cid=rsid("cid")
	rsid.close
	
	'插入操作记录
	'conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Customer&"','"&L_insert_action_01&"','"&cUser&"','"&now()&"')")	

	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	
elseif sType="InfoEdit" then
%>
<% If mid(Session("CRM_qx"), 18, 1) = 1 Then %>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Check.min.js"></script>
<style>body {padding-top:35px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Top_View_Company%>  (<font color="#FF0000">*</font>)</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
        </td>
	</tr>
</table>
<script>function Setting_Customer_AddMust() {$.dialog.open('../system/GetUpdate.asp?action=Setting&sType=CustomerAddMust', {title: '自定义设置', width: 800, height: 480,fixed: true}); };</script>
	
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<form name="Add" action="?action=Customer&sType=SaveEdit" method="post"onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col width="300" /><col width="120" />
							<tr>
								<td class="td_l_r title"><%=L_Customer_cCompany%></td>
								<td class="td_r_l" colspan=3> 
								<input type="text" class="int" name="Company" id="Company" size="50" value="<%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cCompany")%>"  maxlength="50" > 
								</td>
							</tr>
							
							
							<tr> 
								<td class="td_l_r title"><%=L_Customer_cTel%></td>
								<td class="td_r_l" > 
								<input name="Tel" type="text" class="int" value="<%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cTel")%>" id="Tel" size="30">
								</td>
							    <td class="td_l_r title"><%=L_Customer_cType%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_Mtype","Type",""&EasyCrm.getNewItem("Customer","cID",""&cID&"","cType")&"") %>
								<% If mid(Session("CRM_qx"), 6, 1) = 1 Then %>
								<input name="Back" type="button" id="Back" class="button227" value="新增" onclick='Select_Type_InfoAdd()' style="cursor:pointer">
								<script>function Select_Type_InfoAdd() {$.dialog.open('../System/GetUpdate.asp?action=SelectData&sType=Add&oType=Select_Type', {title: '新窗口', width: 400, height: 140,fixed: true}); };
								</script><%end if%>
								</td>
							
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Customer_cHomepage%></td>
								<td class="td_r_l"> 
								<input name="Homepage" type="text" class="int" value="<%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cHomepage")%>"  id="Homepage" size="35" > 
								</td>
							    <td class="td_l_r title"><%=L_Customer_cFax%></td>
								<td class="td_r_l"><input name="Fax" type="text" value="<%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cFax")%>"  class="int" id="Fax" size="30"></td>
							
							</tr>

							<tr> 
								<td class="td_l_r title"><%=L_Customer_cInfo%></td>
								<td class="td_r_l" colspan=3>
								<input name="Info" type="text" class="int" value="<%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cInfo")%>"  id="Info" size="30">
								</td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Customer_cBeizhu%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;"> 
								<textarea name="Beizhu" id="Beizhu" class="int"  style="height:50px;width:98%;">
								<%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cBeizhu")%>
								</textarea>
								</td>
							</tr>
							

						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="cId" type="hidden" id="cId" value="<% = cId %>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>


<%end if%>
<%

elseif sType="SaveEdit" then
	cId = CLng(ABS(Request("cId")))
	cCompany = Trim(Request("Company"))

	cAddress = Trim(Request("Address"))
	cZip = Trim(Request("Zip"))
	cLinkman = Trim(Request("Linkman"))
	cZhiwei = Trim(Request("Zhiwei"))
	cMobile = Trim(Request("Mobile"))
	cTel = Trim(Request("Tel"))
	cFax = Trim(Request("Fax"))
	cHomepage = Trim(Request("Homepage"))
	cEmail = Trim(Request("Email"))  
	cTrade = Trim(Request("Trade"))
	if Trim(Request("Strades"))<>"" then 
		cStrade = Trim(Request("Strades"))
	else
		if Trim(Request("Strade")) <> "" then 
		cstrade = Trim(Request("Strade"))
		else
		cstrade = ""
		end if
	end if
	cType = Trim(Request("Type"))
	cStart = Trim(Request("Start"))
	cSource = Trim(Request("Source"))    
	cInfo = Trim(Request("Info"))
	cBeizhu = Trim(Request("Beizhu"))

	Set rs = Server.CreateObject("ADODB.Recordset")		
	rs.Open "Select Top 1 * From Customer Where cId = " & cId ,conn,3,2
	rs("cCompany") = cCompany
	rs("cArea") = cArea
	rs("cSquare") = cSquare
	rs("cAddress") = cAddress
	rs("cZip") = cZip
	rs("cLinkman") = cLinkman
	rs("cZhiwei") = cZhiwei
	rs("cMobile") = cMobile
	rs("cTel") = cTel
	rs("cFax") = cFax
	rs("cHomepage") = cHomepage
	rs("cEmail") = cEmail
	rs("cTrade") = cTrade
	rs("cStrade") = cStrade
	rs("cType") = cType
	rs("cStart") = cStart
	rs("cSource") = cSource
	rs("cInfo") = cInfo
	rs("cBeizhu") = cBeizhu
	rs("cLastUpdated") = now()

	rs.Update
	rs.Close
	Set rs = Nothing
	
	'插入操作记录
	'conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Customer&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	

	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="InfoView" then
	otype	=	Request.QueryString("otype")
%>
<style>body {padding-top:35px;padding-bottom:55px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Customer_cCompany%> : <%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cCompany")%> </td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			
        </td>
	</tr>
</table>


<table width="100%" border="0" cellpadding="0" cellspacing="0">
	
</table>
	<%if otype="Customer" or otype="" then%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" /><col width="300" /><col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2" style="border-right:0;"><B>基本资料</B></td>
								<td class="td_l_r" COLSPAN="2"><%=L_Customer_cLastUpdated%>：
								<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Customer","cID",""&cID&"","cLastUpdated"),1)%>
								</td>
							</tr>
							
							<tr> 
								<td class="td_l_r title"><%=L_Customer_cTel%></td>
								<td class="td_r_l"> <%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cTel")%> </td>
								<td class="td_l_r title"><%=L_Customer_cType%></td>
								<td class="td_r_l"> <%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cType")%> </td>
							
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Customer_cHomepage%></td>
								<td class="td_r_l"> <%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cHomepage")%> </td>
								<td class="td_l_r title"><%=L_Customer_cFax%></td>
								<td class="td_r_l"> <%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cFax")%> </td>
							</tr>
							
							<tr> 
								<td class="td_l_r title"><%=L_Customer_cInfo%></td>
								<td class="td_r_l" colspan=3 style="height:43px;"> <%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cInfo")%> </td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_Customer_cBeizhu%></td>
								<td class="td_r_l" colspan=3 style="height:43px;"> <%=EasyCrm.getNewItem("Customer","cID",""&cID&"","cBeizhu")%> </td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<% If mid(Session("CRM_qx"), 18, 1) = 1 Then %>
			<%if YNRange = "" then%>
			<input type="button" class="button45" value="编辑" onclick='Customer_InfoEdit();' style="cursor:pointer" />　
			<%end if%>
			<%end if%>
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>

	<%
	elseif otype="Linkmans" then '联系人 ?action=Customer&sType=InfoView&otype=Linkmans&cID=cID
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<%if Linkmans_lName = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lName%></td>
								<%end if%>
								<%if Linkmans_lSex = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lSex%></td>
								<%end if%>
								<%if Linkmans_lZhiwei = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lZhiwei%></td>
								<%end if%>
								<%if Linkmans_lMobile = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lMobile%></td>
								<%end if%>
								<%if Linkmans_lTel = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lTel%></td>
								<%end if%>
								<%if Linkmans_lEmail = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lEmail%></td>
								<%end if%>
								<%if Linkmans_lQQ = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lQQ%></td>
								<%end if%>
								<%if Linkmans_lMSN = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lMSN%></td>
								<%end if%>
								<%if Linkmans_lALWW = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lALWW%></td>
								<%end if%>
								<%if Linkmans_lBirthday = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lBirthday%></td>
								<%end if%>
								<%if Linkmans_lContent = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lContent%></td>
								<%end if%>
								<%if Linkmans_lTime = 1 then %>
								<td class="td_l_c"><%=L_Linkmans_lTime%></td>
								<%end if%>
								<%if YNRange = "" then%>
								<td width="90" class="td_l_c">管理</td>
								<%end if%>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [linkmans] where cId = "&cId&" Order By lId asc ",conn,1,1
						If rs.RecordCount > 0 Then
						i=0 
						Do While Not rs.BOF And Not rs.EOF
						i=i+1
						%>
							<tr class="tr" <%if i=1 then%>style="color:red"<%end if%>>
								<%if Linkmans_lName = 1 then %>
								<td class="td_l_c" ><%=rs("lName")%></td>
								<%end if%>
								<%if Linkmans_lSex = 1 then %>
								<td class="td_l_c"><%=rs("lSex")%></td>
								<%end if%>
								<%if Linkmans_lZhiwei = 1 then %>
								<td class="td_l_c"><%=rs("lZhiwei")%></td>
								<%end if%>
								<%if Linkmans_lMobile = 1 then %>
								<td class="td_l_c"><%=rs("lMobile")%></td>
								<%end if%>
								<%if Linkmans_lTel = 1 then %>
								<td class="td_l_c"><%=rs("lTel")%></td>
								<%end if%>
								<%if Linkmans_lEmail = 1 then %>
								<td class="td_l_c"><%=rs("lEmail")%></td>
								<%end if%>
								<%if Linkmans_lQQ = 1 then %>
								<td class="td_l_c"><%=rs("lQQ")%></td>
								<%end if%>
								<%if Linkmans_lMSN = 1 then %>
								<td class="td_l_c"><%=rs("lMSN")%></td>
								<%end if%>
								<%if Linkmans_lALWW = 1 then %>
								<td class="td_l_c"><%=rs("lALWW")%></td>
								<%end if%>
								<%if Linkmans_lBirthday = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("lBirthday"),2)%></td>
								<%end if%>
								<%if Linkmans_lContent = 1 then %>
								<td class="td_l_c"><%if rs("lContent")<>"" then%><input type="button" class="button226" value="查看" onclick='Linkmans_InfoView<%=rs("lId")%>()' style="cursor:pointer" /><%end if%></td>
								<%end if%>
								<%if Linkmans_lTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("lTime"),2)%></td>
								<%end if%>
								<%if YNRange = "" then%>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 23, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Linkmans_InfoEdit<%=rs("lId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 24, 1) = 1 Then %><%if i>1 then%><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Linkmans_InfoDel<%=rs("lId")%>()' style="cursor:pointer" /><%end if%><%end if%></td>
								<%end if%>
							</tr>
							<script>function Linkmans_InfoEdit<%=rs("lId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=Linkmans&sType=Edit&Id=<%=rs("lId")%><%if i=1 then%>&YNUpdate=1<%end if%>', {title: '编辑', width: 700,height: 340, fixed: true}); };</script>
							<script>function Linkmans_InfoDel<%=rs("lId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () { art.dialog.open('../Main/GetUpdateRW.asp?action=Linkmans&sType=Delete&Id=<%=rs("lid")%>');return false;},cancel: true }); };</script>
							<script>function Linkmans_InfoView<%=rs("lId")%>() {art.dialog({ title: '详情备注',content: '<%=rs("lContent")%>',drag: false,resize: false}); };</script>
						<%
						rs.MoveNext
						Loop
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
			<span class="Bottom_pd r fontnobold">红色：第一联系人（仅允许修改，同步更新基本档案）</span>
			<% If mid(Session("CRM_qx"), 22, 1) = 1 Then %>
			<%if YNRange = "" then%>
			<input name="Back" type="button" id="Back" class="button45" value="新增" onclick='Linkmans_InfoAdd()' style="cursor:pointer">　
			<%end if%>
			<%end if%>
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
<script>function Linkmans_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Linkmans&sType=Add&cID=<%=cID%>', {title: '新窗口', width: 700, height: 340,fixed: true}); };</script>
	<%
	elseif otype="Records" then '跟单记录 ?action=Customer&sType=InfoView&otype=Records&cID=cID
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<%if Records_rType = 1 then %>
								<td class="td_l_c"><%=L_Records_rType%></td>
								<%end if%>
								<%if Records_rState = 1 then %>
								<td class="td_l_c"><%=L_Records_rState%></td>
								<%end if%>
								<%if Records_rLinkman = 1 then %>
								<td class="td_l_c"><%=L_Records_rLinkman%></td>
								<%end if%>
								<%if Records_rNextTime = 1 then %>
								<td class="td_l_c"><%=L_Records_rNextTime%></td>
								<%end if%>
								<%if Records_rContent = 1 then %>
								<td class="td_l_c"><%=L_Records_rContent%></td>
								<%end if%>
								<%if Records_rUser = 1 then %>
								<td class="td_l_c"><%=L_Records_rUser%></td>
								<%end if%>
								<%if Records_rTime = 1 then %>
								<td class="td_l_c"><%=L_Records_rTime%></td>
								<%end if%>
								<%if YNRange = "" then%>
								<td width="90" class="td_l_c">管理</td>
								<%end if%>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Records] where cId = "&cId&" Order By rId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
							<tr class="tr">
								<%if Records_rType = 1 then %>
								<td class="td_l_c"><%=rs("rType")%></td>
								<%end if%>
								<%if Records_rState = 1 then %>
								<td class="td_l_c"><%=rs("rState")%></td>
								<%end if%>
								<%if Records_rLinkman = 1 then %>
								<td class="td_l_c"><%=rs("rLinkman")%></td>
								<%end if%>
								<%if Records_rNextTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("rNextTime"),2)%></td>
								<%end if%>
								<%if Records_rContent = 1 then %>
								<td class="td_l_l" style="line-height:25px;"><%if rs("rContent")<>"" then%><%=rs("rContent")%><%end if%></td>
								<%end if%>
								<%if Records_rUser = 1 then %>
								<td class="td_l_c"><%=rs("rUser")%></td>
								<%end if%>
								<%if Records_rTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("rTime"),2)%></td>
								<%end if%>
								<%if YNRange = "" then%>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 28, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Records_InfoEdit<%=rs("rId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 29, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Records_InfoDel<%=rs("rId")%>()' style="cursor:pointer" /><%end if%></td>
								<%end if%>
							</tr>
							<script>function Records_InfoEdit<%=rs("rId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=Records&sType=Edit&Id=<%=rs("rId")%>', {title: '编辑', width: 800,height: 340, fixed: true}); };</script>
							<script>function Records_InfoDel<%=rs("rId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () { art.dialog.open('../Main/GetUpdateRW.asp?action=Records&sType=Delete&Id=<%=rs("rId")%>');return false;},cancel: true }); };</script>
						<%
						rs.MoveNext
						Loop
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
			<% If mid(Session("CRM_qx"), 27, 1) = 1 Then %>
			<%if YNRange = "" then%>
			<input name="Back" type="button" id="Back" class="button45" value="新增" onclick='Records_InfoAdd()' style="cursor:pointer">　
			<%end if%>
			<%end if%>
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
			
		</td>
	</tr>
</table>
</div>
<script>function Records_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Records&sType=Add&cID=<%=cID%>', {title: '新窗口', width: 800, height: 340,fixed: true}); };</script>
	<%
	elseif otype="Order" then '订单记录 ?action=Customer&sType=InfoView&otype=Order&cID=cID
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<td class="td_l_c"><%=L_Order_oCode%></td>
								<%if Order_oLinkman = 1 then %>
								<td class="td_l_c"><%=L_Order_oLinkman%></td>
								<%end if%>
								<%if Order_oSDate = 1 then %>
								<td class="td_l_c"><%=L_Order_oSDate%></td>
								<%end if%>
								<%if Order_oEDate = 1 then %>
								<td class="td_l_c"><%=L_Order_oEDate%></td>
								<%end if%>
								<%if Order_oDeposit = 1 then %>
								<td class="td_l_c"><%=L_Order_oDeposit%></td>
								<%end if%>
								<td class="td_l_c"><%=L_Order_oMoney%></td>
								<%if Order_oState = 1 then %>
								<td class="td_l_c"><%=L_Order_oState%></td>
								<%end if%>
								<%if Order_oContent = 1 then %>
								<td class="td_l_c"><%=L_Order_oContent%></td>
								<%end if%>
								<%if Order_oUser = 1 then %>
								<td class="td_l_c"><%=L_Order_oUser%></td>
								<%end if%>
								<%if Order_oTime = 1 then %>
								<td class="td_l_c"><%=L_Order_oTime%></td>
								<%end if%>
								<td width="130" class="td_l_c">管理</td>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Order] where cId = "&cId&" Order By oId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
							<tr class="tr">
								<td class="td_l_c"><a title="订单产品明细"  onclick='Order_Products_List<%=rs("oId")%>()' style="cursor:pointer" ><%=rs("oCode")%></a></td>
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
								<td class="td_l_c"><%if rs("oMoney")<1 and rs("oMoney")>0 then%>0<%end if%><%=rs("oMoney")%></td>
								<%if Order_oState = 1 then %>
								<td class="td_l_c"><%if rs("oState") = 0 then%>未处理<%elseif rs("oState") = 1 then%>处理中<%elseif rs("oState") = 2 then%>已完成<%elseif rs("oState") = 3 then%>已取消<%end if%></td>
								<%end if%>
								<%if Order_oContent = 1 then %>
								<td class="td_l_c"><%if rs("oContent")<>"" then%><input type="button" class="button226" value="查看" onclick='Order_InfoView<%=rs("oId")%>()' style="cursor:pointer" /><%end if%></td>
								<%end if%>
								<%if Order_oUser = 1 then %>
								<td class="td_l_c"><%=rs("oUser")%></td>
								<%end if%>
								<%if Order_oTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("oTime"),2)%></td>
								<%end if%>
								<td class="td_l_c"><input type="button" class="button_info_add" value=" " title="快速添加产品"  onclick='Order_Products_Add<%=rs("oId")%>()' style="cursor:pointer" /> <% If mid(Session("CRM_qx"), 33, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Order_InfoEdit<%=rs("oId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 34, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Order_InfoDel<%=rs("oId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Order_Products_Add<%=rs("oId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=OrderProducts&sType=Add&Id=<%=rs("oId")%>', {title: '添加', width: 700,height: 400, fixed: true}); };</script>
							<script>function Order_Products_List<%=rs("oId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=OrderProducts&sType=List&Id=<%=rs("oId")%>', {title: '查看', width: 860,height: 440, fixed: true}); };</script>
							<script>function Order_InfoEdit<%=rs("oId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=Order&sType=Edit&Id=<%=rs("oId")%>', {title: '编辑', width: 700,height: 340, fixed: true}); };</script>
							<script>function Order_InfoDel<%=rs("oId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () { art.dialog.open('../Main/GetUpdateRW.asp?action=Order&sType=Delete&Id=<%=rs("oId")%>');return false;},cancel: true }); };</script>
							
							<script>function Order_InfoView<%=rs("oId")%>() {
								art.dialog(
									{ 
										title: '详情备注', 
										content: '<%=EasyCrm.clearWord(""&rs("oContent")&"")%>',
										drag: false,
										resize: false
									}
								); 
							};</script>
						<%
						rs.MoveNext
						Loop
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
			<% If mid(Session("CRM_qx"), 32, 1) = 1 Then %>
			<input name="Back" type="button" id="Back" class="button45" value="新增" onclick='Order_InfoAdd()' style="cursor:pointer">　
			<%end if%>
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
<script>function Order_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Order&sType=Add&cID=<%=cID%>', {title: '新窗口', width: 700, height: 340,fixed: true}); };</script>
	<%
	elseif otype="Hetong" then '合同记录 ?action=Customer&sType=InfoView&otype=Hetong&cID=cID
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<%if Hetong_hNum = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hNum%></td>
								<%end if%>
								<%if Hetong_oId = 1 then %>
								<td class="td_l_c"><%=L_Hetong_oId%></td>
								<%end if%>
								<%if Hetong_hSdate = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hSdate%></td>
								<%end if%>
								<%if Hetong_hEdate = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hEdate%></td>
								<%end if%>
								<%if Hetong_hType = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hType%></td>
								<%end if%>
								<%if Hetong_hMoney = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hMoney%></td>
								<%end if%>
								<%if Hetong_hRevenue = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hRevenue%></td>
								<%end if%>
								<%if Hetong_hOwed = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hOwed%></td>
								<%end if%>
								<%if Hetong_hInvoice = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hInvoice%></td>
								<%end if%>
								<%if Hetong_hTax = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hTax%></td>
								<%end if%>
								<td class="td_l_c"><%=L_Hetong_hState%></td>
								<%if Hetong_hContent = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hContent%></td>
								<%end if%>
								<%if Hetong_hAudit = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hAudit%></td>
								<%end if%>
								<%if Hetong_hAuditTime = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hAuditTime%></td>
								<%end if%>
								<%if Hetong_hAuditReasons = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hAuditReasons%></td>
								<%end if%>
								<%if Hetong_hUser = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hUser%></td>
								<%end if%>
								<%if Hetong_hTime = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hTime%></td>
								<%end if%>
								<td width="130" class="td_l_c">管理</td>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Hetong] where cId = "&cId&" Order By hId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
							<tr class="tr">
								<%if Hetong_hNum = 1 then %>
								<td class="td_l_c"><a title="续费记录"  onclick='Hetong_Renew_List<%=rs("hId")%>()' style="cursor:pointer" ><%=rs("hNum")%></td>
								<%end if%>
								<%if Hetong_oId = 1 then %>
								<td class="td_l_c"><%if rs("oId")<>"" then%><a title="订单产品明细"  onclick='Order_Products_List<%=rs("oId")%>()' style="cursor:pointer" ><%=EasyCrm.getNewItem("Order","oid",rs("oId"),"oCode")%></a><%end if%></td>
								<%end if%>
								<%if Hetong_hSdate = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("hSdate"),2)%></td>
								<%end if%>
								<%if Hetong_hEdate = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("hEdate"),2)%></td>
								<%end if%>
								<%if Hetong_hType = 1 then %>
								<td class="td_l_c"><%=rs("hType")%></td>
								<%end if%>
								<%if Hetong_hMoney = 1 then %>
								<td class="td_l_c"><%=rs("hMoney")%></td>
								<%end if%>
								<%if Hetong_hRevenue = 1 then %>
								<td class="td_l_c"><%=rs("hRevenue")%></td>
								<%end if%>
								<%if Hetong_hOwed = 1 then %>
								<td class="td_l_c"><%=rs("hOwed")%></td>
								<%end if%>
								<%if Hetong_hInvoice = 1 then %>
								<td class="td_l_c"><%=rs("hInvoice")%></td>
								<%end if%>
								<%if Hetong_hTax = 1 then %>
								<td class="td_l_c"><%=rs("hTax")%></td>
								<%end if%>
								<td class="td_l_c"><%=rs("hState")%></td>
								<%if Hetong_hContent = 1 then %>
								<td class="td_l_c"><%if rs("hContent")<>"" then%><input type="button" class="button226" value="查看" onclick='Hetong_InfoView<%=rs("hId")%>()' style="cursor:pointer" /><%end if%></td>
								<%end if%>
								<%if Hetong_hAudit = 1 then %>
								<td class="td_l_c"><%=rs("hAudit")%></td>
								<%end if%>
								<%if Hetong_hAuditTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("hAuditTime"),2)%></td>
								<%end if%>
								<%if Hetong_hAuditReasons = 1 then %>
								<td class="td_l_c"><%if rs("hAuditReasons")<>"" then%><input type="button" class="button226" value="查看" onclick='Hetong_AuditReasons<%=rs("hId")%>()' style="cursor:pointer" /><%end if%></td>
								<%end if%>
								<%if Hetong_hUser = 1 then %>
								<td class="td_l_c"><%=rs("hUser")%></td>
								<%end if%>
								<%if Hetong_hTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("hTime"),2)%></td>
								<%end if%>
								
								<td class="td_l_c"><input type="button" class="button_info_add" value=" " title="快速续费"  onclick='Hetong_Renew_InfoAdd<%=rs("hId")%>()' style="cursor:pointer" /> <% If mid(Session("CRM_qx"), 38, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Hetong_InfoEdit<%=rs("hId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 39, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Hetong_InfoDel<%=rs("hId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Hetong_Renew_List<%=rs("hId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=RenewList&Id=<%=rs("hId")%>', {title: '查看', width: 860,height: 440, fixed: true}); };</script>
							<script>function Order_Products_List<%=rs("oId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=OrderProducts&sType=List&Id=<%=rs("oId")%>', {title: '查看', width: 860,height: 440, fixed: true}); };</script>
							<script>function Hetong_Renew_InfoAdd<%=rs("hId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=AddRenew&Id=<%=rs("hId")%>', {title: '续费', width: 600,height: 340, fixed: true}); };</script>
							<script>function Hetong_InfoEdit<%=rs("hId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=Edit&Id=<%=rs("hId")%>', {title: '编辑', width: 700,height: 380, fixed: true}); };</script>
							<script>function Hetong_InfoDel<%=rs("hId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () { art.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=Delete&Id=<%=rs("hId")%>');return false;},cancel: true }); };</script>
							<script>function Hetong_InfoView<%=rs("hId")%>() {art.dialog({ title: '详情备注', content: '<%=EasyCrm.clearWord(""&rs("hContent")&"")%>',drag: false,resize: false}); };</script>
							<script>function Hetong_AuditReasons<%=rs("hId")%>() {art.dialog({ title: '审核原因',content: '<%=EasyCrm.clearWord(""&rs("hAuditReasons")&"")%>',drag: false,resize: false}); };</script>
						<%
						rs.MoveNext
						Loop
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
			<% If mid(Session("CRM_qx"), 37, 1) = 1 Then %>
			<input name="Back" type="button" id="Back" class="button45" value="新增" onclick='Hetong_InfoAdd()' style="cursor:pointer">　
			<%end if%>
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
<script>function Hetong_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Hetong&sType=Add&cID=<%=cID%>', {title: '新窗口', width: 700, height: 380,fixed: true}); };</script>
	<%
	elseif otype="Service" then '服务记录 ?action=Customer&sType=InfoView&otype=Service&cID=cID
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<%if Service_sSolve = 1 then %>
								<td width="60" class="td_l_c"><%=L_Service_sSolve%></td>
								<%end if%>
								<%if Service_sTitle = 1 then %>
								<td class="td_l_c"><%=L_Service_sTitle%></td>
								<%end if%>
								<%if Service_sLinkman = 1 then %>
								<td class="td_l_c"><%=L_Service_sLinkman%></td>
								<%end if%>
								<%if Service_sType = 1 then %>
								<td class="td_l_c"><%=L_Service_sType%></td>
								<%end if%>
								<%if Service_sSDate = 1 then %>
								<td class="td_l_c"><%=L_Service_sSDate%></td>
								<%end if%>
								<%if Service_sContent = 1 then %>
								<td class="td_l_c"><%=L_Service_sContent%></td>
								<%end if%>
								<%if Service_sInfo = 1 then %>
								<td class="td_l_c"><%=L_Service_sInfo%></td>
								<%end if%>
								<%if Service_sUser = 1 then %>
								<td class="td_l_c"><%=L_Service_sUser%></td>
								<%end if%>
								<%if Service_sTime = 1 then %>
								<td class="td_l_c"><%=L_Service_sTime%></td>
								<%end if%>
								<td width="90" class="td_l_c">管理</td>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Service] where cId = "&cId&" Order By sId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
							<tr class="tr">
								<%if Service_sSolve = 1 then %>
								<td class="td_l_c"><img src="<%=SiteUrl&skinurl%>images/ico/<%if rs("sSolve") = 0 then%>no<%else%>yes<%end if%>.gif" border=0></td>
								<%end if%>
								<%if Service_sTitle = 1 then %>
								<td class="td_l_c"><%=rs("sTitle")%></td>
								<%end if%>
								<%if Service_sLinkman = 1 then %>
								<td class="td_l_c"><%=rs("sLinkman")%></td>
								<%end if%>
								<%if Service_sType = 1 then %>
								<td class="td_l_c"><%=rs("sType")%></td>
								<%end if%>
								<%if Service_sSDate = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("sSDate"),2)%></td>
								<%end if%>
								<%if Service_sContent = 1 then %>
								<td class="td_l_c"><%if rs("sContent")<>"" then%><input type="button" class="button226" value="查看" onclick='Service_ContentView<%=rs("sId")%>()' style="cursor:pointer" /><%end if%></td>
								<%end if%>
								<%if Service_sInfo = 1 then %>
								<td class="td_l_c"><%if rs("sInfo")<>"" then%><input type="button" class="button226" value="查看" onclick='Service_InfoView<%=rs("sId")%>()' style="cursor:pointer" /><%end if%></td>
								<%end if%>
								<%if Service_sUser = 1 then %>
								<td class="td_l_c"><%=rs("sUser")%></td>
								<%end if%>
								<%if Service_sTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("sTime"),2)%></td>
								<%end if%>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 43, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Service_InfoEdit<%=rs("sId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 44, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Service_InfoDel<%=rs("sId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Service_InfoEdit<%=rs("sId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=Service&sType=Edit&Id=<%=rs("sId")%>', {title: '编辑', width: 800,height: 370, fixed: true}); };</script>
							<script>function Service_InfoDel<%=rs("sId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () { art.dialog.open('../Main/GetUpdateRW.asp?action=Service&sType=Delete&Id=<%=rs("sId")%>');return false;},cancel: true }); };</script>
							<script>function Service_ContentView<%=rs("sId")%>() {art.dialog({ title: '详情备注', content: '<%=EasyCrm.clearWord(""&rs("sContent")&"")%>',drag: false,resize: false}); };</script>
							<script>function Service_InfoView<%=rs("sId")%>() {art.dialog({ title: '处理结果',content: '<%=EasyCrm.clearWord(""&rs("sInfo")&"")%>',drag: false,resize: false}); };</script>
						<%
						rs.MoveNext
						Loop
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
			<% If mid(Session("CRM_qx"), 42, 1) = 1 Then %>
			<input name="Back" type="button" id="Back" class="button45" value="新增" onclick='Service_InfoAdd()' style="cursor:pointer">　
			<%end if%>
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
<script>function Service_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=Service&sType=Add&cID=<%=cID%>', {title: '新窗口', width: 800, height: 370,fixed: true}); };</script>
	<%
	elseif otype="Expense" then '费用记录 ?action=Customer&sType=InfoView&otype=Expense&cID=cID
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<%if Expense_eDate = 1 then %>
								<td class="td_l_c"><%=L_Expense_eDate%></td>
								<%end if%>
								<%if Expense_eOutIn = 1 then %>
								<td class="td_l_c"><%=L_Expense_eOutIn%></td>
								<%end if%>
								<%if Expense_eType = 1 then %>
								<td class="td_l_c"><%=L_Expense_eType%></td>
								<%end if%>
								<%if Expense_eMoney = 1 then %>
								<td class="td_l_c"><%=L_Expense_eMoney%></td>
								<%end if%>
								<%if Expense_eContent = 1 then %>
								<td class="td_l_c"><%=L_Expense_eContent%></td>
								<%end if%>
								<%if Expense_eUser = 1 then %>
								<td class="td_l_c"><%=L_Expense_eUser%></td>
								<%end if%>
								<%if Expense_eTime = 1 then %>
								<td class="td_l_c"><%=L_Expense_eTime%></td>
								<%end if%>
								<td width="90" class="td_l_c">管理</td>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Expense] where cId = "&cId&" Order By eId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
							<tr class="tr">
								<%if Expense_eDate = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("eDate"),2)%></td>
								<%end if%>
								<%if Expense_eOutIn = 1 then %>
								<td class="td_l_c"><%if rs("eOutIn") = 1 then %>收入<%else%>支出<%end if%></td>
								<%end if%>
								<%if Expense_eType = 1 then %>
								<td class="td_l_c"><%=rs("eType")%></td>
								<%end if%>
								<%if Expense_eMoney = 1 then %>
								<td class="td_l_c"><%=rs("eMoney")%></td>
								<%end if%>
								<%if Expense_eContent = 1 then %>
								<td class="td_l_c"><%if rs("eContent")<>"" then%><input type="button" class="button226" value="查看" onclick='Expense_ContentView<%=rs("eId")%>()' style="cursor:pointer" /><%end if%></td>
								<%end if%>
								<%if Expense_eUser = 1 then %>
								<td class="td_l_c"><%=rs("eUser")%></td>
								<%end if%>
								<%if Expense_eTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("eTime"),2)%></td>
								<%end if%>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 48, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Expense_InfoEdit<%=rs("eId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 49, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Expense_InfoDel<%=rs("eId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Expense_InfoEdit<%=rs("eId")%>() {$.dialog.open('../Main/GetUpdateRW.asp?action=Expense&sType=Edit&eOutIn=<%=rs("eOutIn")%>&Id=<%=rs("eId")%>', {title: '编辑', width: 500,height: 270, fixed: true}); };</script>
							<script>function Expense_InfoDel<%=rs("eId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () { art.dialog.open('../Main/GetUpdateRW.asp?action=Expense&sType=Delete&Id=<%=rs("eId")%>');return false;},cancel: true }); };</script>
							<script>function Expense_ContentView<%=rs("eId")%>() {art.dialog({ title: '详情备注', content: '<%=EasyCrm.clearWord(""&rs("eContent")&"")%>',drag: false,resize: false}); };</script>
						<%
						rs.MoveNext
						Loop
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
			<span class="Bottom_pd r">总收入：<%=EasyCrm.getSUMItem("Expense","eMoney","eMoneystr"," and cID = "&cID&" and eOutIn = 1 ")%> 元　 总支出：<%=EasyCrm.getSUMItem("Expense","eMoney","eMoneystr"," and cID = "&cID&" and eOutIn = 0 ")%> 元</span>
			<% If mid(Session("CRM_qx"), 47, 1) = 1 Then %>
			<input name="Back" type="button" id="Back" class="button45" value="新增收入" onclick='Expense_InfoAdd_IN()' style="cursor:pointer">　
			<input name="Back" type="button" id="Back" class="button46" value="新增支出" onclick='Expense_InfoAdd_OUT()' style="cursor:pointer">　
			<%end if%>
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
			
		</td>
	</tr>
</table>
</div>
<script>function Expense_InfoAdd_IN() {$.dialog.open('../Main/GetUpdateRW.asp?action=Expense&sType=Add&eOutIn=1&cID=<%=cID%>', {title: '新窗口', width: 500, height: 270,fixed: true}); };</script>
<script>function Expense_InfoAdd_OUT() {$.dialog.open('../Main/GetUpdateRW.asp?action=Expense&sType=Add&eOutIn=0&cID=<%=cID%>', {title: '新窗口', width: 500, height: 270,fixed: true}); };</script>
	<%
	elseif otype="File" then '附件记录 ?action=Customer&sType=InfoView&otype=File&cID=cID
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<td class="td_l_c" width="80"><%=L_File_fId%></td>
								<td class="td_l_l"><%=L_File_fTitle%></td>
								<td class="td_l_c" width="80">文件大小</td>
								<td class="td_l_c" width="80"><%=L_File_fFile%></td>
								<td class="td_l_c" width="80"><%=L_File_fContent%></td>
								<td class="td_l_c" width="80"><%=L_File_fUser%></td>
								<td class="td_l_c" width="80"><%=L_File_fTime%></td>
								<td width="50" class="td_l_c">管理</td>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [File] where cId = "&cId&" Order By fId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
							<tr class="tr">
								<td class="td_l_c"><%=rs("fId")%></td>
								<td class="td_l_l"><%=rs("fTitle")%></td>
								<td class="td_l_c"><%if rs("fFile")<>"" then%><%=EasyCrm.showsize(rs("fFile"))%><%end if%></td>
								<td class="td_l_c"><%if rs("fFile")<>"" then%><input type="button" class="button222" <%if inStr("'gif','jpg','png','bmp'", right(rs("fFile"),3) ) > 0 then %>value="查看"  onclick="javascript:window.open('<%=rs("fFile")%>')" <%else%>value="下载" onClick=window.location.href="<%=rs("fFile")%>"<%end if%> style="cursor:pointer" /><%else%>无<%end if%></td>
								<td class="td_l_c"><%if rs("fContent")<>"" then%><input type="button" class="button226" value="查看" onclick='File_ContentView<%=rs("fId")%>()' style="cursor:pointer" /><%else%>无<%end if%></td>
								<td class="td_l_c"><%=rs("fUser")%></td>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("fTime"),2)%></td>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 54, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='File_InfoDel<%=rs("fId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function File_InfoDel<%=rs("fId")%>() {$.dialog( { content: '<%=Alert_del_YN%>',icon: 'error',ok: function () { art.dialog.open('../Main/GetUpdateRW.asp?action=File&sType=Delete&Id=<%=rs("fId")%>');return false;},cancel: true }); };</script>
							<script>function File_ContentView<%=rs("fId")%>() {art.dialog({ title: '详情备注', content: '<%=EasyCrm.clearWord(""&rs("fContent")&"")%>',drag: false,resize: false}); };</script>
							<script>function File_InfoView<%=rs("fId")%>() {art.dialog({ title: '查看图片', content: '<img src="<%=rs("fFile")%>" />',lock: true}); };</script>
						<%
						rs.MoveNext
						Loop
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
			<% If mid(Session("CRM_qx"), 52, 1) = 1 Then %>
			<input name="Back" type="button" id="Back" class="button45" value="新增" onclick='File_InfoAdd()' style="cursor:pointer">　
			<%end if%>
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
			
		</td>
	</tr>
</table>
</div>
<script>function File_InfoAdd() {$.dialog.open('../Main/GetUpdateRW.asp?action=File&sType=Add&cID=<%=cID%>', {title: '新窗口', width: 500, height: 210,fixed: true}); };</script>
	<%
	elseif otype="Share" then '共享记录
	%>
	<script>
	function Setdisabled(evt)
	{
		var evt=evt || window.event;   
		var e =evt.srcElement || evt.target;
		
		 if(e.value=="1")
		 {
			var a = document.all.cShareRange; 
			for (var i=0; i<a.length; i++)   
			{ 
				a[i].disabled=false; 
				a[i].readOnly=false; 
			} 
		 }
		 else
		 {
			var a = document.all.cShareRange; 
			for (var i=0; i<a.length; i++)   
			{ 
				a[i].disabled=true; 
				a[i].readOnly=true; 
			} 
		 }
	}
	</script>
		<form name="Save" action="?action=Customer&sType=InfoView&otype=ShareSave&cID=<%=cId%>" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2">选择共享对象</td>
							</tr>
							<tr class="tr_f"> 
								<td class="td_l_c red">是否共享 </td>
								<td class="td_l_l red"><input type="radio" id="cShare" name="cShare" value= '0' <%if EasyCrm.getNewItem("Customer","cID",""&cID&"","cShare")=0 then %>checked <%end if%> onclick="Setdisabled()"> 否　<input type="radio" id="cShare" name="cShare" value= '1' <%if EasyCrm.getNewItem("Customer","cID",""&cID&"","cShare")=1 then %>checked <%end if%> onclick="Setdisabled()"> 是</td>
							</tr>
							<%
								Set rsg = Server.CreateObject("ADODB.Recordset")
								rsg.Open "Select * From [system_group]",conn,1,1
								Do While Not rsg.BOF And Not rsg.EOF
							%>
							<tr> 
								<td class="td_l_c title"><%=rsg("gName")%></td>
								<td  class="td_l_l">
								<%
									Set rsm = Server.CreateObject("ADODB.Recordset")
									rsm.Open "Select * From [user] where uGroup="&rsg("gId")&" ",conn,1,1
									Do While Not rsm.BOF And Not rsm.EOF
								%>
									<input type="checkbox" id="cShareRange" name="cShareRange" value= '<%=rsm("uName")%>' <%if EasyCrm.getNewItem("Customer","cID",""&cID&"","cShare")=0 then %>disabled readOnly <%end if%> <%if inStr(EasyCrm.getNewItem("Customer","cID",""&cID&"","cShareRange"),rsm("uName"))>0 then%>checked<%end if%>> <%=rsm("uName")%>　
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
						<input name="Back" type="Submit" id="Back" class="button45" value="保存" onclick='Share_InfoSave()' style="cursor:pointer">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
						
					</td>
				</tr>
			</table>
			</div>
		</form>

	<%
	elseif otype="ShareSave" then '保存共享
		cShare = Request.Form("cShare")
		cShareRange = Request.Form("cShareRange")
		conn.execute("update [Customer] set cShare='"&cShare&"',cShareRange='"&cShareRange&"' where cId = "&cID&" ")
		Response.Write("<script>location.href='?action=Customer&sType=InfoView&otype=Share&cID="&cId&"&tipinfo=操作成功！';</script>")
		Response.End()
	%>
	
	<%
	elseif otype="History" then '历史记录 ?action=Customer&sType=InfoView&otype=File&cID=cID
	%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 

						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t">
								<td class="td_l_c" width="80">编号</td>
								<td class="td_l_c" width="80">数据表</td>
								<td class="td_l_c" width="80">行为</td>
								<td class="td_l_l">原因</td>
								<td class="td_l_c" width="80">操作人</td>
								<td class="td_l_c" width="130">时间</td>
							</tr>
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Logfile] where lCid = "&cId&" Order By lId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
							<tr class="tr">
								<td class="td_l_c"><%=rs("lId")%></td>
								<td class="td_l_c"><%=rs("lClass")%></td>
								<td class="td_l_c"><%=rs("lAction")%></td>
								<td class="td_l_l"><%=rs("lReason")%></td>
								<td class="td_l_c"><%=rs("lUser")%></td>
								<td class="td_l_c"><%=rs("lTime")%></td>
							</tr>
						<%
						rs.MoveNext
						Loop
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
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
			
		</td>
	</tr>
</table>
</div>
	<%
	end if
	%>
<%
elseif sType="DelReason" then '删除客户填写操作原因
%>	<script language="JavaScript">
	<!-- 跟单记录必填项提示
	function CheckInput()
	{
		if(document.all.lReason.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=alert04%>'});document.all.lReason.focus();return false;}
	}
	-->
	</script>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				
				<form name="Save" action="?action=Customer&sType=DelTrue" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<textarea name="lReason" rows="4" id="lReason" class="int" style="height:80px;width:98%;"></textarea>
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="cId" type="hidden" id="cId" value="<% = cId %>">
			<input type="submit" name="Submit" class="button45" value="保存" >　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
<%
elseif sType="DelTrue" then '执行删除客户操作
	cId = CLng(ABS(Request("cId")))
	lReason = Trim(Request("lReason"))
	conn.execute("update Customer set cYn = 0 where cId = "&cId&" ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lReason,lUser,lTime) values('"&cid&"','"&L_Customer&"','"&L_insert_action_03&"','"&lReason&"','"&Session("CRM_name")&"','"&now()&"')")

		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
%>

<%
end if
%>

<%
End Sub
%>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../data/calendar/WdatePicker.js"></script>
</body>
<% Set EasyCrm = nothing %>