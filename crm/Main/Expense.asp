<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 46, 1) = 1 Then %>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Expense.asp"
Session("CRM_pagenum") = PNN

If subAction = "searchItem" Then
    Dim Company,eOutIn,eType,eUser,TimeBegin,TimeEnd
	Company = EasyCrm.Searchcode(Request.Form("Company"))
	eOutIn = EasyCrm.Searchcode(Request.Form("eOutIn"))
	eTypeA = EasyCrm.Searchcode(Request.Form("eTypeA"))
	eTypeB = EasyCrm.Searchcode(Request.Form("eTypeB"))
	IF eTypeA<>"" THEN
	eType = eTypeA
	ELSE
	eType = eTypeB
	END IF
	eUser = EasyCrm.Searchcode(Request.Form("User"))
	TimeBegin = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	TimeEnd = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	
	Session("Search_Expense_Company") = EasyCrm.Searchcode(Request.Form("Company"))
	Session("Search_Expense_eOutIn") = EasyCrm.Searchcode(Request.Form("eOutIn"))
	IF eTypeA<>"" THEN
	Session("Search_Expense_eTypeA") = EasyCrm.Searchcode(Request.Form("eTypeA"))
	Session("Search_Expense_eTypeB") = ""
	ELSE
	Session("Search_Expense_eTypeA") = ""
	Session("Search_Expense_eTypeB") = EasyCrm.Searchcode(Request.Form("eTypeB"))
	END IF
	Session("Search_Expense_eUser") = EasyCrm.Searchcode(Request.Form("User"))
	Session("Search_Expense_TimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Expense_TimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	
	Dim sql
    sql = ""
	
	CompanyWhere = EasyCrm.seachKey("cCompany",Company)
	
    If Company <> "" Then
        sql = sql & " And cId in ( select cId from [Client] where 1=1 "&CompanyWhere&" )"
	End If
	
    If eOutIn <> "" Then
	    sql = sql & " And eOutIn = '" & eOutIn & "'"
	End If
	
    If eType <> "" Then
	    sql = sql & " And eType = '" & eType & "'"
	End If
	
    If eUser <> "" Then
	    sql = sql & " And eUser = '" & eUser & "'"
	End If
	
	if Accsql =1 then
	If TimeBegin <> "" and  TimeEnd <> "" Then
	    sql = sql & " And eDate >= '" & TimeBegin & "' And eDate <= '" & TimeEnd & "' "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And eDate = '" & TimeBegin & "' "
	End If
	else
	If TimeBegin <> "" and TimeEnd <> "" Then
	    sql = sql & " And eDate >= #" & TimeBegin & "# And eDate <= #" & TimeEnd & "# "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And eDate = #" & TimeBegin & "# "
	End If
	end if
	
	If Session("CRM_level") < 9 Then
		sql = sql & " And eUser In (" & arrUser & ")"
	End If
	
End If

If Company = "" And eOutIn = "" And eType = "" And eUser = "" And TimeBegin = "" And TimeEnd = "" Then
    If Session("CRM_Expense_Search") <> "" Then
        sql = Session("CRM_Expense_Search")
	Else
	    If Session("CRM_level") < 9 Then
			sql = " And eUser In (" & arrUser & ")"
		End If
	End If
Else
    Session("CRM_Expense_Search") = sql
End If

If subAction = "killSession" Then
	Session("CRM_Expense_Search") = ""
	Session("Search_Expense_Company") = ""
	Session("Search_Expense_eOutIn") = ""
	Session("Search_Expense_eType") = ""
	Session("Search_Expense_eUser") = ""
	Session("Search_Expense_TimeBegin") = ""
	Session("Search_Expense_TimeEnd") = ""
	If Session("CRM_level") < 9 Then
		sql = " And eUser In (" & arrUser & ")"
	else
		sql=""
	end if
End If


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<link href="<%=SiteUrl&skinurl%>chosen/chosen.css" rel="stylesheet" />
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Float.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>

<body oncontextmenu=self.event.returnValue=false> 
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Company%> > <%=L_Expense%> <%=sql%></td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
			<% If mid(Session("CRM_qx"), 5, 1) = 1 Then %>
			<input type="button" class="button_top_set" value=" " title="设置" onclick='Setting_Expense()' style="cursor:pointer" />
			<%end if%>
        </td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdt10">   
			<div class="MenuboxS">
				 <ul>
					<li <%if otype="Main" or otype="" then%>class="hover"<%end if%> id="CheckB"><span><a href="?otype=Main">信息列表</a></span></li>
					<li class="" id="CheckA"><span><a href="javascript:voId(0)" style="cursor:pointer">高级搜索</a></span></li>
				</ul>
			</div>
		</td>
	</tr>
</table>
<script>function Setting_Expense() {$.dialog.open('../system/GetUpdate.asp?action=Setting&sType=Expense', {title: '自定义设置', width: 900, height: 480,fixed: true}); };</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr> 
		<td valign="top" class="td_n">
	<script>
	function Show()
	{
		if (document.getElementById('eOutIn').value=="1") 
		 {
			document.getElementById("eTypeA").style.display = "block";
			document.getElementById("eTypeB").style.display = "none";
		 }
		else if (document.getElementById('eOutIn').value=="0") 
		 {
			document.getElementById("eTypeA").style.display = "none";
			document.getElementById("eTypeB").style.display = "block";
		 }
		 else if (document.getElementById('eOutIn').value=="") 
		 {
			document.getElementById("eTypeA").style.display = "none";
			document.getElementById("eTypeB").style.display = "none";
		 }
	}
	</script>
			<div id="SearchBox" style="position: absolute; width:100%; height:450px; background:#ffffff; display:none; z-index:10;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0" style="background:#ffffff;">
					<tr>
						<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
							<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
								<tr class="tr_t"> 
									<td class="td_l_l" COLSPAN="2" style="border-right:0;"><B><%=L_Top_Search%></B></td>
								</tr>
							</table>
							<form name="searchForm" action="?subAction=searchItem" method="post">
							<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
								<col width="120" /><col width="80" />
								<tr>
									<td class="td_l_r title" style="border-top:0;"><%=L_Expense_cID%></td>
									<td class="td_r_l" colspan=2 style="border-top:0;"><input name="Company" type="text" class="int" id="Company" size="30" value="<%=Session("Search_Expense_Company")%>" ></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Expense_eOutIn%></td>
									<td class="td_r_l" style="border-right:0;"><select name='eOutIn' onchange="Show();"><option value="">请选择</option><option value="0" <%IF Session("Search_Expense_eOutIn")="0" THEN%>SELECTED <%END IF%>>支出</option><option value="1" <%IF Session("Search_Expense_eOutIn")="1" THEN%>SELECTED <%END IF%>>收入</option></select> 
									</td>
									<td class="td_r_l">
									<span id=eTypeA<%IF Session("Search_Expense_eTypeA")="" THEN%> STYLE="display:none;" <%END IF%>><% = EasyCrm.getSelect("SelectData","Select_ExpenseIN","eTypeA","") %></span>
									<span id=eTypeB<%IF Session("Search_Expense_eTypeB")="" THEN%> STYLE="display:none;" <%END IF%>><% = EasyCrm.getSelect("SelectData","Select_ExpenseOUT","eTypeB","") %></span>
									</td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Expense_eDate%></td>
									<td class="td_r_l" colspan=2><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Expense_TimeBegin")%>" /> ~ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Expense_TimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Expense_eUser%></td>
									<td class="td_r_l" colspan=2>
										<% If Session("CRM_level") = 9 Then %>
										<% = EasyCrm.UserList(2,"User",Session("Search_Expense_eUser")) %>
										<% Else %>
										<% = EasyCrm.UserList(1,"User",Session("Search_Expense_eUser")) %>
										<% End If %>
									</td>
								</tr>
								<tr>
									<td class="td_r_l" colspan="3">
										<input type="submit" name="Submit" class="button42" value="<%=L_Search%>">
										<input type="button" name="button" class="button43" value="<%=L_Clear%>" onClick=window.location.href="?SubAction=killSession" />
									</td>
								</tr>
							</table>   
							</form>
						</td> 
					</tr>
				</table>
			</div>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" colspan=2 style="padding:10px 10px 0;" class="td_n">
						<form name="Search" action="?action=CheckSub&SubAction=Search&PN=<%=PNN%>" method="post">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1" style="border-bottom:0px;">
							<tr class="tr_t"style="border-bottom:0px;"> 
								<td class="td_l_l"><B>信息列表</B></td>
							</tr>
						</table> 
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_b">
								<td width="80" class="td_l_c"><%=L_Expense_eId%></td>
								<td class="td_l_l"><%=L_Expense_cId%></td>
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
								<td class="td_l_l"><%=L_Expense_eContent%></td>
								<%end if%>
								<%if Expense_eUser = 1 then %>
								<td class="td_l_c"><%=L_Expense_eUser%></td>
								<%end if%>
								<%if Expense_eTime = 1 then %>
								<td class="td_l_c"><%=L_Expense_eTime%></td>
								<%end if%>
								<td width="100" class="td_l_c"><%=L_Top_Manage%></td>
							</tr>
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [Expense] where cID<>'' "&sql&" Order By eId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Expense] where cID<>'' "&sql&" and eId < ( SELECT Min(eId) FROM ( SELECT TOP "&pagenum&" eId FROM [Expense]  where cID<>'' "&sql&" Order BY eId desc ) AS T ) Order By eId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(eId) As Expensenum From [Expense] where cID<>'' "&sql&" ",1,1)
							
							TotalExpense=Rsstr("Expensenum") 
							if Int(TotalExpense/intPageSize)=TotalExpense/intPageSize then
							TotalPages=TotalExpense/intPageSize
							else
							TotalPages=Int(TotalExpense/intPageSize)+1
							end if
							Rsstr.Close
							Set Rsstr=Nothing
							If PN > TotalPages Then PN = TotalPages
							If rs.RecordCount > 0 Then
							Do While Not rs.BOF And Not rs.EOF
							%>
							<tr class="tr">
								<td class="td_l_c"><%=rs("eId")%></td>
								<td class="td_l_l"><a onclick='Client_InfoView<%=rs("cId")%>()' style="cursor:pointer"><%=EasyCrm.getNewItem("Client","cID",rs("cId"),"cCompany")%></a></td>
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
								<td class="td_l_l" style="line-height:25px;"><%if rs("eContent")<>"" then%><%=rs("eContent")%><%end if%></td>
								<%end if%>
								<%if Expense_eUser = 1 then %>
								<td class="td_l_c"><%=rs("eUser")%></td>
								<%end if%>
								<%if Expense_eTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("eTime"),2)%></td>
								<%end if%>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 48, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Expense_InfoEdit<%=rs("eId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 49, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Expense_InfoDel<%=rs("eId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							
							<script>function Client_InfoView<%=rs("cId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&otype=Expense&cId=<%=rs("cId")%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
							
							<script>function Expense_InfoEdit<%=rs("eId")%>() {$.dialog.open('GetUpdateRW.asp?action=Expense&sType=Edit&eOutIn=<%=rs("eOutIn")%>&Id=<%=rs("eId")%>', {title: '编辑', width: 500,height: 270, fixed: true}); };</script>
							
							<script>function Expense_InfoDel<%=rs("eId")%>(){art.dialog({content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {
								<%if YnDelReason = 1 then%> 
								$.dialog.open('GetUpdateRW.asp?action=Expense&sType=DelReason&Id=<%=rs("eId")%>',{title: '删除原因', width: 400,height: 150, fixed: true}); 
								<%else%>
								art.dialog.open('?action=delete&Id=<%=rs("eId")%>');
								<%end if%>
								art.dialog.close();},cancelVal: '关闭',cancel: true});};
							</script>
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
        </td>
	</tr>
	</form>
</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd ">
			<%if sql<>"" then%><span class="r"><input name="Back" type="button" id="Back" class="button227" value="清空" onClick=window.location.href="?SubAction=killSession"></span><%end if%>
			<%=EasyCrm.pagelist("Expense.asp", PN,TotalPages,TotalExpense)%>
		</td> 
	</tr>
</table>
</div>
<%
Select Case action
Case "delete"
    Call deleteData()
End Select

Sub deleteData()
	id = Trim(Request("id"))
	If id = "" Then
	Exit Sub
	End If
	cID = EasyCrm.getNewItem("Expense","eId",""&id&"","cID")
	conn.execute("DELETE FROM [Expense] where eId = "&Id&" ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cID&"','"&L_Expense&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
<script src="../data/calendar/WdatePicker.js"></script>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html><%else%>无权限<%end if%>
<% Set EasyCrm = nothing %>