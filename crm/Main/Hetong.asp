<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 36, 1) = 1 Then %>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Hetong.asp"
Session("CRM_pagenum") = PNN

If subAction = "searchItem" Then
    Dim Company,hNum,hType,hState,rState,hOwed,User,TimeBegin,TimeEnd,ETimeBegin,ETimeEnd
	Company = EasyCrm.Searchcode(Request.Form("company"))
	hNum = EasyCrm.Searchcode(Request.Form("hNum"))
	hType = EasyCrm.Searchcode(Request.Form("hType"))
	hState = EasyCrm.Searchcode(Request.Form("hState"))
	rState = EasyCrm.Searchcode(Request.Form("rState"))
	hOwed = EasyCrm.Searchcode(Request.Form("hOwed"))
	hUser = EasyCrm.Searchcode(Request.Form("User"))
	TimeBegin = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	TimeEnd = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	ETimeBegin = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	ETimeEnd = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	Session("Search_Hetong_Company") = EasyCrm.Searchcode(Request.Form("Company"))
	Session("Search_Hetong_hNum") = EasyCrm.Searchcode(Request.Form("hNum"))
	Session("Search_Hetong_hType") = EasyCrm.Searchcode(Request.Form("hType"))
	Session("Search_Hetong_hState") = EasyCrm.Searchcode(Request.Form("hState"))
	Session("Search_Hetong_rState") = EasyCrm.Searchcode(Request.Form("rState"))
	Session("Search_Hetong_hOwed") = EasyCrm.Searchcode(Request.Form("hOwed"))
	Session("Search_Hetong_hUser") = EasyCrm.Searchcode(Request.Form("User"))
	Session("Search_Hetong_TimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Hetong_TimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	Session("Search_Hetong_ETimeBegin") = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	Session("Search_Hetong_ETimeEnd") = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	Dim sql
    sql = ""
	
	CompanyWhere = EasyCrm.seachKey("cCompany",Company)
	
    If Company <> "" Then
        sql = sql & " And cId in ( select cId from [Client] where 1=1 "&CompanyWhere&" )"
	End If
	
    If hNum <> "" Then
        sql = sql & " And hNum like '%" & hNum & "%' "
	End If
	
    If hType <> "" Then
	    sql = sql & " And hType = '" & hType & "'"
	End If
	
    If hState <> "" Then
	    sql = sql & " And hState = '" & hState & "'"
	End If
	
    If rState <> "" Then
	    sql = sql & " And hID in ( select hid from Hetong_Renew where rState = '"&rState&"' ) "
	End If
	
    If hOwed <> "" Then
		If hOwed = 0 then 
		sql = sql & " And hOwed = 0 "
		elseif hOwed = 1 then
		sql = sql & " And hOwed > 0 "
		end if
	End If
	
    If hUser <> "" Then
	    sql = sql & " And hUser = '" & hUser & "'"
	End If
	
	if Accsql =1 then
	If TimeBegin <> "" and  TimeEnd <> "" Then
	    sql = sql & " And hSdate >= '" & TimeBegin & "' And hSdate <= '" & TimeEnd & "' "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And hSdate = '" & TimeBegin & "' "
	End If
	else
	If TimeBegin <> "" and TimeEnd <> "" Then
	    sql = sql & " And hSdate >= #" & TimeBegin & "# And hSdate <= #" & TimeEnd & "# "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And hSdate = #" & TimeBegin & "# "
	End If
	end if
	
	if Accsql =1 then
	If ETimeBegin <> "" and  ETimeEnd <> "" Then
	    sql = sql & " And hEdate >= '" & ETimeBegin & "' And hEdate <= '" & ETimeEnd & "' "
	End If
	If ETimeBegin <> "" and  ETimeEnd = "" Then
	    sql = sql & " And hEdate = '" & ETimeBegin & "' "
	End If
	else
	If ETimeBegin <> "" and ETimeEnd <> "" Then
	    sql = sql & " And hEdate >= #" & ETimeBegin & "# And hEdate <= #" & ETimeEnd & "# "
	End If
	If ETimeBegin <> "" and  ETimeEnd = "" Then
	    sql = sql & " And hEdate = #" & ETimeBegin & "# "
	End If
	end if
		
	
	If Session("CRM_level") < 9 Then
		sql = sql & " And hUser In (" & arrUser & ")"
	End If
	
End If

If Company = "" And hNum = "" And hType = "" And hState = "" And hUser = "" And TimeBegin = "" And TimeEnd = ""  And ETimeBegin = "" And ETimeEnd = "" Then
    If Session("CRM_Hetong_Search") <> "" Then
        sql = Session("CRM_Hetong_Search")
	Else
	    If Session("CRM_level") < 9 Then
			sql = " And hUser In (" & arrUser & ")"
		End If
	End If
Else
    Session("CRM_Hetong_Search") = sql
End If

If subAction = "killSession" Then
	Session("CRM_Hetong_Search") = ""
	Session("Search_Hetong_Company") = ""
	Session("Search_Hetong_hNum") = ""
	Session("Search_Hetong_hType") = ""
	Session("Search_Hetong_hState") = ""
	Session("Search_Hetong_hUser") = ""
	Session("Search_Hetong_TimeBegin") = ""
	Session("Search_Hetong_TimeEnd") = ""
	Session("Search_Hetong_ETimeBegin") = ""
	Session("Search_Hetong_ETimeEnd") = ""
	If Session("CRM_level") < 9 Then
		sql = " And hUser In (" & arrUser & ")"
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

<body> 
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Company%> > <%=L_Hetong%> </td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
			<% If mid(Session("CRM_qx"), 5, 1) = 1 Then %>
			<input type="button" class="button_top_set" value=" " title="设置" onclick='Setting_Hetong()' style="cursor:pointer" />
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
<script>function Setting_Hetong() {$.dialog.open('../system/GetUpdate.asp?action=Setting&sType=Hetong', {title: '自定义设置', width: 900, height: 480,fixed: true}); };</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr> 
		<td valign="top" class="td_n">
		
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
								<col width="120" />
								<tr>
									<td class="td_l_r title" style="border-top:0;"><%=L_Hetong_cID%></td>
									<td class="td_r_l" style="border-top:0;"><input name="Company" type="text" class="int" id="Company" size="30" value="<%=Session("Search_Hetong_Company")%>" ></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Hetong_hNum%></td>
									<td class="td_r_l"><input name="hNum" type="text" class="int" id="hNum" size="30" value="<%=Session("Search_Hetong_hNum")%>"></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Hetong_hSdate%></td>
									<td class="td_r_l"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Hetong_TimeBegin")%>" /> ~ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Hetong_TimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Hetong_hEdate%></td>
									<td class="td_r_l"><input name="ETimeBegin" type="text" maxlength="10" id="ETimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Hetong_ETimeBegin")%>" /> ~ <input name="ETimeEnd" type="text" maxlength="10" id="ETimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Hetong_ETimeEnd")%>" /></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Hetong_hType%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Hetong","hType",Session("Search_Hetong_hType")) %></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Hetong_hState%></td>
									<td class="td_r_l"><select name='hState'><option value="">请选择</option><option value="<%=L_Hetong_hState_1%>"><%=L_Hetong_hState_1%></option><option value="<%=L_Hetong_hState_2%>"><%=L_Hetong_hState_2%></option><option value="<%=L_Hetong_hState_3%>"><%=L_Hetong_hState_3%></option></select></td>
								</tr>
								<tr>
									<td class="td_l_r title">是否欠款</td>
									<td class="td_r_l"><input type="radio" name="hOwed" value="" <%if Session("Search_Hetong_hOwed") = "" then %>checked<%end if%>> 未知　<input type="radio" name="hOwed" value="1" <%if Session("Search_Hetong_hOwed") = "1" then %>checked<%end if%>> 有　<input type="radio" name="hOwed" value="0" <%if Session("Search_Hetong_hOwed") = "0" then %>checked<%end if%> > 无 </td>
								</tr>
								<tr>
									<td class="td_l_r title">续费状态</td>
									<td class="td_r_l"><select name='rState'><option value="">请选择</option><option value="续费有效">续费有效</option><option value="续费无效">续费无效</option><option value="待审">待审</option></select></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Hetong_hUser%></td>
									<td class="td_r_l">
										<% If Session("CRM_level") = 9 Then %>
										<% = EasyCrm.UserList(2,"User",Session("Search_Hetong_hUser")) %>
										<% Else %>
										<% = EasyCrm.UserList(1,"User",Session("Search_Hetong_hUser")) %>
										<% End If %>
									</td>
								</tr>
								<tr>
									<td class="td_r_l" colspan="2">
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
			
<script language="JavaScript">
<!--
for(var i=0;i<document.getElementById('hState').options.length;i++){
	if(document.getElementById('hState').options[i].value == "<% = Session("Search_Hetong_hState") %>"){
	document.getElementById('hState').options[i].selected = true;}}
for(var i=0;i<document.getElementById('rState').options.length;i++){
	if(document.getElementById('rState').options[i].value == "<% = Session("Search_Hetong_rState") %>"){
	document.getElementById('rState').options[i].selected = true;}}
-->
</script>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" colspan=2 style="padding:10px 10px 0;" class="td_n">
						<form name="Search" action="?action=CheckSub&SubAction=Search&PN=<%=PNN%>" method="post">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1" style="border-bottom:0px;">
							<tr class="tr_t"style="border-bottom:0px;"> 
								<td class="td_l_l" style="border-right:0;"><B>信息列表</B></td>
							</tr>
						</table> 
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_b">
								<td width="80" class="td_l_c"><%=L_Hetong_hId%></td>
								<td class="td_l_l"><%=L_Hetong_cId%></td>
								<%if Hetong_hNum = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hNum%></td>
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
								<%if Hetong_hAudit = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hAudit%></td>
								<%end if%>
								<%if Hetong_hAuditTime = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hAuditTime%></td>
								<%end if%>

  <%							
Set rss = Server.CreateObject("ADODB.Recordset")
rss.Open "Select * From [CustomField] where cTable = 'Hetong' and cList = '1' order by Id asc ",conn,3,1
If rss.RecordCount > 0 Then
Do While Not rss.BOF And Not rss.EOF
%>
<td class="td_l_c"><%=rss("cTitle")%></td>
<%
rss.MoveNext
Loop
end if
rss.Close
Set rss = Nothing
%>
                                
















							<%If mid(Session("CRM_qx"), 14, 1) = "1" Then%>
								<td class="td_l_c">审核</td>
							<%end if%>
								<%if Hetong_hUser = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hUser%></td>
								<%end if%>
								<%if Hetong_hTime = 1 then %>
								<td class="td_l_c"><%=L_Hetong_hTime%></td>
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
							rs.Open "Select top "&intPageSize&" * From [Hetong] where cID<>'' "&sql&" Order By hId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Hetong] where cID<>'' "&sql&" and hId < ( SELECT Min(hId) FROM ( SELECT TOP "&pagenum&" hId FROM [Hetong]  where cID<>'' "&sql&" Order BY hId desc ) AS T ) Order By hId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(hId) As Hetongum From [Hetong] where cID<>'' "&sql&" ",1,1)
							
							TotalHetong=Rsstr("Hetongum") 
							if Int(TotalHetong/intPageSize)=TotalHetong/intPageSize then
							TotalPages=TotalHetong/intPageSize
							else
							TotalPages=Int(TotalHetong/intPageSize)+1
							end if
							Rsstr.Close
							Set Rsstr=Nothing
							If PN > TotalPages Then PN = TotalPages
							If rs.RecordCount > 0 Then
							Do While Not rs.BOF And Not rs.EOF
							%>
							<tr class="tr">
								<td class="td_l_c"><%=rs("hId")%></td>
								<td class="td_l_l"><a onclick='Hetong_InfoView<%=rs("cId")%>()' style="cursor:pointer"><%=EasyCrm.getNewItem("Client","cID",rs("cId"),"cCompany")%></a></td>
								<%if Hetong_hNum = 1 then %>
								<td class="td_l_c"><a <%if EasyCrm.getCountItem("Hetong_Renew","rId","IDstr"," and hID = "&rs("hId")&" and rState = '待审' ") > 0 then %> title="续费待审" style="color:red;cursor:pointer"<%else%>title="续费记录" style="cursor:pointer"<%end if%>  onclick='Hetong_Renew_List<%=rs("hId")%>()'  ><%=rs("hNum")%></td>
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
								<%if Hetong_hAudit = 1 then %>
								<td class="td_l_c"><%=rs("hAudit")%></td>
								<%end if%>
								<%if Hetong_hAuditTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("hAuditTime"),2)%></td>
								<%end if%>

                                
  <%                            
                               	cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID",""&rs("cId")&" and hid = "&rs("hId")&" ","cContent")
								cContentArr = split(cContentStr,"|")	
															
								Set rss1 = Server.CreateObject("ADODB.Recordset")
								rss1.Open "Select * From [CustomField] where cTable = 'Hetong' and cList = '1' order by id asc ",conn,1,1
								If rss1.RecordCount > 0 Then
								Do While Not rss1.BOF And Not rss1.EOF

cname = rss1("cname")
								
For x=LBound(cContentArr) to UBound(cContentArr)-1

cZiduan=split(cContentArr(x),":")

if inStr(cContentArr(x),cZiduan(0))>0 then

  if cZiduan(0) = cname then
  y = x
  end if										
										
end if

Next
'Response.Write i '--词语数量			
								
								k=y
								if Ubound(cContentArr) > k then
								cContent = split(cContentArr(k),":")
								%>

                                   <td class="td_l_c">
									<%if inStr(cContentArr(k),cContent(0))>0 then%>
										<%=cContent(1)%>
									<%end if%>
									</td>
								<%
								else
								%>
								    <td class="td_l_c">
									</td>

								<%
								end if
								k=k+1
								rss1.MoveNext
								Loop
								end if
								rss1.Close
								Set rss1 = Nothing
							    %>
























							<%If mid(Session("CRM_qx"), 14, 1) = "1" Then%>
								<td class="td_l_c">
									<input type="button" class="button222" value="审核"  onclick='Hetong_InfoAudit<%=rs("hId")%>()' style="cursor:pointer" />
								</td>
							<%end if%>
								<%if Hetong_hUser = 1 then %>
								<td class="td_l_c"><%=rs("hUser")%></td>
								<%end if%>
								<%if Hetong_hTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("hTime"),2)%></td>
								<%end if%>
								
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 38, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Hetong_InfoEdit<%=rs("hId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 39, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Hetong_InfoDel<%=rs("hId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Hetong_InfoView<%=rs("cId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&otype=Hetong&cId=<%=rs("cId")%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
							<script>function Hetong_Renew_List<%=rs("hId")%>() {$.dialog.open('GetUpdateRW.asp?action=Hetong&sType=RenewList&Id=<%=rs("hId")%>', {title: '查看', width: 860,height: 440, fixed: true}); };</script>
							<script>function Order_Products_List<%=rs("oId")%>() {$.dialog.open('GetUpdateRW.asp?action=OrderProducts&sType=List&Id=<%=rs("oId")%>', {title: '查看', width: 860,height: 440, fixed: true}); };</script>
							<script>function Hetong_InfoEdit<%=rs("hId")%>() {$.dialog.open('GetUpdateRW.asp?action=Hetong&sType=Edit&Id=<%=rs("hId")%>', {title: '编辑', width: 600,height: 380, fixed: true}); };</script>
							
							<script>function Hetong_InfoAudit<%=rs("hId")%>() {$.dialog.open('GetUpdateRW.asp?action=Hetong&sType=Audit&Id=<%=rs("hId")%>', {title: '审核', width: 400,height: 180, fixed: true}); };</script>
							
							<script>function Hetong_InfoDel<%=rs("hId")%>(){art.dialog({content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {
								<%if YnDelReason = 1 then%> 
								$.dialog.open('GetUpdateRW.asp?action=Hetong&sType=DelReason&Id=<%=rs("hId")%>',{title: '删除原因', width: 400,height: 150, fixed: true}); 
								<%else%>
								art.dialog.open('?action=delete&Id=<%=rs("hId")%>');
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
			<%=EasyCrm.pagelist("Hetong.asp", PN,TotalPages,TotalHetong)%>
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
	cID = EasyCrm.getNewItem("Hetong","hId",""&id&"","cID")
	conn.execute("DELETE FROM [Hetong] where hId = "&Id&" ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cID&"','"&L_Hetong&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
<script src="../data/calendar/WdatePicker.js"></script>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html><%else%>无权限<%end if%>
<% Set EasyCrm = nothing %>