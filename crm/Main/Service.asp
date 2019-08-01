<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 41, 1) = 1 Then %>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Service.asp"
Session("CRM_pagenum") = PNN

If subAction = "searchItem" Then
    Dim Company,sTitle,sLinkman,sType,sSolve,User,TimeBegin,TimeEnd,ETimeBegin,ETimeEnd
	Company = EasyCrm.Searchcode(Request.Form("Company"))
	sTitle = EasyCrm.Searchcode(Request.Form("sTitle"))
	sLinkman = EasyCrm.Searchcode(Request.Form("sLinkman"))
	sType = EasyCrm.Searchcode(Request.Form("sType"))
	sSolve = EasyCrm.Searchcode(Request.Form("sSolve"))
	sUser = EasyCrm.Searchcode(Request.Form("User"))
	TimeBegin = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	TimeEnd = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	ETimeBegin = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	ETimeEnd = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	Session("Search_Service_Company") = EasyCrm.Searchcode(Request.Form("Company"))
	Session("Search_Service_sTitle") = EasyCrm.Searchcode(Request.Form("sTitle"))
	Session("Search_Service_sLinkman") = EasyCrm.Searchcode(Request.Form("sLinkman"))
	Session("Search_Service_sType") = EasyCrm.Searchcode(Request.Form("sType"))
	Session("Search_Service_sSolve") = EasyCrm.Searchcode(Request.Form("sSolve"))
	Session("Search_Service_sUser") = EasyCrm.Searchcode(Request.Form("User"))
	Session("Search_Service_TimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Service_TimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	Session("Search_Service_ETimeBegin") = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	Session("Search_Service_ETimeEnd") = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	Dim sql
    sql = ""
	
	CompanyWhere = EasyCrm.seachKey("cCompany",Company)
	
    If Company <> "" Then
        sql = sql & " And cId in ( select cId from [Client] where 1=1 "&CompanyWhere&" )"
	End If
	
    If sTitle <> "" Then
        sql = sql & " And sTitle like '%" & sTitle & "%' "
	End If
	
    If sLinkman <> "" Then
        sql = sql & " And sLinkman like '%" & sLinkman & "%' "
	End If
	
    If sType <> "" Then
	    sql = sql & " And sType = '" & sType & "'"
	End If
	
    If sSolve <> "" Then
	    sql = sql & " And sSolve = '" & sSolve & "'"
	End If
	
    If sUser <> "" Then
	    sql = sql & " And sUser = '" & sUser & "'"
	End If
	
	if Accsql =1 then
	If TimeBegin <> "" and  TimeEnd <> "" Then
	    sql = sql & " And sSDate >= '" & TimeBegin & "' And sSDate <= '" & TimeEnd & "' "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And DATEDIFF(d,sSDate,'"&TimeBegin&"')=0 "
	End If
	else
	If TimeBegin <> "" and TimeEnd <> "" Then
	    sql = sql & " And sSDate >= #" & TimeBegin & "# And sSDate <= #" & TimeEnd & "# "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And DATEDIFF('d',sSDate,'"&TimeBegin&"')=0 "
	End If
	end if
	
	if Accsql =1 then
	If ETimeBegin <> "" and  ETimeEnd <> "" Then
	    sql = sql & " And sEDate >= '" & ETimeBegin & "' And sEDate <= '" & ETimeEnd & "' "
	End If
	If ETimeBegin <> "" and  ETimeEnd = "" Then
	    sql = sql & " And DATEDIFF(d,sEDate,'"&ETimeBegin&"')=0 "
	End If
	else
	If ETimeBegin <> "" and ETimeEnd <> "" Then
	    sql = sql & " And sEDate >= #" & ETimeBegin & "# And sEDate <= #" & ETimeEnd & "# "
	End If
	If ETimeBegin <> "" and  ETimeEnd = "" Then
	    sql = sql & " And DATEDIFF('d',sEDate,'"&ETimeBegin&"')=0 "
	End If
	end if
		
	
	If Session("CRM_level") < 9 Then
		if mid(Session("CRM_qx"), 13, 1) = "0" then
		sql = sql & " And sUser In (" & arrUser & ")"
		end if
	End If
	
End If

If Company = "" And sTitle = "" And sLinkman = "" And sType = "" And sSolve = "" And User = "" And TimeBegin = "" And TimeEnd = ""  And ETimeBegin = "" And ETimeEnd = "" Then
    If Session("CRM_Service_Search") <> "" Then
        sql = Session("CRM_Service_Search")
	Else
	    If Session("CRM_level") < 9 Then
		if mid(Session("CRM_qx"), 13, 1) = "0" then
			sql = " And sUser In (" & arrUser & ")"
		end if
		End If
	End If
Else
    Session("CRM_Service_Search") = sql
End If

If subAction = "killSession" Then
	Session("CRM_Service_Search") = ""
	Session("Search_Service_Company") = ""
	Session("Search_Service_sTitle") = ""
	Session("Search_Service_sLinkman") = ""
	Session("Search_Service_sType") = ""
	Session("Search_Service_sSolve") = ""
	Session("Search_Service_sUser") = ""
	Session("Search_Service_TimeBegin") = ""
	Session("Search_Service_TimeEnd") = ""
	Session("Search_Service_ETimeBegin") = ""
	Session("Search_Service_ETimeEnd") = ""
	If Session("CRM_level") < 9 Then
		if mid(Session("CRM_qx"), 13, 1) = "0" then
		sql = " And sUser In (" & arrUser & ")"
		end if
	else
		sql=""
	end if
End If


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
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
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Company%> > <%=L_Service%> </td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
			<% If mid(Session("CRM_qx"), 5, 1) = 1 Then %>
			<input type="button" class="button_top_set" value=" " title="设置" onclick='Setting_Service()' style="cursor:pointer" />
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
<script>function Setting_Service() {$.dialog.open('../system/GetUpdate.asp?action=Setting&sType=Service', {title: '自定义设置', width: 900, height: 480,fixed: true}); };</script>
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
									<td class="td_l_r title" style="border-top:0;"><%=L_Service_cID%></td>
									<td class="td_r_l" style="border-top:0;"><input name="Company" type="text" class="int" id="Company" size="30" value="<%=Session("Search_Service_Company")%>" ></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Service_sTitle%></td>
									<td class="td_r_l"><input name="sTitle" type="text" class="int" id="sTitle" size="30" value="<%=Session("Search_Service_sTitle")%>"></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Service_sLinkman%></td>
									<td class="td_r_l"><input name="sLinkman" type="text" class="int" id="sLinkman" size="30" value="<%=Session("Search_Service_sLinkman")%>"></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Service_sType%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Service","sType",Session("Search_Service_sType")) %></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Service_sSDate%></td>
									<td class="td_r_l"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Service_TimeBegin")%>" /> ~ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Service_TimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Service_sEDate%></td>
									<td class="td_r_l"><input name="ETimeBegin" type="text" maxlength="10" id="ETimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Service_ETimeBegin")%>" /> ~ <input name="ETimeEnd" type="text" maxlength="10" id="ETimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Service_ETimeEnd")%>" /></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Service_sSolve%></td>
									<td class="td_r_l"><select name='sSolve'><option value="">请选择</option><option value="0"><%=L_Service_sSolve_0%></option><option value="1"><%=L_Service_sSolve_1%></option></select></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Service_sUser%></td>
									<td class="td_r_l">
										<% If Session("CRM_level") = 9 Then %>
										<% = EasyCrm.UserList(2,"User",Session("Search_Service_sUser")) %>
										<% Else %>
										<% = EasyCrm.UserList(1,"User",Session("Search_Service_sUser")) %>
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
								<td width="80" class="td_l_c"><%=L_Service_sId%></td>
								<td class="td_l_l"><%=L_Service_cId%></td>
								<%if Service_sSolve = 1 then %>
								<td width="60" class="td_l_c"><%=L_Service_sSolve%></td>
								<%end if%>
								<%if Service_sTitle = 1 then %>
								<td class="td_l_c"><%=L_Service_sTitle%></td>
								<%end if%>
								<%if Service_sLinkman = 1 then %>
								<td class="td_l_c"><%=L_Service_sLinkman%></td>
								<%end if%>

								<%							
Set rss = Server.CreateObject("ADODB.Recordset")
rss.Open "Select * From [CustomField] where cTable = 'Service' and cList = '1' order by Id asc ",conn,3,1
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
								<%if Service_sType = 1 then %>
								<td class="td_l_c"><%=L_Service_sType%></td>
								<%end if%>
								<%if Service_sSDate = 1 then %>
								<td class="td_l_c"><%=L_Service_sSDate%></td>
								<%end if%>
							<%If mid(Session("CRM_qx"), 13, 1) = "1" Then%>
								<td class="td_l_c">处理</td>
							<%end if%>

								<%if Service_sUser = 1 then %>
								<td class="td_l_c"><%=L_Service_sUser%></td>
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
							rs.Open "Select top "&intPageSize&" * From [Service] where cID<>'' "&sql&" Order By sId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Service] where cID<>'' "&sql&" and sId < ( SELECT Min(sId) FROM ( SELECT TOP "&pagenum&" sId FROM [Service]  where cID<>'' "&sql&" Order BY sId desc ) AS T ) Order By sId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(sId) As Servicenum From [Service] where cID<>'' "&sql&" ",1,1)
							
							TotalService=Rsstr("Servicenum") 
							if Int(TotalService/intPageSize)=TotalService/intPageSize then
							TotalPages=TotalService/intPageSize
							else
							TotalPages=Int(TotalService/intPageSize)+1
							end if
							Rsstr.Close
							Set Rsstr=Nothing
							If PN > TotalPages Then PN = TotalPages
							If rs.RecordCount > 0 Then
							Do While Not rs.BOF And Not rs.EOF
							%>
							<tr class="tr">
								<td class="td_l_c"><%=rs("sId")%></td>
								<td class="td_l_l"><a onclick='Client_InfoView<%=rs("cId")%>()' style="cursor:pointer"><%=EasyCrm.getNewItem("Client","cID",rs("cId"),"cCompany")%></a></td>
								<%if Service_sSolve = 1 then %>
								<td class="td_l_c"><img src="<%=SiteUrl&skinurl%>images/ico/<%if rs("sSolve") = 0 then%>no<%else%>yes<%end if%>.gif" border=0></td>
								<%end if%>
								<%if Service_sTitle = 1 then %>
								<td class="td_l_c"><%=rs("sTitle")%></td>
								<%end if%>
								<%if Service_sLinkman = 1 then %>
								<td class="td_l_c"><%=rs("sLinkman")%></td>
								<%end if%>


 <%                            
                               	cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID",""&rs("cId")&" and sid = "&rs("sId")&" ","cContent")
								cContentArr = split(cContentStr,"|")	
															
								Set rss1 = Server.CreateObject("ADODB.Recordset")
								rss1.Open "Select * From [CustomField] where cTable = 'Service' and cList = '1' order by id asc ",conn,1,1
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




































								<%if Service_sType = 1 then %>
								<td class="td_l_c"><%=rs("sType")%></td>
								<%end if%>
								<%if Service_sSDate = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("sSDate"),2)%></td>
								<%end if%>
							<%If mid(Session("CRM_qx"), 13, 1) = "1" Then%>
								<td class="td_l_c"><input type="button" class="button222" value="处理" onclick='Service_InfoAudit<%=rs("sId")%>()' style="cursor:pointer" /></td>
							<%end if%>
								<%if Service_sUser = 1 then %>
								<td class="td_l_c"><%=rs("sUser")%></td>
								<%end if%>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 43, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Service_InfoEdit<%=rs("sId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 44, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Service_InfoDel<%=rs("sId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Client_InfoView<%=rs("cId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&otype=Service&cId=<%=rs("cId")%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
							
							<script>function Service_InfoEdit<%=rs("sId")%>() {$.dialog.open('GetUpdateRW.asp?action=Service&sType=Edit&Id=<%=rs("sId")%>', {title: '编辑', width: 800,height: 370, fixed: true}); };</script>
							
							<script>function Service_InfoAudit<%=rs("sId")%>() {$.dialog.open('GetUpdateRW.asp?action=Service&sType=Audit&Id=<%=rs("sId")%>', {title: '问题处理', width: 800,height: 370, fixed: true}); };</script>
							
							<script>function Service_InfoDel<%=rs("sId")%>(){art.dialog({content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {
								<%if YnDelReason = 1 then%> 
								$.dialog.open('GetUpdateRW.asp?action=Service&sType=DelReason&Id=<%=rs("sId")%>',{title: '删除原因', width: 400,height: 150, fixed: true}); 
								<%else%>
								art.dialog.open('?action=delete&Id=<%=rs("sId")%>');
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
			<%=EasyCrm.pagelist("Service.asp", PN,TotalPages,TotalService)%>
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
	cID = EasyCrm.getNewItem("Service","sId",""&id&"","cID")
	conn.execute("DELETE FROM [Service] where sId = "&Id&" ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cID&"','"&L_Service&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
<script src="../data/calendar/WdatePicker.js"></script>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html><%else%>无权限<%end if%><% Set EasyCrm = nothing %>
