<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 26, 1) = 1 Then %>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Records.asp"
Session("CRM_pagenum") = PNN

If subAction = "searchItem" Then
    Dim Company,rLinkman,rState,rType,rUser,TimeBegin,TimeEnd,ETimeBegin,ETimeEnd
	Company = EasyCrm.Searchcode(Request.Form("company"))
	rLinkman = EasyCrm.Searchcode(Request.Form("rLinkman"))
	rState = EasyCrm.Searchcode(Request.Form("rState"))
	rType = EasyCrm.Searchcode(Request.Form("rType"))
	rUser = EasyCrm.Searchcode(Request.Form("User"))
	TimeBegin = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	TimeEnd = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	ETimeBegin = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	ETimeEnd = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	Session("Search_Records_Company") = EasyCrm.Searchcode(Request.Form("Company"))
	Session("Search_Records_rLinkman") = EasyCrm.Searchcode(Request.Form("rLinkman"))
	Session("Search_Records_rState") = EasyCrm.Searchcode(Request.Form("rState"))
	Session("Search_Records_rType") = EasyCrm.Searchcode(Request.Form("rType"))
	Session("Search_Records_rUser") = EasyCrm.Searchcode(Request.Form("User"))
	Session("Search_Records_TimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Records_TimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	Session("Search_Records_ETimeBegin") = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	Session("Search_Records_ETimeEnd") = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	Dim sql
    sql = ""
	
	CompanyWhere = EasyCrm.seachKey("cCompany",Company)
	
    If Company <> "" Then
        sql = sql & " And cId in ( select cId from [Client] where 1=1 "&CompanyWhere&" )"
	End If
	
    If rLinkman <> "" Then
        sql = sql & " And rLinkman like '%" & rLinkman & "%' "
	End If
	
    If rState <> "" Then
	    sql = sql & " And rState = '" & rState & "'"
	End If
	
    If rType <> "" Then
	    sql = sql & " And rType = '" & rType & "'"
	End If
	
    If rUser <> "" Then
	    sql = sql & " And rUser = '" & rUser & "'"
	End If
	
	if Accsql =1 then
	If TimeBegin <> "" and  TimeEnd <> "" Then
	    sql = sql & " And rTime >= '" & TimeBegin & "' And rTime <= '" & TimeEnd & "' "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And rTime = '" & TimeBegin & "' "
	End If
	else
	If TimeBegin <> "" and TimeEnd <> "" Then
	    sql = sql & " And rTime >= #" & TimeBegin & "# And rTime <= #" & TimeEnd & "# "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And rTime = #" & TimeBegin & "# "
	End If
	end if
	
	if Accsql =1 then
	If ETimeBegin <> "" and  ETimeEnd <> "" Then
	    sql = sql & " And rNextTime >= '" & ETimeBegin & "' And rNextTime <= '" & ETimeEnd & "' "
	End If
	If ETimeBegin <> "" and  ETimeEnd = "" Then
	    sql = sql & " And DATEDIFF(d,rNextTime,'"&ETimeBegin&"')=0 "
	End If
	else
	If ETimeBegin <> "" and ETimeEnd <> "" Then
	    sql = sql & " And rNextTime >= #" & ETimeBegin & "# And rNextTime <= #" & ETimeEnd & "# "
	End If
	If ETimeBegin <> "" and  ETimeEnd = "" Then
	    sql = sql & " And DATEDIFF('d',rNextTime,'"&ETimeBegin&"')=0 "
	End If
	end if
		
	
	If Session("CRM_level") < 9 Then
		sql = sql & " And rUser In (" & arrUser & ")"
	End If
	
End If

If Company = "" And rLinkman = "" And rState = "" And rType = "" And rUser = "" And TimeBegin = "" And TimeEnd = ""  And ETimeBegin = "" And ETimeEnd = "" Then
    If Session("CRM_Records_Search") <> "" Then
        sql = Session("CRM_Records_Search")
	Else
	    If Session("CRM_level") < 9 Then
			sql = " And rUser In (" & arrUser & ")"
		End If
	End If
Else
    Session("CRM_Records_Search") = sql
End If

If subAction = "killSession" Then
	Session("CRM_Records_Search") = ""
	Session("Search_Records_Company") = ""
	Session("Search_Records_rLinkman") = ""
	Session("Search_Records_rState") = ""
	Session("Search_Records_rType") = ""
	Session("Search_Records_rUser") = ""
	Session("Search_Records_TimeBegin") = ""
	Session("Search_Records_TimeEnd") = ""
	Session("Search_Records_ETimeBegin") = ""
	Session("Search_Records_ETimeEnd") = ""
	If Session("CRM_level") < 9 Then
		sql = " And rUser In (" & arrUser & ")"
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
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Company%> > <%=L_Records%> </td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
			<% If mid(Session("CRM_qx"), 5, 1) = 1 Then %>
			<input type="button" class="button_top_set" value=" " title="设置" onclick='Setting_Records()' style="cursor:pointer" />
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
					<li class="" id="CheckA"><span><a href="javascript:void(0)" style="cursor:pointer">高级搜索</a></span></li>
				</ul>
			</div>
		</td>
	</tr>
</table>
<script>function Setting_Records() {$.dialog.open('../system/GetUpdate.asp?action=Setting&sType=Records', {title: '自定义设置', width: 900, height: 480,fixed: true}); };</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr> 
		<td valign="top" class="td_n">
		
			<div id="SearchBox" style="position: absolute; width:100%; height:450px; background:#ffffff; display:none; z-index:10;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0" style="background:#ffffff;">
					<tr>
						<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
							<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
								<tr class="tr_t"> 
									<td class="td_l_l" COLSPAN="6" style="border-right:0;"><B><%=L_Top_Search%></B></td>
								</tr>
							</table>
							<form name="searchForm" action="?subAction=searchItem" method="post">
							<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
								<col width="120" />
								<tr>
									<td class="td_l_r title" style="border-top:0;"><%=L_Records_cID%></td>
									<td class="td_r_l" style="border-top:0;"><input name="Company" type="text" class="int" id="Company" size="30" value="<%=Session("Search_Records_Company")%>" ></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Records_rTime%></td>
									<td class="td_r_l"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Records_TimeBegin")%>" /> ~ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Records_TimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Records_rNextTime%></td>
									<td class="td_r_l"><input name="ETimeBegin" type="text" maxlength="10" id="ETimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Records_ETimeBegin")%>" /> ~ <input name="ETimeEnd" type="text" maxlength="10" id="ETimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Records_ETimeEnd")%>" /></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Records_rLinkman%></td>
									<td class="td_r_l"><input name="rLinkman" type="text" class="int" id="rLinkman" size="30" value="<%=Session("Search_Records_rLinkman")%>"></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Records_rState%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Type","rState",Session("Search_Records_rState")) %></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Records_rType%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Records","rType",Session("Search_Records_rType")) %></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Records_rUser%></td>
									<td class="td_r_l">
										<% If Session("CRM_level") = 9 Then %>
										<% = EasyCrm.UserList(2,"User",Session("Search_Records_rUser")) %>
										<% Else %>
										<% = EasyCrm.UserList(1,"User",Session("Search_Records_rUser")) %>
										<% End If %>
									</td>
								</tr>
								<tr>
									<td class="td_r_l" colspan="4">
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
								<td class="td_l_l" style="border-right:0;"><B>信息列表</B></td>
							</tr>
						</table> 
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_b">
								<td width="80" class="td_l_c"><%=L_Records_rId%></td>
								<td class="td_l_l"><%=L_Records_cId%></td>
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
								<td class="td_l_l"><%=L_Records_rContent%></td>
								<%end if%>
								<%if Records_rUser = 1 then %>
								<td class="td_l_c"><%=L_Records_rUser%></td>
								<%end if%>


                                 <%							
Set rss = Server.CreateObject("ADODB.Recordset")
rss.Open "Select * From [CustomField] where cTable = 'Records' and cList = '1' order by Id asc ",conn,3,1
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











								<%if Records_rTime = 1 then %>
								<td class="td_l_c"><%=L_Records_rTime%></td>
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
							rs.Open "Select top "&intPageSize&" * From [Records] where cID<>'' "&sql&" Order By rId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Records] where cID<>'' "&sql&" and rId < ( SELECT Min(rId) FROM ( SELECT TOP "&pagenum&" rId FROM [Records]  where cID<>'' "&sql&" ORDER BY rId desc ) AS T ) Order By rId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(rId) As RecordSum From [Records] where cID<>'' "&sql&" ",1,1)
							
							TotalRecords=Rsstr("RecordSum") 
							if Int(TotalRecords/intPageSize)=TotalRecords/intPageSize then
							TotalPages=TotalRecords/intPageSize
							else
							TotalPages=Int(TotalRecords/intPageSize)+1
							end if
							Rsstr.Close
							Set Rsstr=Nothing
							If PN > TotalPages Then PN = TotalPages
							If rs.RecordCount > 0 Then
							Do While Not rs.BOF And Not rs.EOF
							%>
							<tr class="tr">
								<td class="td_l_c"><%=rs("rId")%></td>
								<td class="td_l_l"><a onclick='Records_InfoView<%=rs("cId")%>()' style="cursor:pointer"><%=EasyCrm.getNewItem("Client","cID",rs("cId"),"cCompany")%></a></td>
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


                                 <%                            
                               	cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID",""&rs("cId")&" and rid = "&rs("rId")&" ","cContent")
								cContentArr = split(cContentStr,"|")	
															
								Set rss1 = Server.CreateObject("ADODB.Recordset")
								rss1.Open "Select * From [CustomField] where cTable = 'Records' and cList = '1' order by id asc ",conn,1,1
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
























								<%if Records_rTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("rTime"),2)%></td>
								<%end if%>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 28, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Records_InfoEdit<%=rs("rId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 29, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Records_InfoDel<%=rs("rId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Records_InfoView<%=rs("cId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&otype=Records&cId=<%=rs("cId")%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
							<script>function Records_InfoEdit<%=rs("rId")%>() {$.dialog.open('GetUpdateRW.asp?action=Records&sType=Edit&Id=<%=rs("rId")%>', {title: '编辑', width: 800,height: 340, fixed: true}); };</script>
							<script>function Records_InfoDel<%=rs("rId")%>(){art.dialog({content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {
								<%if YnDelReason = 1 then%> 
								$.dialog.open('GetUpdateRW.asp?action=Records&sType=DelReason&Id=<%=rs("rId")%>',{title: '删除原因', width: 400,height: 150, fixed: true}); 
								<%else%>
								art.dialog.open('?action=delete&Id=<%=rs("rid")%>');
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
			<%=EasyCrm.pagelist("Records.asp", PN,TotalPages,TotalRecords)%> 
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
	cID = EasyCrm.getNewItem("Records","rID",""&id&"","cID")
	conn.execute("DELETE FROM [Records] where rId = "&Id&" ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cID&"','"&L_Records&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
<script src="../data/calendar/WdatePicker.js"></script>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html><%else%>无权限<%end if%><% Set EasyCrm = nothing %>
