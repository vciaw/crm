<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 31, 1) = 1 Then %>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Order.asp"
Session("CRM_pagenum") = PNN

If subAction = "searchItem" Then
    Dim Company,oCode,oLinkman,oState,User,TimeBegin,TimeEnd,ETimeBegin,ETimeEnd
	Company = EasyCrm.Searchcode(Request.Form("company"))
	oCode = EasyCrm.Searchcode(Request.Form("oCode"))
	oLinkman = EasyCrm.Searchcode(Request.Form("oLinkman"))
	oState = EasyCrm.Searchcode(Request.Form("oState"))
	oUser = EasyCrm.Searchcode(Request.Form("User"))
	TimeBegin = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	TimeEnd = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	ETimeBegin = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	ETimeEnd = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	Session("Search_Order_Company") = EasyCrm.Searchcode(Request.Form("Company"))
	Session("Search_Order_oCode") = EasyCrm.Searchcode(Request.Form("oCode"))
	Session("Search_Order_oLinkman") = EasyCrm.Searchcode(Request.Form("oLinkman"))
	Session("Search_Order_oState") = EasyCrm.Searchcode(Request.Form("oState"))
	Session("Search_Order_oUser") = EasyCrm.Searchcode(Request.Form("User"))
	Session("Search_Order_TimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Order_TimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	Session("Search_Order_ETimeBegin") = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	Session("Search_Order_ETimeEnd") = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	Dim sql
    sql = ""
	
	CompanyWhere = EasyCrm.seachKey("cCompany",Company)
	
    If Company <> "" Then
        sql = sql & " And cId in ( select cId from [Client] where 1=1 "&CompanyWhere&" )"
	End If
	
    If oCode <> "" Then
        sql = sql & " And oCode like '%" & oCode & "%' "
	End If
	
    If oLinkman <> "" Then
        sql = sql & " And oLinkman like '%" & oLinkman & "%' "
	End If
	
    If oState <> "" Then
	    sql = sql & " And oState = '" & oState & "'"
	End If
	
    If oUser <> "" Then
	    sql = sql & " And oUser = '" & oUser & "'"
	End If
	
	if Accsql =1 then
	If TimeBegin <> "" and  TimeEnd <> "" Then
	    sql = sql & " And oSDate >= '" & TimeBegin & "' And oSDate <= '" & TimeEnd & "' "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And oSDate = '" & TimeBegin & "' "
	End If
	else
	If TimeBegin <> "" and TimeEnd <> "" Then
	    sql = sql & " And oSDate >= #" & TimeBegin & "# And oSDate <= #" & TimeEnd & "# "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And oSDate = #" & TimeBegin & "# "
	End If
	end if
	
	if Accsql =1 then
	If ETimeBegin <> "" and  ETimeEnd <> "" Then
	    sql = sql & " And oEDate >= '" & ETimeBegin & "' And oEDate <= '" & ETimeEnd & "' "
	End If
	If ETimeBegin <> "" and  ETimeEnd = "" Then
	    sql = sql & " And oEDate = '" & ETimeBegin & "' "
	End If
	else
	If ETimeBegin <> "" and ETimeEnd <> "" Then
	    sql = sql & " And oEDate >= #" & ETimeBegin & "# And oEDate <= #" & ETimeEnd & "# "
	End If
	If ETimeBegin <> "" and  ETimeEnd = "" Then
	    sql = sql & " And oEDate = #" & ETimeBegin & "# "
	End If
	end if
		
	
	If Session("CRM_level") < 9 Then
		sql = sql & " And oUser In (" & arrUser & ")"
	End If
	
End If

If Company = "" And oCode = "" And oLinkman = "" And oState = "" And User = "" And TimeBegin = "" And TimeEnd = ""  And ETimeBegin = "" And ETimeEnd = "" Then
    If Session("CRM_Order_Search") <> "" Then
        sql = Session("CRM_Order_Search")
	Else
	    If Session("CRM_level") < 9 Then
			sql = " And oUser In (" & arrUser & ")"
		End If
	End If
Else
    Session("CRM_Order_Search") = sql
End If

If subAction = "killSession" Then
	Session("CRM_Order_Search") = ""
	Session("Search_Order_Company") = ""
	Session("Search_Order_oCode") = ""
	Session("Search_Order_oLinkman") = ""
	Session("Search_Order_oState") = ""
	Session("Search_Order_oUser") = ""
	Session("Search_Order_TimeBegin") = ""
	Session("Search_Order_TimeEnd") = ""
	Session("Search_Order_ETimeBegin") = ""
	Session("Search_Order_ETimeEnd") = ""
	If Session("CRM_level") < 9 Then
		sql = " And oUser In (" & arrUser & ")"
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
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Company%> > <%=L_Order%> </td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
			<% If mid(Session("CRM_qx"), 5, 1) = 1 Then %>
			<input type="button" class="button_top_set" value=" " title="设置" onclick='Setting_Order()' style="cursor:pointer" />
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
<script>function Setting_Order() {$.dialog.open('../system/GetUpdate.asp?action=Setting&sType=Order', {title: '自定义设置', width: 900, height: 480,fixed: true}); };</script>
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
									<td class="td_l_r title" style="border-top:0;"><%=L_Order_cID%></td>
									<td class="td_r_l" style="border-top:0;"><input name="Company" type="text" class="int" id="Company" size="30" value="<%=Session("Search_Order_Company")%>" ></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Order_oCode%></td>
									<td class="td_r_l"><input name="oCode" type="text" class="int" id="oCode" size="30" value="<%=Session("Search_Order_oCode")%>"></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Order_oSDate%></td>
									<td class="td_r_l"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Order_TimeBegin")%>" /> ~ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Order_TimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Order_oEDate%></td>
									<td class="td_r_l"><input name="ETimeBegin" type="text" maxlength="10" id="ETimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Order_ETimeBegin")%>" /> ~ <input name="ETimeEnd" type="text" maxlength="10" id="ETimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Order_ETimeEnd")%>" /></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Order_oLinkman%></td>
									<td class="td_r_l"><input name="oLinkman" type="text" class="int" id="oLinkman" size="30" value="<%=Session("Search_Order_oLinkman")%>"></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Order_oState%></td>
									<td class="td_r_l"><select name='oState'><option value="">请选择</option><option value="0">未处理</option><option value="1">处理中</option><option value="2">已完成</option></select></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Order_oUser%></td>
									<td class="td_r_l">
										<% If Session("CRM_level") = 9 Then %>
										<% = EasyCrm.UserList(2,"User",Session("Search_Order_oUser")) %>
										<% Else %>
										<% = EasyCrm.UserList(1,"User",Session("Search_Order_oUser")) %>
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
								<td class="td_l_l"><B>信息列表</B></td>
							</tr>
						</table> 
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_b">
								<td width="80" class="td_l_c"><%=L_Order_oId%></td>
								<td class="td_l_l"><%=L_Order_cId%></td>
								<%if Order_oCode = 1 then %>
								<td class="td_l_c"><%=L_Order_oCode%></td>
								<%end if%>
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
								<%if Order_oUser = 1 then %>
								<td class="td_l_c"><%=L_Order_oUser%></td>

                          <%							
Set rss = Server.CreateObject("ADODB.Recordset")
rss.Open "Select * From [CustomField] where cTable = 'Order' and cList = '1' order by Id asc ",conn,3,1
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
       





								<%end if%>
								<%if Order_oTime = 1 then %>
								<td class="td_l_c"><%=L_Order_oTime%></td>
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
							rs.Open "Select top "&intPageSize&" * From [Order] where cID<>'' "&sql&" Order By oId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Order] where cID<>'' "&sql&" and oId < ( SELECT Min(oId) FROM ( SELECT TOP "&pagenum&" oId FROM [Order]  where cID<>'' "&sql&" ORDER BY oId desc ) AS T ) Order By oId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(oId) As Orderum From [Order] where cID<>'' "&sql&" ",1,1)
							
							TotalOrder=Rsstr("Orderum") 
							if Int(TotalOrder/intPageSize)=TotalOrder/intPageSize then
							TotalPages=TotalOrder/intPageSize
							else
							TotalPages=Int(TotalOrder/intPageSize)+1
							end if
							Rsstr.Close
							Set Rsstr=Nothing
							If PN > TotalPages Then PN = TotalPages
							If rs.RecordCount > 0 Then
							Do While Not rs.BOF And Not rs.EOF
							%>
							<tr class="tr">
								<td class="td_l_c"><%=rs("oId")%></td>
								<td class="td_l_l"><a onclick='Order_InfoView<%=rs("cId")%>()' style="cursor:pointer"><%=EasyCrm.getNewItem("Client","cID",rs("cId"),"cCompany")%></a></td>
								<%if Order_oCode = 1 then %>
								<td class="td_l_c"><a title="订单产品明细"  onclick='Order_Products_List<%=rs("oId")%>()' style="cursor:pointer" ><%=rs("oCode")%></a></td>
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
								<td class="td_l_c"><% if Session("CRM_level")=9 then%><a onclick='Order_InfoAudit<%=rs("oId")%>()' style="cursor:pointer" ><%else%><a style="cursor:pointer"><%end if%><%if rs("oState") = 0 then%>未处理<%elseif rs("oState") = 1 then%>处理中<%elseif rs("oState") = 2 then%>已完成<%elseif rs("oState") = 3 then%>已取消<%end if%> </a></td>
								<%end if%>
								<%if Order_oUser = 1 then %>
								<td class="td_l_c"><%=rs("oUser")%></td>
								<%end if%>
                                
                               <%                            
                               	cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID",""&rs("cId")&" and oid = "&rs("oId")&" ","cContent")
								cContentArr = split(cContentStr,"|")	
															
								Set rss1 = Server.CreateObject("ADODB.Recordset")
								rss1.Open "Select * From [CustomField] where cTable = 'order' and cList = '1' order by id asc ",conn,1,1
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




								<%if Order_oTime = 1 then %>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("oTime"),2)%></td>
								<%end if%>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 33, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Order_InfoEdit<%=rs("oId")%>()' style="cursor:pointer" /><%end if%> <% If mid(Session("CRM_qx"), 34, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Order_InfoDel<%=rs("oId")%>()' style="cursor:pointer" /><%end if%></td>
							</tr>
							<script>function Order_InfoView<%=rs("cId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&otype=Order&cId=<%=rs("cId")%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
							<script>function Order_Products_List<%=rs("oId")%>() {$.dialog.open('GetUpdateRW.asp?action=OrderProducts&sType=List&Id=<%=rs("oId")%>', {title: '查看', width: 860,height: 440, fixed: true}); };</script>
							<script>function Order_InfoAudit<%=rs("oId")%>() {$.dialog.open('GetUpdateRW.asp?action=Order&sType=Audit&Id=<%=rs("oId")%>', {title: '审核', width: 400,height: 180, fixed: true}); };</script>
							<script>function Order_InfoEdit<%=rs("oId")%>() {$.dialog.open('GetUpdateRW.asp?action=Order&sType=Edit&Id=<%=rs("oId")%>', {title: '编辑', width: 700,height: 340, fixed: true}); };</script>
							<script>function Order_InfoDel<%=rs("oId")%>(){art.dialog({content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {
								<%if YnDelReason = 1 then%> 
								$.dialog.open('GetUpdateRW.asp?action=Order&sType=DelReason&Id=<%=rs("oId")%>',{title: '删除原因', width: 400,height: 150, fixed: true}); 
								<%else%>
								art.dialog.open('?action=delete&Id=<%=rs("oId")%>');
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
			<%=EasyCrm.pagelist("Order.asp", PN,TotalPages,TotalOrder)%> 
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
	cID = EasyCrm.getNewItem("Order","oID",""&id&"","cID")
	conn.execute("DELETE FROM [Order] where oId = "&Id&" ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cID&"','"&L_Order&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
<script src="../data/calendar/WdatePicker.js"></script>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html><%else%>无权限<%end if%><% Set EasyCrm = nothing %>
