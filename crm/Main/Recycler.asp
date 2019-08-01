<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Recycler.asp"
Session("CRM_pagenum") = PNN

If subAction = "searchItem" Then
    Dim cCompany,czhuying,cLinkman,cTel,cArea,cSquare,cStart,cType,cSource,cUser,cTimeBegin,cTimeEnd,cETimeBegin,cETimeEnd,SHYN
	cCompany = EasyCrm.Searchcode(Request.Form("company"))
	czhuying = EasyCrm.Searchcode(Request.Form("czhuying"))
	cLinkman = EasyCrm.Searchcode(Request.Form("Linkman"))
	cMobile = EasyCrm.Searchcode(Request.Form("Mobile"))
	cTel = EasyCrm.Searchcode(Request.Form("Tel"))
	cArea = EasyCrm.Searchcode(Request.Form("Area"))
	cSquare = EasyCrm.Searchcode(Request.Form("Squares"))
	cStart = EasyCrm.Searchcode(Request.Form("Start"))
	cType = EasyCrm.Searchcode(Request.Form("Type"))
	cSource = EasyCrm.Searchcode(Request.Form("Source"))
	cUser = EasyCrm.Searchcode(Request.Form("User"))
	cTimeBegin = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	cTimeEnd = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	cETimeBegin = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	cETimeEnd = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	SHYN = EasyCrm.Searchcode(Request("SHYN"))
	
	Session("Search_Recycler_cCompany") = EasyCrm.Searchcode(Request.Form("company"))
	Session("Search_Recycler_czhuying") = EasyCrm.Searchcode(Request.Form("czhuying"))
	Session("Search_Recycler_cLinkman") = EasyCrm.Searchcode(Request.Form("Linkman"))
	Session("Search_Recycler_cMobile") = EasyCrm.Searchcode(Request.Form("Mobile"))
	Session("Search_Recycler_cTel") = EasyCrm.Searchcode(Request.Form("Tel"))
	Session("Search_Recycler_cArea") = EasyCrm.Searchcode(Request.Form("Area"))
	Session("Search_Recycler_cSquare") = EasyCrm.Searchcode(Request.Form("Squares"))
	Session("Search_Recycler_cStart") = EasyCrm.Searchcode(Request.Form("Start"))
	Session("Search_Recycler_cType") = EasyCrm.Searchcode(Request.Form("Type"))
	Session("Search_Recycler_cSource") = EasyCrm.Searchcode(Request.Form("Source"))
	Session("Search_Recycler_cUser") = EasyCrm.Searchcode(Request.Form("User"))
	Session("Search_Recycler_cTimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Recycler_cTimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	Session("Search_Recycler_cETimeBegin") = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	Session("Search_Recycler_cETimeEnd") = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	Session("Search_Recycler_SHYN") = EasyCrm.Searchcode(Request.Form("SHYN"))
	
	Dim sql
    sql = ""
	
	cCompanyWhere = EasyCrm.seachKey("cCompany",cCompany)
	
    If cCompany <> "" Then
        sql = sql & cCompanyWhere
	End If	
	
    If cLinkman <> "" Then
        sql = sql & " And cId in ( select cId from [Client] where cLinkman like '%" & cLinkman & "%' )"
	End If
	
    If cMobile <> "" Then
        sql = sql & " And cId in ( select cId from [Client] where cMobile like '%" & cMobile & "%' )"
	End If
		If czhuying <> "" Then
	    sql = sql & " And cInfo Like '%" & czhuying & "%'"
	End If
	If cTel <> "" Then
	    sql = sql & " And cTel Like '%" & cTel & "%'"
	End If
	
    If cArea <> "" Then
	    sql = sql & " And cArea = '" & cArea & "'"
	End If
	
    If cSquare <> "" Then
	    sql = sql & " And cSquare = '" & cSquare & "'"
	End If
	
    If cstart <> "" Then
	    sql = sql & " And cStart = '" & cstart & "'"
	End If
	
    If cType <> "" Then
	    sql = sql & " And cType = '" & cType & "'"
	End If
	
    If cSource <> "" Then
	    sql = sql & " And cSource = '" & cSource & "'"
	End If
	
	if Accsql =1 then
	If cTimeBegin <> "" and  cTimeEnd <> "" Then
	    sql = sql & " And cDate >= '" & cTimeBegin & "' And cDate <= '" & cTimeEnd & "' "
	End If
	If cTimeBegin <> "" and  cTimeEnd = "" Then
	    sql = sql & " And cDate = '" & cTimeBegin & "' "
	End If
	else
	If cTimeBegin <> "" and cTimeEnd <> "" Then
	    sql = sql & " And cDate >= #" & cTimeBegin & "# And cDate <= #" & cTimeEnd & "# "
	End If
	If cTimeBegin <> "" and  cTimeEnd = "" Then
	    sql = sql & " And cDate = #" & cTimeBegin & "# "
	End If
	end if
	
	if Accsql =1 then
	If cETimeBegin <> "" and  cETimeEnd <> "" Then
	    sql = sql & " And cLastUpdated >= '" & cETimeBegin & "' And cLastUpdated <= '" & cETimeEnd & "' "
	End If
	If cETimeBegin <> "" and  cETimeEnd = "" Then
	    sql = sql & " And DATEDIFF(d,cLastUpdated,'"&cETimeBegin&"')=0 "
	End If
	else
	If cETimeBegin <> "" and cETimeEnd <> "" Then
	    sql = sql & " And cLastUpdated >= #" & cETimeBegin & "# And cLastUpdated <= #" & cETimeEnd & "# "
	End If
	If cETimeBegin <> "" and  cETimeEnd = "" Then
	    sql = sql & " And DATEDIFF('d',cLastUpdated,'"&cETimeBegin&"')=0 "
	End If
	end if
	
	If cUser <> "" Then
	    sql = sql & " And cUser = '"&cUser &"' "
	End If
	
    If SHYN <> "" Then
		if SHYN = ""&L_Shi&"" then
	    sql = sql & " And cOldUser <> '' "
		else
	    sql = sql & " And ( cOldUser = '' or cOldUser is null ) "
		end if
	End If
	
End If

If cCompany = "" And czhuying = "" And cLinkman = ""  And cMobile = "" And cTel = "" And cArea = "" And cSquare = ""  And cType = "" And cSource = "" And cUser = "" And cStart = "" And cTimeBegin = "" And cTimeEnd = "" And cETimeBegin = "" And cETimeEnd = "" And SHYN = "" Then
    If Session("CRM_Search_Recycler") <> "" Then
        sql = Session("CRM_Search_Recycler")
	End If
Else
    Session("CRM_Search_Recycler") = sql
End If

If subAction = "killSession" Then
	Session("CRM_Search_Recycler") = ""
	Session("Search_Recycler_cCompany") = ""
	Session("Search_Recycler_czhuying") = ""
	Session("Search_Recycler_cLinkman") = ""
	Session("Search_Recycler_cMobile") = ""
	Session("Search_Recycler_cTel") = ""
	Session("Search_Recycler_cArea") = ""
	Session("Search_Recycler_cSquare") = ""
	Session("Search_Recycler_cStart") = ""
	Session("Search_Recycler_cType") = ""
	Session("Search_Recycler_cSource") = ""
	Session("Search_Recycler_cUser") = ""
	Session("Search_Recycler_cTimeBegin") = ""
	Session("Search_Recycler_cTimeEnd") = ""
	Session("Search_Recycler_cETimeBegin") = ""
	Session("Search_Recycler_cETimeEnd") = ""
	Session("Search_Recycler_SHYN") = ""
	Session("Search_Recycler_NewUser") = ""
	sql=""
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
<script language="JavaScript">
<!--
function CheckAll(form) {
for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
}
-->
</script>
</head>

<body oncontextmenu=self.event.returnValue=false> 
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Company%> > <%=L_Page_Recycler%> </td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>
<%
Select Case action
Case "CheckSub"		'批量操作
    Call CheckSubject()
Case "ReDel"		'撤销删除
    Call ReDel()
Case "ReApp"		'客户申请
    Call ReApp()
Case "ReConfirm"	'通过申请
    Call ReConfirm()
Case "ReDenied"		'拒绝申请
    Call ReDenied()
Case "RealDel"		'彻底删除
    Call RealDel()
Case Else
	Call Main()
End Select

Sub Main()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdt10">   
			<div class="MenuboxS">
				 <ul>
					<li <%if otype="Main" or otype="" then%>class="hover"<%end if%> id="CheckB"><span><a href="?otype=Main">客户列表</a></span></li>
					<li class="" id="CheckA"><span><a href="javascript:void(0)" style="cursor:pointer">高级搜索</a></span></li>
				</ul>
			</div>
		</td>
	</tr>
</table>
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
								<col width="120" /><col width="260" /><col width="120" />
								<tr>
									<td class="td_l_r title" style="border-top:0;"><%=L_Client_cCompany%></td>
									<td class="td_r_l" style="border-top:0;"><input name="Company" type="text" class="int" id="Company" size="30" value="<%=Session("Search_Recycler_cCompany")%>" ></td>
									<td class="td_l_r title" style="border-top:0;"><%=L_Client_cDate%></td>
									<td class="td_r_l" style="border-top:0;"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Recycler_cTimeBegin")%>" /> ~ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Recycler_cTimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cLinkman%></td>
									<td class="td_r_l"><input name="Linkman" type="text" class="int" id="Linkman" size="30" value="<%=Session("Search_Recycler_cLinkman")%>" ></td>
									<td class="td_l_r title"><%=L_Client_cLastUpdated%></td>
									<td class="td_r_l"><input name="ETimeBegin" type="text" maxlength="10" id="ETimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Recycler_cETimeBegin")%>" /> ~ <input name="ETimeEnd" type="text" maxlength="10" id="ETimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Recycler_cETimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cMobile%></td>
									<td class="td_r_l"><input name="Mobile" type="text" class="int" id="Mobile" size="30" value="<%=Session("Search_Recycler_cMobile")%>" ></td>
									<td class="td_l_r title"><%=L_Client_cArea%><%=L_Client_cSquare%></td>
									<td class="td_r_l">
										<select name="Area" onchange="getArea(this.options[this.selectedIndex].id);">
										<option value=""><%=L_Please_choose_01%></option>
										<% 
											Set rsb = Conn.Execute("select * from AreaData where aFId = '0' ")
											If Not rsb.Eof then
											Do While Not rsb.Eof
											aId= rsb("aId")
											aName= rsb("aName")
										%>
											<option value="<%=aName%>" id="<%=aId%>"><%=aName%></option>
										<%
											rsb.Movenext
											Loop
											End If
											rsb.Close
											Set rss = Nothing 
										%>
										</select> 
										<span id="Squarediv"  style="margin-left:10px;padding:0;">
											<select name="Squares">
												<option value=""><%=L_Please_choose_02%></option>
												<% 
												IF ""&cArea&""<>"" then
												Set rss = Conn.Execute("select * from AreaData where aFId='"&EasyCrm.getNewItem("AreaData","aName","'"&cArea&"'","aId")&"' ")
												If Not rss.Eof then
												Do While Not rss.Eof
												aName= rss("aName")
												%>
												<option value="<%=aName%>"><%=aName%></option>
												<%rss.Movenext
												Loop
												End If
												rss.Close
												Set rss = Nothing 
												End If
												%>
											</select>
											
										</span>
										
									</td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cInfo%></td>
									<td class="td_r_l"><input name="czhuying" type="text"  id="czhuying" size="30" value="<%=Session("Search_Recycler_czhuying")%>"></td>
									<td class="td_l_r title"></td>
									<td class="td_r_l"></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cTel%></td>
									<td class="td_r_l"><input name="Tel" type="text" class="int" id="Tel" size="30" value="<%=Session("Search_Recycler_cTel")%>"></td>
									<td class="td_l_r title"><%=L_Client_cType%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Type","Type",Session("Search_Recycler_cType")) %></td>
								</tr>
								<tr>
									<td class="td_l_r title">审核中</td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_YN","SHYN",Session("Search_Recycler_SHYN")) %></td>
									<td class="td_l_r title"><%=L_Client_cStart%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Star","Start",Session("Search_Recycler_cStart")) %></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cUser%></td>
									<td class="td_r_l">
										<% If Session("CRM_level") = 9 Then %>
										<% = EasyCrm.UserList(2,"User",Session("Search_Recycler_cUser")) %>
										<% Else %>
										<% = EasyCrm.UserList(1,"User",Session("Search_Recycler_cUser")) %>
										<% End If %>
									</td>
									<td class="td_l_r title"><%=L_Client_cSource%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Source","Source",Session("Search_Recycler_cSource")) %></td>
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
								<td class="td_l_l">
								<% If mid(Session("CRM_qx"), 9, 1) = 1 Then %>
									<span class="tips01" style="float:left;display:none;padding:0 10px;height:34px;line-height:34px;text-align:left;position:fixed;right:10px;top:43px;color:#000;" id="CheckSub">
									<% If mid(Session("CRM_qx"), 12, 1) = 1 Then %>
										新：<% If Session("CRM_level") = 9 Then %><% = EasyCrm.UserList(2,"NewUser","") %><%else%><% = EasyCrm.UserList(1,"NewUser","") %><%end if%>
										<input type="submit" name="Checkexecute" id="Checkexecute" class="button224" value="<%=L_Transfer%>">
									<%end if%>
									<% If mid(Session("CRM_qx"), 15, 1) = 1 Then %>
										<input type="submit" name="Checkexecute" id="Checkexecute" class="button242" value="<%=L_ReConfirm%>">
										<input type="submit" name="Checkexecute" id="Checkexecute" class="button243" value="<%=L_ReDenied%>">
									<%end if%>
									<%if Session("CRM_level") = 9 then%>
										<input type="submit" name="Checkexecute" id="Checkexecute" class="button247" value="<%=L_RealDel%>">
									<%end if%>
									</span>
								<%end if%>
								<B>信息列表</B>
								</td>
							</tr>
						</table> 
						<script type="text/javascript">
						 function getBlock(ck, d) {
							var c = document.getElementById(ck);
							var d = document.getElementById(d);
							if (c.checked == true) {
								d.style.display = "block"
							} 
							//else {d.style.display = "none"}
						}
						</script>

						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_b">
								<td width="40" class="td_l_c"><input name="chkall" type="checkbox" id="chkall" onclick="CheckAll(this.form);getBlock('chkall','CheckSub')" value="checkbox"></td>
								<td width="80" class="td_l_c"><%=L_Client_cId%></td>
								<td class="td_l_l"><%=L_Client_cCompany%></td>
								<td width="80" class="td_l_c"><%=L_Client_cLastUpdated%></td>
								<td width="100" class="td_l_c"><%=L_Client_cUser%></td>
								<td width="100" class="td_l_c"><%=L_Client_cOldUser%></td>
								<td width="100" class="td_l_c"><%=L_Top_Manage%></td>
							</tr>
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [client] where cYn = 0 "&sql&" Order By cId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [client] where cYn = 0 "&sql&" and cid < ( SELECT Min(cid) FROM ( SELECT TOP "&pagenum&" cid FROM [client]  where cYn = 0 "&sql&" ORDER BY cId desc ) AS T ) Order By cId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(cid) As RecordSum From [client] where cYn = 0 "&sql&" ",1,1)
						
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
								<td class="td_l_c"><input type="checkbox" name="cId" id="cId<%=rs("cId")%>" value="<%=rs("cId")%>" onclick="getBlock('cId<%=rs("cId")%>','CheckSub')"></td>
								<td class="td_l_c"><%=rs("cId")%></td>
								<td class="td_l_l"><a onclick='Client_InfoView<%=rs("cId")%>()' style="cursor:pointer"><%=rs("cCompany")%></a><%if rs("cShare") = 1 then%><span class="info_share" title="已共享">&nbsp;</span><%end if%><%If EasyCrm.getCountItem("File","fId","fIdstr"," and cID="&rs("cID")&" ") > 0 Then%><span class="info_file" title="有附件">&nbsp;</span><%end if%></td>
								<td class="td_l_c"><font color=red><%=Datediff("d",rs("cLastUpdated"),Now())%> </font>天前</td>
								<td class="td_l_c"><%=rs("cUser")%></td>
								<td class="td_l_c"><%=rs("cOldUser")%></td>
								<td class="td_l_c">
								<%if rs("cOldUser")="" or IsNull(rs("cOldUser")) then%>
									<%if rs("cUser")=Session("CRM_name") then%>
									<input type="button" class="button_info_restore" value=" " title="<%=L_ReDel%>"  onclick='Recycler_InfoReDel<%=rs("cId")%>()' style="cursor:pointer" /> 
									<%else%>
									<%if YNRecycler = 1 then%>
									<input type="button" class="button_info_add" value=" " title="<%=L_ReApp%>"  onclick='Recycler_InfoReApp<%=rs("cId")%>()' style="cursor:pointer" /> 
									<%else%>
									<input type="button" class="button_info_add" value=" " title="<%=L_ReApp%>"  onclick='Recycler_InfoReConfirm<%=rs("cId")%>()' style="cursor:pointer" /> 
									<%end if%>
									<%end if%>
									<%if Session("CRM_level") = 9 then%>
									<input type="button" class="button_info_delete" value=" " title="<%=L_RealDel%>" onclick='Recycler_InfoRealDel<%=rs("cId")%>()' style="cursor:pointer" />
									<%end if%>
								<%else%>
								<% If mid(Session("CRM_qx"), 15, 1) = 1 Then %>
									<input type="button" class="button_info_yes" value=" " title="<%=L_ReConfirm%>"  onclick='Recycler_InfoReConfirm<%=rs("cId")%>()' style="cursor:pointer" /> 
									<input type="button" class="button_info_del" value=" " title="<%=L_ReDenied%>" onclick='Recycler_InfoReDenied<%=rs("cId")%>()' style="cursor:pointer" />
								<%end if%>
								<%end if%>
								</td>
							</tr>
							<script>function Client_InfoView<%=rs("cId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&cId=<%=rs("cId")%>&YNRange=0', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
							<script>function Recycler_InfoReDel<%=rs("cId")%>() //撤销删除
							{
								art.dialog({
									content: '<%=Alert_del_restore%>',
									icon: 'face-smile',
									ok: function () {
										art.dialog.open('?action=ReDel&cId=<%=rs("cid")%>');
										art.dialog.close();
									},
									cancelVal: '关闭',
									cancel: true
								});
							};
							</script>
							<script>function Recycler_InfoReApp<%=rs("cId")%>() //申请客户
							{
										art.dialog.open('?action=ReApp&cId=<%=rs("cid")%>&uId=<%=Session("CRM_uId")%>');
										art.dialog.close();
							};
							</script>
							<script>function Recycler_InfoReConfirm<%=rs("cId")%>() //通过申请
							{		
									<%if YNRecycler = 1 then%>
										art.dialog.open('?action=ReConfirm&cId=<%=rs("cid")%>');
									<%else%>
										art.dialog.open('?action=ReConfirm&cId=<%=rs("cid")%>&uId=<%=Session("CRM_uId")%>');
									<%end if%>
										art.dialog.close();
							};
							</script>
							<script>function Recycler_InfoReDenied<%=rs("cId")%>() //拒绝申请
							{
										art.dialog.open('?action=ReDenied&cId=<%=rs("cid")%>');
										art.dialog.close();
							};
							</script>
							<script>function Recycler_InfoRealDel<%=rs("cId")%>() //彻底删除
							{
								art.dialog({
									content: '<%=Alert_del_client_Del%>',
									icon: 'warning',
									ok: function () {
										art.dialog.open('?action=RealDel&cId=<%=rs("cid")%>');
										art.dialog.close();
									},
									cancelVal: '关闭',
									cancel: true
								});
							};
							</script>
							<%
							rs.MoveNext
							Loop
							else
							%>
							<tr><td class="td_l_l" colspan="8"><%=L_Notfound%></td></tr>
							<%
							end if
							rs.Close
							Set rs = Nothing
							%>
							
						</table> 
						</form>
					</td>
				</tr>
			</table>
        </td>
	</tr>
</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd ">
			<%if sql<>"" then%><span class="r"><input name="Back" type="button" id="Back" class="button227" value="清空" onClick=window.location.href="?SubAction=killSession"></span><%end if%>
			<%=EasyCrm.pagelist("Recycler.asp", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>

<script language="JavaScript">
<!--
for(var i=0;i<document.getElementById('Area').options.length;i++){
    if(document.getElementById('Area').options[i].value == "<% = Session("Search_Recycler_cArea") %>"){
    document.getElementById('Area').options[i].selected = true;}}
for(var i=0;i<document.getElementById('Squares').options.length;i++){
    if(document.getElementById('Squares').options[i].value == "<% = Session("Search_Recycler_cSquare") %>"){
    document.getElementById('Squares').options[i].selected = true;}}
-->
</script>
<%
end Sub 

Sub CheckSubject()
cId=Trim(Request("cId"))
NewUser=Request("NewUser")

PN = CLng(ABS(Request("PN")))

If Request("Checkexecute")=""&L_Transfer&"" Then
	If cId="" Then
		Response.Write "<script>alert('"&alert04&"');location.href='?PN="&PN&"' ;</script>"
		Response.End
	elseif NewUser="" then
		Response.Write "<script>alert('"&alert04&"');location.href='?PN="&PN&"' ;</script>"
		Response.End
	else
		cidarr = split(cid,",")
		' 限制客户量，则循环判断是否已经达到最大客户量
		if CLng(ABS(EasyCrm.getNewItem("User","uName","'"&NewUser&"'","uClientNum"))) > 0 then
		for i = 0 to Ubound(cidarr) 
		
			if CLng(ABS(EasyCrm.getCountItem("Client","cid","cidstr"," and cYn=1 and cUser='"&NewUser&"' "))) < CLng(ABS(EasyCrm.getNewItem("User","uName","'"&NewUser&"'","uClientNum"))) then
				set rs=conn.execute("update [Client] set cYn = '1' ,cUser = '"&NewUser&"' where cId =" & cidarr(i) & " ")
				if SaveOldUser = 0 then
				set rs=conn.execute("update [Linkmans] set lUser = '"&NewUser&"' where cId =" & cidarr(i) & " ")
				set rs=conn.execute("update [Records] set rUser = '"&NewUser&"' where cId =" & cidarr(i) & " ")
				set rs=conn.execute("update [Order] set oUser = '"&NewUser&"' where cId =" & cidarr(i) & " ")
				set rs=conn.execute("update [Hetong] set hUser = '"&NewUser&"' where cId =" & cidarr(i) & " ")
				set rs=conn.execute("update [Service] set sUser = '"&NewUser&"' where cId =" & cidarr(i) & " ")
				set rs=conn.execute("update [Expense] set eUser = '"&NewUser&"' where cId =" & cidarr(i) & " ")
				set rs=conn.execute("update [File] set fUser = '"&NewUser&"' where cId =" & cidarr(i) & " ")
				end if
			else
				if ""&YNalert&"" = 1 then
				Response.Write "<script>alert('成功转移"&i&"条记录，"&NewUser&"的客户量达到最大配额！');</script>"
				end if
				Response.Write("<script>location.href='?PN="&PN&"' ;</script>")
				Response.end
			end if
		next
		
		else
				set rs=conn.execute("update [Client] set cYn = '1' ,cUser = '"&NewUser&"' where cId In(" & cId & ")")
				if SaveOldUser = 0 then
				set rs=conn.execute("update [Linkmans] set lUser = '"&NewUser&"' where cId In(" & cId & ") ")
				set rs=conn.execute("update [Records] set rUser = '"&NewUser&"' where cId In(" & cId & ") ")
				set rs=conn.execute("update [Order] set oUser = '"&NewUser&"' where cId In(" & cId & ") ")
				set rs=conn.execute("update [Hetong] set hUser = '"&NewUser&"' where cId In(" & cId & ") ")
				set rs=conn.execute("update [Service] set sUser = '"&NewUser&"' where cId In(" & cId & ") ")
				set rs=conn.execute("update [Expense] set eUser = '"&NewUser&"' where cId In(" & cId & ") ")
				set rs=conn.execute("update [File] set fUser = '"&NewUser&"' where cId In(" & cId & ") ")
				end if
				if ""&YNalert&"" = 1 then
				Response.Write("<script>alert('"&alert4&"');</script>")
				end if
				Response.Write("<script>location.href='?PN="&PN&"' ;</script>")
				Response.end
		end if
	End If
elseIf Request("Checkexecute")=""&L_ReConfirm&"" Then	'通过审核，最大客户量已判断，这里直接通过
	If cId="" Then
		Response.Write "<script>alert('"&alert04&"');location.href='?PN="&PN&"' ;</script>"
		Response.End
	else
		cidarr = split(cid,",")
		' 限制客户量，则循环判断是否已经达到最大客户量
		for i = 0 to Ubound(cidarr) 
			cUser = EasyCrm.getNewItem("Client","cId",""&cidarr(i)&"","cOlduser") '申请人
			if cUser <> "" then
			cGroup = EasyCrm.getNewItem("User","uName","'"&cUser&"'","uGroup") '申请人所在的部门
		
				conn.execute("update Client set cYn='1' ,cUser = '"&cUser&"',cOlduser = '',cGroup='"&cGroup&"' where cId = " & cidarr(i) & " ")
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cidarr(i)&"','"&L_Client&"','"&L_ReConfirm&"','"&Session("CRM_name")&"','"&now()&"')")
			else
				Response.Write "<script>alert('"&alert04&"');location.href='?PN="&PN&"' ;</script>"
				Response.End
			end if
		next
	End If
				Response.Write("<script>location.href='?PN="&PN&"' ;</script>")
				Response.end

elseIf Request("Checkexecute")=""&L_ReDenied&"" Then	'拒绝申请
	If cId="" Then
		Response.Write "<script>alert('"&alert04&"');location.href='?PN="&PN&"' ;</script>"
		Response.End
	else
				conn.execute("update Client set cOlduser = '' where cId In(" & cId & ") ")
	End If
				Response.Write("<script>location.href='?PN="&PN&"' ;</script>")
				Response.end
				
elseIf Request("Checkexecute")=""&L_RealDel&"" Then		'彻底删除
	If cId="" Then
		Response.Write "<script>alert("""&alert_trans_no_select&""");</script><SCRIPT LANGUAGE=JavaScript>document.Search.chkall.focus()</SCRIPT>"
		Response.End
	else
		set rs=conn.execute("Delete from [Client] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Linkmans] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Records] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Order] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Hetong] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Service] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Expense] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [File] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Logfile] Where lcId In(" & cId & ")")
		if ""&YNalert&"" = 1 then
			Response.Write("<script>alert('"&alert4&"');</script>")
		end if
		Response.Write("<script>location.href='?PN="&PN&"' ;</script>")
	end if
End If
   
End Sub

Sub ReApp() '申请客户
    Dim cid,uId
	cid = Trim(Request("cid"))
	uId = Trim(Request("uId"))
	If cid = "" or uId = "" Then
	Exit Sub
	End If
	cUser = EasyCrm.getNewItem("User","uId",""&uId&"","uName") '申请人
	ReCNum = EasyCrm.getCountItem("Client","cid","cidstr"," and cYn=1 and cUser='"&cUser&"' ") '申请人客户量
	ReRNum = EasyCrm.getCountItem("Client","cid","cidstr"," and cYn=0 and cOlduser='"&cUser&"' ") '已申请客户量
	'判断是否限制了当前申请人的客户量
	if CLng(ABS(EasyCrm.getNewItem("User","uId",""&uId&"","uClientNum"))) > 0 then
		if CLng(ABS(ReCNum))+CLng(ABS(ReRNum)) < CLng(ABS(EasyCrm.getNewItem("User","uId",""&uId&"","uClientNum"))) then
			conn.execute("update Client set cOlduser = '"&cUser&"' where cId = "&cId&" ")
			conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_ReApp&"','"&cUser&"','"&now()&"')")
		else
			Response.Write "<script>alert('您的客户量达到最大配额！');</script>"
		end if
	else
			conn.execute("update Client set cOlduser = '"&cUser&"' where cId = "&cId&" ")
			conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_ReApp&"','"&cUser&"','"&now()&"')")
	end if
			
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

Sub ReConfirm() '通过申请
    Dim cid,uId
	cid = Trim(Request("cid"))
	uId = Trim(Request("uId"))
	If cid = "" Then
	Exit Sub
	End If
if YNRecycler = 1 then '公海申请需要审核
	cUser = EasyCrm.getNewItem("Client","cId",""&cId&"","cOlduser") '申请人
	cGroup = EasyCrm.getNewItem("User","uName","'"&cUser&"'","uGroup") '申请人所在的部门
	conn.execute("update Client set cYn='1' ,cUser = '"&cUser&"',cOlduser = '',cGroup='"&cGroup&"' where cId = "&cId&" ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_ReConfirm&"','"&Session("CRM_name")&"','"&now()&"')")
	
else '公海申请直接通过
	cUser = EasyCrm.getNewItem("User","uId",""&uId&"","uName") '申请人
	cGroup = EasyCrm.getNewItem("User","uName","'"&cUser&"'","uGroup") '申请人所在的部门
	ReCNum = EasyCrm.getCountItem("Client","cid","cidstr"," and cYn=1 and cUser='"&cUser&"' ") '申请人客户量
	ReRNum = EasyCrm.getCountItem("Client","cid","cidstr"," and cYn=0 and cOlduser='"&cUser&"' ") '已申请客户量
	'判断是否限制了当前申请人的客户量
	if CLng(ABS(EasyCrm.getNewItem("User","uId",""&uId&"","uClientNum"))) > 0 then
		if CLng(ABS(ReCNum))+CLng(ABS(ReRNum)) < CLng(ABS(EasyCrm.getNewItem("User","uId",""&uId&"","uClientNum"))) then
			conn.execute("update Client set cYn='1' ,cUser = '"&cUser&"',cOlduser = '',cGroup='"&cGroup&"' where cId = "&cId&" ")
			conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_ReApp&"','"&cUser&"','"&now()&"')")
		else
			Response.Write "<script>alert('您的客户量达到最大配额！');</script>"
		end if
	else
			conn.execute("update Client set cYn='1' ,cUser = '"&cUser&"',cOlduser = '',cGroup='"&cGroup&"' where cId = "&cId&" ")
			conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_ReApp&"','"&cUser&"','"&now()&"')")
	end if

end if 

	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

Sub ReDenied() '拒绝申请
    Dim cid
	cid = Trim(Request("cid"))
	If cid = "" Then
	Exit Sub
	End If
	conn.execute("update Client set cOlduser = '' where cId = "&cId&" ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_ReDenied&"','"&Session("CRM_name")&"','"&now()&"')")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

Sub ReDel() '撤销删除
    Dim cid
	cid = Trim(Request("cid"))
	If cid = "" Then
	Exit Sub
	End If
	cUser = EasyCrm.getNewItem("Client","cID",""&cID&"","cUser")
	ReCNum = EasyCrm.getCountItem("Client","cid","cidstr"," and cYn=1 and cUser='"&cUser&"' ") '申请人客户量
	'判断是否限制了当前客户的原业务员客户量
	if CLng(ABS(EasyCrm.getNewItem("User","uName","'"&cUser&"'","uClientNum"))) > 0 then
		if CLng(ABS(ReCNum)) < CLng(ABS(EasyCrm.getNewItem("User","uName","'"&cUser&"'","uClientNum"))) then
			conn.execute("update Client set cYn = 1 where cId = "&cId&" ")
			conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_ReDel&"','"&Session("CRM_name")&"','"&now()&"')")
		else
			Response.Write "<script>alert('"&cUser&"的客户量达到最大配额！');</script>"
		end if
	else
			conn.execute("update Client set cYn = 1 where cId = "&cId&" ")
			conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_ReDel&"','"&Session("CRM_name")&"','"&now()&"')")
	end if
			
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

Sub RealDel() '彻底删除
    Dim cid
	cid = Trim(Request("cid"))
	If cid = "" Then
	Exit Sub
	End If
		set rs=conn.execute("Delete from [Client] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Linkmans] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Records] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Order] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Hetong] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Service] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Expense] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [File] Where cId In(" & cId & ")")
		set rs=conn.execute("Delete from [Logfile] Where lcId In(" & cId & ")")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
<script src="../data/calendar/WdatePicker.js"></script>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html><% Set EasyCrm = nothing %>
