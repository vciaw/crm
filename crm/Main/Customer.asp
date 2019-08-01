<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Customer.asp"
Session("CRM_pagenum") = PNN
MyShare = Trim(Request.QueryString("MyShare"))
HeShare = Trim(Request.QueryString("HeShare"))

If subAction = "searchItem" Then
    Dim cCompany,cLinkman,cAddress,cTel,cFax,cInfo,cBeizhu,cType,cHomepage,cTrade,cStrade,cSource,cUser,cGroup,arrUser,cdate,cTimeBegin,cTimeEnd,cYn,SearchRange
	cCompany = EasyCrm.Searchcode(Request.Form("company"))
	cTel = EasyCrm.Searchcode(Request.Form("cTel"))
	cType = EasyCrm.Searchcode(Request.Form("cType"))  
	cHomepage = EasyCrm.Searchcode(Request.Form("cHomepage"))
	cFax = EasyCrm.Searchcode(Request.Form("cFax"))
	cInfo = EasyCrm.Searchcode(Request.Form("cInfo"))
	cBeizhu = EasyCrm.Searchcode(Request.Form("cBeizhu"))
	cUser = EasyCrm.Searchcode(Request.Form("User"))
	cGroup = EasyCrm.Searchcode(Request.Form("group"))
	If cGroup <> "" Then
	    cGroup = CInt(Abs(cGroup))
	End If
	cTimeBegin = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	cTimeEnd = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	cETimeBegin = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	cETimeEnd = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	SearchRange = EasyCrm.Searchcode(Request.Form("SearchRange"))
	
	Session("Search_Customer_cCompany") = EasyCrm.Searchcode(Request.Form("company"))
	Session("Search_Customer_cTel") = EasyCrm.Searchcode(Request.Form("cTel"))
	Session("Search_Customer_cType") = EasyCrm.Searchcode(Request.Form("cType"))  
	Session("Search_Customer_cHomepage") = EasyCrm.Searchcode(Request.Form("cHomepage"))
	Session("Search_Customer_cFax") = EasyCrm.Searchcode(Request.Form("cFax"))
	Session("Search_Customer_cInfo") = EasyCrm.Searchcode(Request.Form("cInfo"))
	Session("Search_Customer_cBeizhu") = EasyCrm.Searchcode(Request.Form("cBeizhu"))
	Session("Search_Customer_cUser") = EasyCrm.Searchcode(Request.Form("User"))
	Session("Search_Customer_cGroup") = EasyCrm.Searchcode(Request.Form("group"))
	Session("Search_Customer_cTimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Customer_cTimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	Session("Search_Customer_cETimeBegin") = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	Session("Search_Customer_cETimeEnd") = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	if EasyCrm.Searchcode(Request.Form("SearchRange")) <> "" then
	Session("Search_Customer_SearchRange") = EasyCrm.Searchcode(Request.Form("SearchRange"))
	end if
	
	Dim sql
    sql = ""
	
	cCompanyWhere = EasyCrm.seachKey("cCompany",cCompany)
	
    If cCompany <> "" Then
        sql = sql & cCompanyWhere
	End If	
    If cAddress <> "" Then
	    sql = sql & " And cAddress Like '%" & cAddress & "%'"
	End If
	
	If cHomepage <> "" Then
	    sql = sql & " And cHomepage Like '%" & cHomepage & "%'"
	End If
	
	If cTel <> "" Then
	    sql = sql & " And cTel Like '%" & cTel & "%'"
	End If
	
    If cFax <> "" Then
	    sql = sql & " And cFax   Like '%" & cFax & "%'"
	End If
	
    If cInfo <> "" Then
	    sql = sql & " And  cInfo Like '%" & cInfo & "%'"
	End If
	
    If cBeizhu <> "" Then
	    sql = sql & " And cBeizhu Like '%" & cBeizhu & "%'"
	End If
	
    If cType <> "" Then
	    sql = sql & " And cType = '" & cType & "'"
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
		
	If cGroup <> "" And IsNumeric(cGroup) Then
	    sql = sql & " And cGroup = "&cGroup&" "
	End If
	
	If Session("CRM_level") < 9 Then
		if MyShare <> "" then 
			sql = sql & " and cShare = 1 and cUser = '"&Session("CRM_Name")&"' "
		end if
		if HeShare <> "" then 
			sql = sql & " and cShare = 1 and cShareRange like '%"&Session("CRM_Name")&"%' "
		end if
		
		if Session("Search_Customer_SearchRange") = "0" then
		sql = sql & " "
		else
			if MyShare = "" or HeShare = "" then 
		sql = sql & " And cUser In (" & arrUser & ")"
			end if
		end if
	End If
	
End If
	
	
If cCompany = ""And cTel = "" And cHomepage = "" And cFax = "" And cInfo = ""  And cType = "" And cBeizhu = ""And cUser = "" And cGroup = "" And cTimeBegin = "" And cTimeEnd = "" And cETimeBegin = "" And cETimeEnd = ""  And SearchRange = "" Then
    If Session("CRM_Search") <> "" Then
        sql = Session("CRM_Search")
	Else
	    If Session("CRM_level") < 9 Then
			if MyShare <> "" then 
				sql = " and cShare = 1 and cUser = '"&Session("CRM_Name")&"' "
			elseif HeShare <> "" then 
				sql = " and cShare = 1 and cShareRange like '%"&Session("CRM_Name")&"%' "
			else
				sql = " And cUser In (" & arrUser & ")"
			end if
		End If
	End If
Else
    Session("CRM_Search") = sql
End If

If subAction = "killSession" Then
	Session("CRM_Search") = ""
	Session("Search_Customer_cCompany") = ""
	Session("Search_Customer_cTel") = ""
	Session("Search_Customer_cType") =  ""
	Session("Search_Customer_cHomepage") =  ""
	Session("Search_Customer_cFax") =  ""
	Session("Search_Customer_cInfo") =  ""
	Session("Search_Customer_cBeizhu") = ""
	Session("Search_Customer_cUser") = ""
	Session("Search_Customer_cGroup") = ""
	Session("Search_Customer_cTimeBegin") = ""
	Session("Search_Customer_cTimeEnd") = ""
	Session("Search_Customer_cETimeBegin") = ""
	Session("Search_Customer_cETimeEnd") = ""
	Session("Search_Customer_SearchRange") = ""
	Session("Search_Customer_NewUser") = ""
	If Session("CRM_level") < 9 Then
			if MyShare <> "" then 
				sql = " and cShare = 1 and cUser = '"&Session("CRM_Name")&"' "
			elseif HeShare <> "" then 
				sql = " and cShare = 1 and cShareRange like '%"&Session("CRM_Name")&"%' "
			else
				sql = " And cUser In (" & arrUser & ")"
			end if
	else
	sql=""
	end if
End If


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<title><%=EasyCrm.URLDecode("%cb%f9%d3%d0%bf%cd%bb%a7")%></title>
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

<body> 
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Company%> > <%=L_Page_Customer%></td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>
<%
Select Case action
Case "CheckSub"
    Call CheckSubject()
Case "delete"
    Call deleteData()
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
					<% If mid(Session("CRM_qx"), 11, 1) = 1 Then %>
					<li class="" id="CheckA"><span><a href="javascript:void(0)" style="cursor:pointer">高级搜索</a></span></li>
					<% end If %>
					<% If mid(Session("CRM_qx"), 17, 1) = 1 Then %>
					<li class="" id="CheckC"><span><a href="#" onclick='Customer_InfoAdd()' style="cursor:pointer">新增客户</a></span></li>
					<% end If %>
				</ul>
			</div>
		</td>
	</tr>
</table>
<script>function Customer_InfoAdd() {$.dialog.open('GetUpdatenew.asp?action=Customer&sType=Add', {title: '新增', width: 900, height: 480,fixed: true}); };</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr> 
		<td valign="top" class="td_n">
		
			<div id="SearchBox" style="position: absolute; width:100%; height:500px; background:#ffffff; display:none; z-index:10;">
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
									<td class="td_l_r title" style="border-top:0;"><%=L_Customer_cCompany%></td>
									<td class="td_r_l" style="border-top:0;"><input name="Company" type="text" class="int" id="Company" size="30" value="<%=Session("Search_Customer_cCompany")%>" ></td>
									<td class="td_l_r title" style="border-top:0;"><%=L_Customer_cDate%></td>
									<td class="td_r_l" style="border-top:0;"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Customer_cTimeBegin")%>" /> ~ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Customer_cTimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Customer_cLastUpdated%></td>
									<td class="td_r_l"><input name="ETimeBegin" type="text" maxlength="10" id="ETimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Customer_cETimeBegin")%>" /> ~ <input name="ETimeEnd" type="text" maxlength="10" id="ETimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Customer_cETimeEnd")%>" /> </td>
								    <td class="td_l_r title"><%=L_Customer_cType%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Mtype","cType",Session("Search_Customer_cType")) %></td>
								
								</tr>
								<tr>
									<td class="td_l_r title" ><%=L_Customer_cTel%></td>
									<td class="td_r_l">
									<input name="cTel" type="text" class="int" id="cTel" size="30" value="<%=Session("Search_Customer_cTel")%>" ></td>
									<td class="td_l_r title"><%=L_Customer_cHomepage%></td>
									<td class="td_r_l" >
                                    <input name="cHomepage" type="text" class="int" id="cHomepage" size="30" value="<%=Session("Search_Customer_cHomepage")%>" >
									
									</td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Customer_cFax%></td>
									<td class="td_r_l" >
									<input name="cFax" type="text" class="int" id="cFax" size="30" value="<%=Session("Search_Customer_cFax")%>" ></td>
									<td class="td_l_r title" ><%=L_Customer_cInfo%></td>
									<td class="td_r_l" >
                                    <input name="cInfo" type="text" class="int" id="cInfo" size="30" value="<%=Session("Search_Customer_cInfo")%>" >
									</td>
								</tr>
								<tr>
									<td class="td_l_r title" ><%=L_Customer_cBeizhu%></td>
									<td class="td_r_l"  colspan="3">
									<input name="cBeizhu" type="text" class="int" id="cBeizhu" size="30" value="<%=Session("Search_Customer_cBeizhu")%>" >
									</td>
									
								</tr>
								
								<tr>
									<td class="td_l_r title"><%=L_Customer_cUser%></td>
									<td class="td_r_l">
										<% If Session("CRM_level") = 9 Then %>
										<% = EasyCrm.UserList(2,"User",Session("Search_Customer_cUser")) %>
										<% Else %>
										<% = EasyCrm.UserList(1,"User",Session("Search_Customer_cUser")) %>
										<% End If %>
									</td>
									<td class="td_l_r title">查询范围</td>
									<td class="td_r_l"> 
										<input type="radio" name="SearchRange" value="1" <%if Session("Search_Customer_SearchRange") = "1" or Session("Search_Customer_SearchRange") = "" then%>checked<%end if%>> 权限内　
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
								<td class="td_l_l">
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
								<td width="80" class="td_l_c"><%=L_Customer_cId%></td>
								<td width="100" class="td_l_c"><%=L_Customer_cDate%></td>
								<td class="td_l_l"><%=L_Customer_cCompany%></td>
								<td width="120" class="td_l_c"><%=L_Customer_cTel%></td>
								<td width="100" class="td_l_c"><%=L_Customer_cType%></td>
							   	<td width="120" class="td_l_c"><%=L_Customer_cHomepage%></td>
								<td width="120" class="td_l_c"><%=L_Customer_cFax%></td>
								<td width="120" class="td_l_c"><%=L_Customer_cInfo%></td>
								<td width="120" class="td_l_c"><%=L_Customer_cBeizhu%></td>
								<td width="100" class="td_l_c"><%=L_Customer_cUser%></td>
								<td width="100" class="td_l_c"><%=L_Top_Manage%></td>
							</tr>
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							
						if Session("Search_Customer_SearchRange") = "0" then
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [Customer] where 1 = 1 "&sql&" "&Share&" Order By cId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Customer] where 1 = 1 "&sql&" "&Share&" and cid < ( SELECT Min(cid) FROM ( SELECT TOP "&pagenum&" cid FROM [Customer]  where 1 = 1 "&sql&" "&Share&" ORDER BY cId desc ) AS T ) Order By cId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(cid) As RecordSum From [Customer] where 1 = 1 "&sql&" "&Share&" ",1,1)
						else
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [Customer] where cYn = 1 "&sql&" "&Share&" Order By cId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Customer] where cYn = 1 "&sql&" "&Share&" and cid < ( SELECT Min(cid) FROM ( SELECT TOP "&pagenum&" cid FROM [Customer]  where cYn = 1 "&sql&" "&Share&" ORDER BY cId desc ) AS T ) Order By cId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(cid) As RecordSum From [Customer] where cYn = 1 "&sql&" "&Share&" ",1,1)
						end if
						
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
								
								<td class="td_l_c"><%=rs("cDate")%></td>
								
								
						
								<td class="td_l_l">
									<a onclick='Customer_InfoView<%=rs("cId")%>()' style="cursor:pointer"><%=rs("cCompany")%></a>
								</td>
							
						
							
								
								<td class="td_l_c"><%=rs("cTel")%></td>
								<td class="td_l_c"><%=rs("cType")%></td>
							    <td class="td_l_c"><%=rs("cHomepage")%></td>
								<td class="td_l_c"><%=rs("cFax")%></td>
								<td class="td_l_c"><%=rs("cInfo")%></td>
								<td class="td_l_c"><%=rs("cBeizhu")%></td>	
								<td class="td_l_c"><%=rs("cUser")%></td>
							   



							   <%if Session("CRM_level")<9 then%>
								<td class="td_l_c">
									<% if inStr(""&arrUser&"",rs("cUser"))>0 then %>
									<% If mid(Session("CRM_qx"), 18, 1) = 1 Then %>
									<input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Customer_InfoEdit<%=rs("cId")%>()' style="cursor:pointer" />
									<%end if%>
									<% If mid(Session("CRM_qx"), 19, 1) = 1 Then %>
									<input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Customer_InfoDel<%=rs("cId")%>()' style="cursor:pointer" />
									<%end if%>
									<%end if%>
								</td>
							<%else%>
								<td class="td_l_c"><% If mid(Session("CRM_qx"), 18, 1) = 1 Then %><input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Customer_InfoEdit<%=rs("cId")%>()' style="cursor:pointer" /> <%end if%><% If mid(Session("CRM_qx"), 19, 1) = 1 Then %><input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Customer_InfoDel<%=rs("cId")%>()' style="cursor:pointer" /><%end if%></td>
							<%end if%>
							</tr>
							<script>function Customer_InfoView<%=rs("cId")%>() {$.dialog.open('GetUpdatenew.asp?action=Customer&sType=InfoView&cId=<%=rs("cId")%><%if rs("cYn") = 0 then%>&YNRange=0<%end if%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
							<script>function Customer_InfoEdit<%=rs("cId")%>() {$.dialog.open('GetUpdatenew.asp?action=Customer&sType=InfoEdit&cId=<%=rs("cId")%>', {title: '编辑', width: 900,height: 480, fixed: true}); };</script>
							
							
							
							<script>function Customer_InfoDel<%=rs("cId")%>()
							{
								art.dialog({
									content: '<%=Alert_del_YN%>',
									icon: 'error',
									ok: function () {
										<%if YnDelReason = 1 then%> 
										$.dialog.open('GetUpdatenew.asp?action=Customer&sType=DelReason&cId=<%=rs("cId")%>', 
										{
											title: '删除原因', 
											width: 400,
											height: 150, 
											fixed: true
										}); 
										<%else%>
										art.dialog.open('?action=delete&cId=<%=rs("cid")%>');
										<%end if%>
										//return false;
										art.dialog.close();
									},
									cancelVal: '关闭',
									cancel: true //为true等价于function(){}
								});
							};
							</script>
							<%
							rs.MoveNext
							Loop
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
			<%=EasyCrm.pagelist("Customer.asp", PN,TotalPages,TotalRecords)%>

		</td>
	</tr>
</table>
</div>
<%
end Sub 
Sub deleteData()
    Dim cid
	cid = Trim(Request("cid"))
	If cid = "" Then
	Exit Sub
	End If
	conn.execute("update Customer set cYn = 0 where cId = "&cId&" ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Customer&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
<script src="../data/calendar/WdatePicker.js"></script>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html>
<% Set EasyCrm = nothing %>