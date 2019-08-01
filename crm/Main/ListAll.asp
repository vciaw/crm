<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Listall.asp"
Session("CRM_pagenum") = PNN
MyShare = Trim(Request.QueryString("MyShare"))
HeShare = Trim(Request.QueryString("HeShare"))

If subAction = "searchItem" Then
    Dim cCompany,cLinkman,cAddress,cTel,cArea,cSquare,cStart,cType,cTrade,cStrade,cSource,cUser,cGroup,arrUser,cdate,cTimeBegin,cTimeEnd,cYn,SearchRange
	cCompany = EasyCrm.Searchcode(Request.Form("company"))
	cLinkman = EasyCrm.Searchcode(Request.Form("Linkman"))
	cMobile = EasyCrm.Searchcode(Request.Form("Mobile"))
	cAddress = EasyCrm.Searchcode(Request.Form("address"))
	cTel = EasyCrm.Searchcode(Request.Form("Tel"))
	cArea = EasyCrm.Searchcode(Request.Form("Area"))
	cSquare = EasyCrm.Searchcode(Request.Form("Squares"))
	cStart = EasyCrm.Searchcode(Request.Form("Start"))
	cType = EasyCrm.Searchcode(Request.Form("Type"))
	cSource = EasyCrm.Searchcode(Request.Form("Source"))
	cTrade = EasyCrm.Searchcode(Request.Form("Trade"))
	cStrade = EasyCrm.Searchcode(Request.Form("Strades"))
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
	cRecords = EasyCrm.Searchcode(Request.Form("Records"))
	cOrder = EasyCrm.Searchcode(Request.Form("Order"))
	cHetong = EasyCrm.Searchcode(Request.Form("Hetong"))
	cService = EasyCrm.Searchcode(Request.Form("Service"))
	cFile = EasyCrm.Searchcode(Request.Form("File"))
	
	Session("Search_Client_cCompany") = EasyCrm.Searchcode(Request.Form("company"))
	Session("Search_Client_cLinkman") = EasyCrm.Searchcode(Request.Form("Linkman"))
	Session("Search_Client_cMobile") = EasyCrm.Searchcode(Request.Form("Mobile"))
	Session("Search_Client_cAddress") = EasyCrm.Searchcode(Request.Form("address"))
	Session("Search_Client_cTel") = EasyCrm.Searchcode(Request.Form("Tel"))
	Session("Search_Client_cArea") = EasyCrm.Searchcode(Request.Form("Area"))
	Session("Search_Client_cSquare") = EasyCrm.Searchcode(Request.Form("Squares"))
	Session("Search_Client_cStart") = EasyCrm.Searchcode(Request.Form("Start"))
	Session("Search_Client_cType") = EasyCrm.Searchcode(Request.Form("Type"))
	Session("Search_Client_cSource") = EasyCrm.Searchcode(Request.Form("Source"))
	Session("Search_Client_cTrade") = EasyCrm.Searchcode(Request.Form("Trade"))
	Session("Search_Client_cStrade") = EasyCrm.Searchcode(Request.Form("Strades"))
	Session("Search_Client_cUser") = EasyCrm.Searchcode(Request.Form("User"))
	Session("Search_Client_cGroup") = EasyCrm.Searchcode(Request.Form("group"))
	Session("Search_Client_cTimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Client_cTimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	Session("Search_Client_cETimeBegin") = EasyCrm.Searchcode(Request.Form("ETimeBegin"))
	Session("Search_Client_cETimeEnd") = EasyCrm.Searchcode(Request.Form("ETimeEnd"))
	
	if EasyCrm.Searchcode(Request.Form("SearchRange")) <> "" then
	Session("Search_Client_SearchRange") = EasyCrm.Searchcode(Request.Form("SearchRange"))
	end if
	
	Session("Search_Client_Records") = EasyCrm.Searchcode(Request.Form("Records"))
	Session("Search_Client_Order") = EasyCrm.Searchcode(Request.Form("Order"))
	Session("Search_Client_Hetong") = EasyCrm.Searchcode(Request.Form("Hetong"))
	Session("Search_Client_Service") = EasyCrm.Searchcode(Request.Form("Service"))
	Session("Search_Client_File") = EasyCrm.Searchcode(Request.Form("File"))
	
	Dim sql
    sql = ""
	
	cCompanyWhere = EasyCrm.seachKey("cCompany",cCompany)
	
    If cCompany <> "" Then
        sql = sql & cCompanyWhere
	End If	
	
   If cLinkman <> "" Then
        sql = sql & " And cId in ( select cId from [Client]  where cLinkman  like '%" & cLinkman & "%' )"
	End If
    If cMobile <> "" Then
        sql = sql & " And cId in ( select cId from [Client]  where cMobile  like '%" & cMobile & "%' )"
	End If
	
    If cAddress <> "" Then
	    sql = sql & " And cAddress Like '%" & cAddress & "%'"
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
	
    If cTrade <> "" Then
	    sql = sql & " And cTrade = '" & cTrade & "'"
	End If
	
    If cStrade <> "" Then
	    sql = sql & " And cStrade = '" & cStrade & "'"
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
	
	If cRecords = "1" Then sql = sql & " And cID in ( select cId from [Records] ) "
	If cOrder = "1" Then sql = sql & " And cID in ( select cId from [Order] ) "
	If cHetong = "1" Then sql = sql & " And cID in ( select cId from [Hetong] ) "
	If cService = "1" Then sql = sql & " And cID in ( select cId from [Service] ) "
	If cFile = "1" Then sql = sql & " And cID in ( select cId from [File] ) "
	
	'If Session("CRM_level") < 9 Then
		'if MyShare <> "" then 
			'sql = sql & " and cShare = 1 and cUser = '"&Session("CRM_Name")&"' "
		'end if
		'if HeShare <> "" then 
			'sql = sql & " and cShare = 1 and cShareRange like '%"&Session("CRM_Name")&"%' "
		'end if
		
		'if Session("Search_Client_SearchRange") = "0" then
		'sql = sql & " "
		'else
			'if MyShare = "" or HeShare = "" then 
		'sql = sql & " And cUser In (" & arrUser & ")"
			'end if
		'end if
	'End If
	
End If

If cCompany = "" And cLinkman = ""  And cMobile = "" And cTel = "" And cAddress = "" And cArea = "" And cSquare = ""  And cType = "" And cTrade = "" And cStrade = "" And cSource = "" And cUser = "" And cGroup = "" And cstart = "" And cTimeBegin = "" And cTimeEnd = "" And cETimeBegin = "" And cETimeEnd = "" And cRecords = "" And cOrder = "" And cHetong = "" And cService = "" And cFile = "" And SearchRange = "" Then
    If Session("CRM_Search") <> "" Then
        sql = Session("CRM_Search")
	Else
	    'If Session("CRM_level") < 9 Then
			'if MyShare <> "" then 
				'sql = " and cShare = 1 and cUser = '"&Session("CRM_Name")&"' "
			'elseif HeShare <> "" then 
				'sql = " and cShare = 1 and cShareRange like '%"&Session("CRM_Name")&"%' "
			'else
				'sql = " And cUser In (" & arrUser & ")"
			'end if
		'End If
		sql = ""
	End If
Else
    Session("CRM_Search") = sql
End If

If subAction = "killSession" Then
	Session("CRM_Search") = ""
	Session("Search_Client_cCompany") = ""
	Session("Search_Client_cLinkman") = ""
	Session("Search_Client_cMobile") = ""
	Session("Search_Client_cAddress") = ""
	Session("Search_Client_cTel") = ""
	Session("Search_Client_cArea") = ""
	Session("Search_Client_cSquare") = ""
	Session("Search_Client_cStart") = ""
	Session("Search_Client_cType") = ""
	Session("Search_Client_cSource") = ""
	Session("Search_Client_cTrade") = ""
	Session("Search_Client_cStrade") = ""
	Session("Search_Client_cUser") = ""
	Session("Search_Client_cGroup") = ""
	Session("Search_Client_cTimeBegin") = ""
	Session("Search_Client_cTimeEnd") = ""
	Session("Search_Client_cETimeBegin") = ""
	Session("Search_Client_cETimeEnd") = ""
	Session("Search_Client_SearchRange") = ""
	Session("Search_Client_Records") = ""
	Session("Search_Client_Order") = ""
	Session("Search_Client_Hetong") = ""
	Session("Search_Client_Service") = ""
	Session("Search_Client_File") = ""
	Session("Search_Client_NewUser") = ""
	'If Session("CRM_level") < 9 Then
			'if MyShare <> "" then 
				'sql = " and cShare = 1 and cUser = '"&Session("CRM_Name")&"' "
			'elseif HeShare <> "" then 
				'sql = " and cShare = 1 and cShareRange like '%"&Session("CRM_Name")&"%' "
			'else
				'sql = " And cUser In (" & arrUser & ")"
			'end if
	'else
	'sql=""
	'end if
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
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Company%> > <%=L_Page_Listall%></td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
			<% If mid(Session("CRM_qx"), 5, 1) = 1 Then %>
			<input type="button" class="button_top_set" value=" " title="设置" onclick='Setting_ListAll()' style="cursor:pointer" />
			<%end if%>
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
					<% If mid(Session("CRM_qx"), 10, 1) = 1 Then %>
					<li <%if otype="MyShare" then%>class="hover"<%end if%> id="CheckD"><span><a href="?otype=MyShare&MyShare=1">我的共享</a></span></li>
					<li <%if otype="HeShare" then%>class="hover"<%end if%> id="CheckE"><span><a href="?otype=HeShare&HeShare=1">共享给我</a></span></li>
					<% end If %>
					<% If mid(Session("CRM_qx"), 17, 1) = 1 Then %>
					<li class="" id="CheckC"><span><a href="#" onclick='Client_InfoAdd()' style="cursor:pointer">新增客户</a></span></li>
					<% end If %>
				</ul>
			</div>
		</td>
	</tr>
</table>
<script>function Setting_ListAll() {$.dialog.open('../system/GetUpdate.asp?action=Setting&sType=ListAll', {title: '自定义设置', width: 900, height: 480,fixed: true}); };</script>
<%
if CLng(ABS(EasyCrm.getNewItem("User","uName","'"&Session("CRM_name")&"'","uClientNum"))) > 0 and CLng(ABS(EasyCrm.getCountItem("Client","cid","cidstr"," and cYn=1 and cUser='"&Session("CRM_name")&"' "))) >= CLng(ABS(EasyCrm.getNewItem("User","uName","'"&Session("CRM_name")&"'","uClientNum"))) then
%>
<script>function Client_InfoAdd() {art.dialog({title: 'Error',time: 1,icon: 'warning',content: '客户量达到最大配额！'});};</script>
<%
else
%>
<script>function Client_InfoAdd() {$.dialog.open('GetUpdate.asp?action=Client&sType=Add', {title: '新增', width: 900, height: 480,fixed: true}); };</script>
<%
end if
%>
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
									<td class="td_l_r title" style="border-top:0;"><%=L_Client_cCompany%></td>
									<td class="td_r_l" style="border-top:0;"><input name="Company" type="text" class="int" id="Company" size="30" value="<%=Session("Search_Client_cCompany")%>" ></td>
									<td class="td_l_r title" style="border-top:0;"><%=L_Client_cDate%></td>
									<td class="td_r_l" style="border-top:0;"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Client_cTimeBegin")%>" /> ~ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Client_cTimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cLinkman%></td>
									<td class="td_r_l"><input name="Linkman" type="text" class="int" id="Linkman" size="30" value="<%=Session("Search_Client_cLinkman")%>" ></td>
									<td class="td_l_r title"><%=L_Client_cLastUpdated%></td>
									<td class="td_r_l"><input name="ETimeBegin" type="text" maxlength="10" id="ETimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Client_cETimeBegin")%>" /> ~ <input name="ETimeEnd" type="text" maxlength="10" id="ETimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Client_cETimeEnd")%>" /> </td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cMobile%></td>
									<td class="td_r_l"><input name="Mobile" type="text" class="int" id="Mobile" size="30" value="<%=Session("Search_Client_cMobile")%>" ></td>
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
									<td class="td_l_r title"><%=L_Client_cTel%></td>
									<td class="td_r_l"><input name="Tel" type="text" class="int" id="Tel" size="30" value="<%=Session("Search_Client_cTel")%>"></td>
									<td class="td_l_r title"><%=L_Client_cAddress%></td>
									<td class="td_r_l"><input name="Address" type="text" class="int" id="Address" size="30" value="<%=Session("Search_Client_cAddress")%>" ></td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cType%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Type","Type",Session("Search_Client_cType")) %></td>
									<td class="td_l_r title"><%=L_Client_cTrade%></td>
									<td class="td_r_l">
										<select name="Trade" onchange="getTrade(this.options[this.selectedIndex].id);">
										<option value=""><%=L_Please_choose_01%></option>
										<% 
											Set rsb = Conn.Execute("select * from ProductClass where pClassFid = '0' ")
											If Not rsb.Eof then
											Do While Not rsb.Eof
											pClassid= rsb("pClassid")
											pClassname= rsb("pClassname")
										%>
											<option value="<%=pClassname%>" id="<%=pClassid%>"><%=pClassname%></option>
										<%
											rsb.Movenext
											Loop
											End If
											rsb.Close
											Set rsb = Nothing 
										%>
										</select> 
										<span id="Stradediv"  style="margin-left:10px;padding:0;">
											<select name="Strades">
												<option value=""><%=L_Please_choose_02%></option>
												<% 
												IF ""&cTrade&""<>"" then
												Set rsb = Conn.Execute("select * from ProductClass where pClassFid = '"&EasyCrm.getNewItem("ProductClass","pClassname","'"&cTrade&"'","pClassId")&"' ")
												If Not rsb.Eof then
												Do While Not rsb.Eof
												pClassname= rsb("pClassname")
												%>
												<option value="<%=pClassname%>"><%=pClassname%></option>
												<%rsb.Movenext
												Loop
												End If
												rsb.Close
												Set rsb = Nothing 
												end if
												%>
											</select>
										</span>
									</td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cStart%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Star","Start",Session("Search_Client_cStart")) %></td>
								
									<td class="td_l_r title"><%=L_Client_cUser%></td>
									<td class="td_r_l">
										<% If Session("CRM_level") = 9 Then %>
										<% = EasyCrm.UserList(2,"User",Session("Search_Client_cUser")) %>　<%=L_Client_cGroup%>：<% = EasyCrm.getList(2,"system_group","gId","gName","Group",""&viewGroup&"") %>
										<% Else %>
										<% = EasyCrm.UserList(1,"User",Session("Search_Client_cUser")) %>
										<% End If %>
									</td>
								</tr>
								<tr>
									<td class="td_l_r title"><%=L_Client_cSource%></td>
									<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Source","Source",Session("Search_Client_cSource")) %></td>
									<td class="td_l_r title">查询范围</td>
									<td class="td_r_l"> 
										<input type="radio" name="SearchRange" value="1" <%if Session("Search_Client_SearchRange") = "1" or Session("Search_Client_SearchRange") = "" then%>checked<%end if%>> 权限内　
										<input type="radio" name="SearchRange" value="0" <%if Session("Search_Client_SearchRange") = "0" then%>checked<%end if%>> 所有（包括公海）
									</td>
								</tr>
								<tr>
									<td class="td_l_r title">其它</td>
									<td class="td_r_l" colspan=3> 
										<input type="checkbox" name="Records" value="1" <%if Session("Search_Client_Records") = "1" then%>checked<%end if%>> 有跟单　
										<input type="checkbox" name="Order" value="1" <%if Session("Search_Client_Order") = "1" then%>checked<%end if%>> 有订单　
										<input type="checkbox" name="Hetong" value="1" <%if Session("Search_Client_Hetong") = "1" then%>checked<%end if%>> 有合同　
										<input type="checkbox" name="Service" value="1" <%if Session("Search_Client_Service") = "1" then%>checked<%end if%>> 有售后　
										<input type="checkbox" name="File" value="1" <%if Session("Search_Client_File") = "1" then%>checked<%end if%> > 有附件　
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
								<% If mid(Session("CRM_qx"), 9, 1) = 1 Then %>
									<span class="tips01" style="float:left;display:none;padding:0 10px;height:34px;text-align:left;position:fixed;right:10px;top:43px;color:#000;" id="CheckSub">
									<% If mid(Session("CRM_qx"), 12, 1) = 1 Then %>
										新：<% If Session("CRM_level") = 9 Then %><% = EasyCrm.UserList(2,"NewUser","") %><%else%><% = EasyCrm.UserList(1,"NewUser","") %><%end if%>
										<input type="submit" name="Checkexecute" id="Checkexecute" class="button224" value="<%=L_Transfer%>">
										<%if Session("Search_Client_cUser") <> "" then%>
										<input type="submit" name="Checkexecute" id="Checkexecute" class="button247" value="<%=L_Transfer_all%>">
										<%end if%>
									<%end if%>
									<% If mid(Session("CRM_qx"), 19, 1) = 1 Then %>
										<input type="submit" name="Checkexecute" id="Checkexecute" class="button227" value="<%=L_Del%>">
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
								<%if Client_cDate = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cDate%></td>
								<%end if%>
								<%if Client_cCompany = 1 then%>
								<td class="td_l_l"><%=L_Client_cCompany%></td>
								<%end if%>
								<%if Client_cArea = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cArea%></td>
								<%end if%>
								<%if Client_cSquare = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cSquare%></td>
								<%end if%>
								<%if Client_cAddress = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cAddress%></td>
								<%end if%>
								<%if Client_cType = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cType%></td>
								<%end if%>
								<%if Client_cTel = 1 then%>
								<td width="120" class="td_l_c"><%=L_Client_cTel%></td>
								<%end if%>
								<%if Client_cFax = 1 then%>
								<td width="120" class="td_l_c"><%=L_Client_cFax%></td>
								<%end if%>
								<%if Client_cTrade = 1 then%>
								<td width="80" class="td_l_c"><%=L_Client_cTrade%></td>
								<%end if%>
								<%if Client_cStrade = 1 then%>
								<td width="80" class="td_l_c"><%=L_Client_cStrade%></td>
								<%end if%>
								<%if Client_cStart = 1 then%>
								<td width="80" class="td_l_c"><%=L_Client_cStart%></td>
								<%end if%>
								<%if Client_cSource = 1 then%>
								<td width="80" class="td_l_c"><%=L_Client_cSource%></td>
								<%end if%>
								<%if Client_cLinkman = 1 then%>
								<td width="80" class="td_l_c"><%=L_Client_cLinkman%></td>
								<%end if%>
								<%if Client_cZhiwei = 1 then%>
								<td width="80" class="td_l_c"><%=L_Client_cZhiwei%></td>
								<%end if%>
								<%if Client_cMobile = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cMobile%></td>
								<%end if%>
                                
                               <%							
Set rss = Server.CreateObject("ADODB.Recordset")
rss.Open "Select * From [CustomField] where cTable='Client' and cList = '1' order by Id asc ",conn,3,1
If rss.RecordCount > 0 Then
Do While Not rss.BOF And Not rss.EOF
%>
	<td width="100" class="td_l_c"><%=rss("cTitle")%></td>
<%
rss.MoveNext
Loop
end if
rss.Close
Set rss = Nothing
%>
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
								<%if Client_cLastUpdated = 1 then%>
								<td width="80" class="td_l_c"><%=L_Client_cLastUpdated%></td>
								<%end if%>
								<%if Client_cRNextTime = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cRNextTime%></td>
								<%end if%>
								<%if Client_cOEDate = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cOEDate%></td>
								<%end if%>
								<%if Client_cHEdate = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cHEdate%></td>
								<%end if%>
								<%if Client_cHMoney = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cHMoney%></td>
								<%end if%>
								<%if Client_cHOwed = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cHOwed%></td>
								<%end if%>
								<%if Client_cSNum = 1 then%>
								<td width="80" class="td_l_c"><%=L_Client_cSNum%></td>
								<%end if%>
								<%if Client_cUser = 1 then%>
								<td width="100" class="td_l_c"><%=L_Client_cUser%></td>
								<%end if%>
								<td width="100" class="td_l_c"><%=L_Top_Manage%></td>
							</tr>
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							
						if Session("Search_Client_SearchRange") = "0" then
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [client] where 1 = 1 "&sql&" "&Share&" Order By cId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [client] where 1 = 1 "&sql&" "&Share&" and cid < ( SELECT Min(cid) FROM ( SELECT TOP "&pagenum&" cid FROM [client]  where 1 = 1 "&sql&" "&Share&" ORDER BY cId desc ) AS T ) Order By cId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(cid) As RecordSum From [client] where 1 = 1 "&sql&" "&Share&" ",1,1)
						else
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [client] where cYn = 1 "&sql&" "&Share&" Order By cId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [client] where cYn = 1 "&sql&" "&Share&" and cid < ( SELECT Min(cid) FROM ( SELECT TOP "&pagenum&" cid FROM [client]  where cYn = 1 "&sql&" "&Share&" ORDER BY cId desc ) AS T ) Order By cId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(cid) As RecordSum From [client] where cYn = 1 "&sql&" "&Share&" ",1,1)
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
								<%if Client_cDate = 1 then%>
								<td class="td_l_c"><%=rs("cDate")%></td>
								<%end if%>
								
						    	
								<td class="td_l_l">
							
									<a onclick='Client_InfoView<%=rs("cId")%>()' style="cursor:pointer"><%=rs("cCompany")%></a>
								</td>
							
								<%if Client_cArea = 1 then%>
								<td class="td_l_c"><%=rs("cArea")%></td>
								<%end if%>
								<%if Client_cSquare = 1 then%>
								<td class="td_l_c"><%=rs("cSquare")%></td>
								<%end if%>
								<%if Client_cAddress = 1 then%>
								<td class="td_l_c" onmouseover="tip.start(this)" tips="<%=rs("cAddress")%>"><%=left(rs("cAddress"),5)%></td>
								<%end if%>
								<%if Client_cType = 1 then%>
								<td class="td_l_c"><%=rs("cType")%></td>
								<%end if%>
								<%if Client_cTel = 1 then%>
								<td class="td_l_c"><%=rs("cTel")%></td>
								<%end if%>
								<%if Client_cFax = 1 then%>
								<td class="td_l_c"><%=rs("cFax")%></td>
								<%end if%>
								<%if Client_cTrade = 1 then%>
								<td class="td_l_c"><%=rs("cTrade")%></td>
								<%end if%>
								<%if Client_cStrade = 1 then%>
								<td class="td_l_c"><%=rs("cStrade")%></td>
								<%end if%>
								<%if Client_cStart = 1 then%>
								<td class="td_l_c"><%=rs("cStart")%></td>
								<%end if%>
								<%if Client_cSource = 1 then%>
								<td class="td_l_c"><%=rs("cSource")%></td>
								<%end if%>
								<%if Client_cLinkman = 1 then%>
								<td class="td_l_c"><%=rs("cLinkman")%></td>
								<%end if%>
								<%if Client_cZhiwei = 1 then%>
								<td class="td_l_c"><%=rs("cZhiwei")%></td>
								<%end if%>
								<%if Client_cMobile = 1 then%>
								<td class="td_l_c"><%=rs("cMobile")%></td>
								<%end if%>




<%
								cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID",""&rs("cId")&" ","cContent")
								cContentArr = split(cContentStr,"|")	
															
								Set rss1 = Server.CreateObject("ADODB.Recordset")
								rss1.Open "Select * From [CustomField] where cTable = 'Client' and cList = '1' order by id asc ",conn,1,1
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






























								<%if Client_cLastUpdated = 1 then%>
								<td class="td_l_c"><font color=red><%=Datediff("d",rs("cLastUpdated"),Now())%> </font>天前</td>
								<%end if%>
								<%if Client_cRNextTime = 1 then%>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("cRNextTime"),2)%></td>
								<%end if%>
								<%if Client_cOEDate = 1 then%>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("cOEDate"),2)%></td>
								<%end if%>
								<%if Client_cHEdate = 1 then%>
								<td class="td_l_c"><%=EasyCrm.FormatDate(rs("cHEdate"),2)%></td>
								<%end if%>
								<%if Client_cHMoney = 1 then%>
								<td class="td_l_c"><font color=red><%=EasyCrm.getSUMItem("Hetong","hMoney","hMoneystr"," and cid = "&rs("cID")&" ")%> </font> 元</td>
								<%end if%>
								<%if Client_cHOwed = 1 then%>
								<td class="td_l_c"><font color=red><%=EasyCrm.getSUMItem("Hetong","hOwed","hOwedstr"," and cid = "&rs("cID")&" ")%> </font> 元</td>
								<%end if%>
								<%if Client_cSNum = 1 then%>
								<td class="td_l_c"><font color=red><%=EasyCrm.getCountItem("Service","sID","sIDstr"," and cid = "&rs("cID")&" ")%> </font> 次</td>
								<%end if%>
								<%if Client_cUser = 1 then%>
								<td class="td_l_c"><%=rs("cUser")%></td>
								<%end if%>
						
								<td class="td_l_c">
									<input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Client_InfoEdit<%=rs("cId")%>()' style="cursor:pointer" />
									<% if inStr(""&arrUser&"",rs("cUser"))>0 then %>
									<% If mid(Session("CRM_qx"), 19, 1) = 1 Then %>
									<input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Client_InfoDel<%=rs("cId")%>()' style="cursor:pointer" />
									<%end if%>
									<%end if%>
								</td>
							
					
							
							</tr>
							<script>function Client_InfoView<%=rs("cId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&cId=<%=rs("cId")%><%if rs("cYn") = 0 then%>&YNRange=0<%end if%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
							<script>function Client_InfoEdit<%=rs("cId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoEdit&cId=<%=rs("cId")%>', {title: '编辑', width: 900,height: 480, fixed: true}); };</script>
							
							
							
							<script>function Client_InfoDel<%=rs("cId")%>()
							{
								art.dialog({
									content: '<%=Alert_del_YN%>',
									icon: 'error',
									ok: function () {
										<%if YnDelReason = 1 then%> 
										$.dialog.open('GetUpdate.asp?action=Client&sType=DelReason&cId=<%=rs("cId")%>', 
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
			<%if otype="MyShare" then%>
			<%=EasyCrm.pagelist("ListAll.asp?otype=MyShare&MyShare=1", PN,TotalPages,TotalRecords)%>
			<%elseif otype="HeShare" then%>
			<%=EasyCrm.pagelist("ListAll.asp?otype=HeShare&HeShare=1", PN,TotalPages,TotalRecords)%>
			<%else%>
			<%=EasyCrm.pagelist("ListAll.asp", PN,TotalPages,TotalRecords)%>
			<%end if%>
		</td>
	</tr>
</table>
</div>

<script language="JavaScript">
<!--
for(var i=0;i<document.getElementById('Area').options.length;i++){
    if(document.getElementById('Area').options[i].value == "<% = Session("Search_Client_cArea") %>"){
    document.getElementById('Area').options[i].selected = true;}}

	
for(var i=0;i<document.getElementById('Squares').options.length;i++){
    if(document.getElementById('Squares').options[i].value == "<% = Session("Search_Client_cSquare") %>"){
    document.getElementById('Squares').options[i].selected = true;}}

for(var i=0;i<document.getElementById('Trade').options.length;i++){
    if(document.getElementById('Trade').options[i].value == "<% = Session("Search_Client_cTrade") %>"){
    document.getElementById('Trade').options[i].selected = true;}}

for(var i=0;i<document.getElementById('Strades').options.length;i++){
    if(document.getElementById('Strades').options[i].value == "<% = Session("Search_Client_cStrade") %>"){
    document.getElementById('Strades').options[i].selected = true;}}

for(var i=0;i<document.getElementById('Group').options.length;i++){
    if(document.getElementById('Group').options[i].value == "<% = Session("Search_Client_cGroup") %>"){
    document.getElementById('Group').options[i].selected = true;}}
-->
</script>
<%
end Sub 

Sub CheckSubject()
cId=Trim(Request("cId"))
NewUser=Request("NewUser")

PN = CLng(ABS(Request("PN")))
If Request("Checkexecute")="转移" Then
	If cId="" Then
		Response.Write "<script>alert("""&alert_trans_no_select&""");</script><SCRIPT LANGUAGE=JavaScript>document.Search.chkall.focus()</SCRIPT>"
		Response.End
	elseif NewUser="" then
		Response.Write "<script>alert("""&alert_trans_no_newuser&""");</script><SCRIPT LANGUAGE=JavaScript>document.Search.NewUser.focus()</SCRIPT>"
		Response.End
	else
		cidarr = split(cid,",")
		' 限制客户量，则循环判断是否已经达到最大客户量
		if CLng(ABS(EasyCrm.getNewItem("User","uName","'"&NewUser&"'","uClientNum"))) > 0 then
		for i = 0 to Ubound(cidarr) 
		
			if CLng(ABS(EasyCrm.getCountItem("Client","cid","cidstr"," and cYn=1 and cUser='"&NewUser&"' "))) < CLng(ABS(EasyCrm.getNewItem("User","uName","'"&NewUser&"'","uClientNum"))) then
				set rs=conn.execute("update [Client] set cUser = '"&NewUser&"' where cId ='" & cidarr(i) & "' ")
				if SaveOldUser = 0 then
				set rs=conn.execute("update [Linkmans] set lUser = '"&NewUser&"' where cId ='" & cidarr(i) & "' ")
				set rs=conn.execute("update [Records] set rUser = '"&NewUser&"' where cId ='" & cidarr(i) & "' ")
				set rs=conn.execute("update [Order] set oUser = '"&NewUser&"' where cId ='" & cidarr(i) & "' ")
				set rs=conn.execute("update [Hetong] set hUser = '"&NewUser&"' where cId ='" & cidarr(i) & "' ")
				set rs=conn.execute("update [Service] set sUser = '"&NewUser&"' where cId ='" & cidarr(i) & "' ")
				set rs=conn.execute("update [Expense] set eUser = '"&NewUser&"' where cId ='" & cidarr(i) & "' ")
				set rs=conn.execute("update [File] set fUser = '"&NewUser&"' where cId ='" & cidarr(i) & "' ")
				end if
			else
				if ""&YNalert&"" = 1 then
				Response.Write "<script>alert(""成功转移"&i&"条记录，"&NewUser&"的客户量达到最大配额！"");</script>"
				end if
			end if
		next
				Response.Write("<script>location.href='?PN="&PN&"' ;</script>")
				Response.end
		
		else
				set rs=conn.execute("update [Client] set cUser = '"&NewUser&"' where cId In(" & cId & ")")
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
elseIf Request("Checkexecute")="转移所有" Then
	if NewUser="" then
		Response.Write "<script>alert("""&alert_trans_no_newuser&""");history.back(1);</script><SCRIPT LANGUAGE=JavaScript>document.Search.NewUser.focus()</SCRIPT>"
		Response.End
	else
		set rs=conn.execute("update Client set cUser = '"&NewUser&"' where cUser ='"&Session("Search_Client_cUser")&"' ")
		Response.Write("<script>alert('"&alert4&"');</script>")
		Response.Write("<script>location.href='?PN="&PN&"' ;</script>")
		Response.end
	end if
elseIf Request("Checkexecute")="删除" Then

	cidarr = split(cid,",")
	for i = 0 to Ubound(cidarr) 
	'循环删除，并写入操作记录
	conn.execute("update Client set cYn = 0 where cId = " & cidarr(i) & " ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cidarr(i)&"','"&L_Client&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")
	next
	if ""&YNalert&"" = 1 then
		Response.Write("<script>alert('"&alert4&"');</script>")
	end if
	Response.Write("<script>location.href='?PN="&PN&"' ;</script>")
End If
   
End Sub

Sub deleteData()
    Dim cid
	cid = Trim(Request("cid"))
	If cid = "" Then
	Exit Sub
	End If
	conn.execute("update Client set cYn = 0 where cId = "&cId&" ")
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub
%>
<script src="../data/calendar/WdatePicker.js"></script>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html>
<% Set EasyCrm = nothing %>