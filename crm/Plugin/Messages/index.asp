<!--#include file="../../data/conn.asp" -->
<!--#include file="config.asp" --><!--#include file="../../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
Dim mconn,mMDBPath
set mrs=server.CreateObject("adodb.recordset")
Set mconn = Server.CreateObject("ADODB.Connection")
mMDBPath = Server.MapPath("blackdict.mdb")
mconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mMDBPath

'获取当前页码
PNN = Trim(Request.QueryString("PN"))
subAction = Trim(Request("subAction"))
if PNN="" then PNN=1 
otype	=	Request.QueryString("otype")
if otype="" then otype="Main"
Dim Articles,ContactMsg
if Request.QueryString("cMobile")<>"" then
Session("CRM_Message_Mobile") = Request.QueryString("cMobile")
end if
Articles = Session("CRM_Message_Mobile")
		
ContactMsg = Request.QueryString("ContactMsg")
if Articles<>"" then
	'Mobiles=right(Articles,(clng(len(Articles))-instr(1,Articles,",",1)))
	Mobiles=left(Articles,len(Articles)-1)
	if Mobiles <>"" then
	mMobile=left(Mobiles,len(Mobiles))
	end if
end if
if ContactMsg<>"" then
ContactMsg = left(ContactMsg,len(ContactMsg)-1)
end if
if ContactMsg <>"" and Articles="" then 
mMobile = ""&ContactMsg&""
end if

Function replacemobile(str)
	str = Replace(str,",",",")
	str = Replace(str,".",",")
	str = Replace(str,";",",")
	str = Replace(str,"；",",")
	str = Replace(str,"，",",")
	str = Replace(str,"。",",")
	str = Replace(str,"/",",")
	str = Replace(str,"、",",")
	str = Replace(str,"|",",")
	str = Replace(str,"｜",",")
	str = Replace(str,chr(9),",")
	str = Replace(str,chr(10),",")
	str = Replace(str,chr(13),",")
	replacemobile = str
End Function

Function lastMessages(cMobile)
    Dim rs
	cMobile = Replace(cMobile," ","")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select top 1 mTime From Plugin_Messages Where mPhonenum like '%"&cMobile&"%' order by mid desc",conn,1,1
	If rs.RecordCount = 1 Then
	    lastMessages = "<font color=red>"&Datediff("d",rs("mTime"),now())&"</font> 天前"
	Else
	    lastMessages = "未发过"
	End If
    rs.Close
	Set rs = Nothing
End Function

Function lastMessagesinfo(cMobile)
    Dim rs
	cMobile = Replace(cMobile," ","")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select mContent,mTime From [Plugin_Messages] Where mPhonenum like '%"&cMobile&"%' order by mid desc ",conn,1,1
	dim i
	i=0
Do While Not rs.BOF And Not rs.EOF
i = i + 1
	    lastMessagesinfo = lastMessagesinfo & "<div><ol>"
	    lastMessagesinfo = lastMessagesinfo & "<li>"&i&"："&rs(0)&" <font color=red>"&rs(1)&"</font></li>"
	    lastMessagesinfo = lastMessagesinfo & "</ol></div>"
	rs.MoveNext
Loop
    rs.Close
	Set rs = Nothing
End Function

	Dim Sql
	sql=""
	If subAction = "searchItem" Then
    Dim cArea,cSquare,cStart,cType,cSource,cTimeBegin,cTimeEnd,cUser,cCompany,SHYN
	cArea = EasyCrm.Searchcode(Request("area"))
	cSquare = EasyCrm.Searchcode(Request("Squares"))
	cStart = EasyCrm.Searchcode(Request("Start"))
	cType = EasyCrm.Searchcode(Request("Type"))
	cSource = EasyCrm.Searchcode(Request("Source"))
	cUser = EasyCrm.Searchcode(Request("User"))
	cCompany = EasyCrm.Searchcode(Request("cCompany"))
	cTrade = EasyCrm.Searchcode(Request("Trade"))
	cStrade = EasyCrm.Searchcode(Request("Strades"))
	
	Session("Search_Plugin_Messages_cCompany") = EasyCrm.Searchcode(Request("cCompany"))
	Session("Search_Plugin_Messages_cArea") = EasyCrm.Searchcode(Request("area"))
	Session("Search_Plugin_Messages_cSquare") = EasyCrm.Searchcode(Request("Squares"))
	Session("Search_Plugin_Messages_cStart") = EasyCrm.Searchcode(Request("Start"))
	Session("Search_Plugin_Messages_cType") = EasyCrm.Searchcode(Request("Type"))
	Session("Search_Plugin_Messages_cSource") = EasyCrm.Searchcode(Request("Source"))
	Session("Search_Plugin_Messages_cUser") = EasyCrm.Searchcode(Request("User"))
	Session("Search_Plugin_Messages_cTrade") = EasyCrm.Searchcode(Request("Trade"))
	Session("Search_Plugin_Messages_cStrade") = EasyCrm.Searchcode(Request("Strades"))
		
    If cArea <> "" Then
        sql = sql & " And cArea = '" & cArea & "' "
	End If
    If cSquare <> "" Then
        sql = sql & " And cSquare = '" & cSquare & "' "
	End If
    If cStart <> "" Then
        sql = sql & " And cStart = '" & cStart & "' "
	End If
    If cType <> "" Then
        sql = sql & " And cType = '" & cType & "' "
	End If
    If cSource <> "" Then
        sql = sql & " And cSource = '" & cSource & "' "
	End If
    If cUser <> "" Then
        sql = sql & " And cUser = '" & cUser & "' "
	End If
    If cCompany <> "" Then
        sql = sql & " And cCompany like '%" & cCompany & "%' "
	End If
	
    If cTrade <> "" Then
	    sql = sql & " And cTrade = '" & cTrade & "' "
	End If
	
    If cStrade <> "" Then
	    sql = sql & " And cStrade = '" & cStrade & "' "
	End If
	
	End If
	
	sql= sql & " And cMobile <>'' and len(cMobile)=11 and left(cMobile,1)=1"
	
	If Session("CRM_level") < 9 Then
		sql = sql &  " And cUser In (" & arrUser & ") " 
	End If
	
	If cArea = "" And cSquare = ""  And cType = "" And cSource = "" And cUser = "" And cstart = "" And cCompany="" Then
		If Session("Search_Plugin_Messages_Search") <> "" Then
			sql = Session("Search_Plugin_Messages_Search")
		End If
	Else
		Session("Search_Plugin_Messages_Search") = sql
	End If

	If subAction = "killSession" Then
		Session("Search_Plugin_Messages_Search") = ""
		Session("Search_Plugin_Messages_cArea") = ""
		Session("Search_Plugin_Messages_cSquare") = ""
		Session("Search_Plugin_Messages_cStart") = ""
		Session("Search_Plugin_Messages_cType") = ""
		Session("Search_Plugin_Messages_cSource") = ""
		Session("Search_Plugin_Messages_cUser") = ""
		Session("Search_Plugin_Messages_cCompany") = ""
		Session("Search_Plugin_Messages_cTrade") = ""
		Session("Search_Plugin_Messages_cStrade") = ""
		
		Session("CRM_Message_Mobile")=""
		
		sql=" And cMobile <>'' and len(cMobile)=11 and left(cMobile,1)=1 "
		
		If Session("CRM_level") < 9 Then
			sql = sql &  " And cUser In (" & arrUser & ") " 
		End If
	End If

	Dim intTotalRecords,intTotalPages,PN,intPageSize'记录总数，总页数，当前页，分页数量
	PN = CLng(ABS(Request("PN")))

    If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
    intPageSize = DataPageSize
	pageNums = intPageSize*(PN-1)

		Set rs = Server.CreateObject("ADODB.Recordset")
		IF PN=1 THEN
	    rs.Open "Select top "&intPageSize&" * From [Client] where cYn=1 "&sql&" Order By cId desc ",conn,1,1 
		ELSE
	    rs.Open "Select top "&intPageSize&" * From [Client] where cYn=1 "&sql&" and cId < ( SELECT Min(cId) FROM ( SELECT TOP "&pageNums&" cId FROM [Client] where  cYn=1 "&sql&" ORDER BY cId desc ) AS T ) Order By cId desc ",conn,1,1
		END IF
		SQLstr="Select count(cId) As RecordSum From [Client] where cYn=1 "&sql&" " '统计页码

	Dim TotalRecords,TotalPages
	Set Rsstr=conn.Execute(SQLstr,1,1) 
	TotalRecords=Rsstr("RecordSum") 
	if Int(TotalRecords/DataPageSize)=TotalRecords/DataPageSize then
	TotalPages=TotalRecords/DataPageSize
	else
	TotalPages=Int(TotalRecords/DataPageSize)+1
	end if
	Rsstr.Close 
	Set Rsstr=Nothing

    If PN > TotalPages Then PN = TotalPages

'翻页代码开始
	
	 strCounter = strCounter & " "&EasyCrm.pagelist("index.asp", PN,TotalPages,TotalRecords)&""
	
'翻页代码结束

Dim i
i = 0
Do While Not rs.BOF And Not rs.EOF
    i = i + 1
	strToPrint = strToPrint & "			<tr class=""tr"">" & VBCrlf
	'if InStr(Articles,"," & rs("cMobile") & ",")>0 Then
	'strToPrint = strToPrint & "				<td class=""td_l_c""><input id=""Plugin_msg"" type=""checkbox"" name=""Plugin_msg"" onclick=""SetArticleId(this,"&rs("cMobile")&");"" checked value=""" & rs("cId") & """ /></td>" & VBCrlf
	'else
	'strToPrint = strToPrint & "				<td class=""td_l_c""><input id=""Plugin_msg"" type=""checkbox"" name=""Plugin_msg"" onclick=""SetArticleId(this,"&rs("cMobile")&");"" value=""" & rs("cId") & """ /></td>" & VBCrlf
	'end if
	if inStr(Session("CRM_Message_Mobile"),EasyCrm.getNewItem("Client","cID",""&rs("cID")&"","cMobile"))>0 then
	strToPrint = strToPrint & "        <td class=""td_l_c title""><input type=""checkbox"" name=""cId"" id=""cId"" value=""" & rs("cId") & """ onClick=""unselectall(this.form)"" checked></td>" & VBCrlf
	else
	strToPrint = strToPrint & "        <td class=""td_l_c title""><input type=""checkbox"" name=""cId"" id=""cId"" value=""" & rs("cId") & """ onClick=""unselectall(this.form)""></td>" & VBCrlf
	end if
		
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("cID") & "</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_l"" onclick='Client_InfoEdit"&rs("cID")&"()' style='cursor:pointer' >" & rs("cCompany")& "</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("cLinkman") & "</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("cMobile") & "</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c""><a onclick=""Showhiden(this,'box"&rs("cID")&"',false,'"&lastMessages(rs("cMobile"))&"','"&lastMessages(rs("cMobile"))&"')"" style=""cursor:pointer"">"&lastMessages(rs("cMobile"))&"</a> </td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c""> "&rs("cType")&"</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c""> "&rs("cStart")&"</td>" & VBCrlf
	strToPrint = strToPrint & "			</tr>" & VBCrlf
	strToPrint = strToPrint & "			<tr class=""tr"" id=""box"&rs("cID")&""" style=""display:none;"">" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_l"" colspan=8> "&lastMessagesinfo(rs("cMobile"))&"</td>" & VBCrlf
	strToPrint = strToPrint & "			</tr>" & VBCrlf
	strToPrint = strToPrint & "			<script>function Client_InfoEdit"&rs("cID")&"() {$.dialog.open('"&SiteUrl&"Main/GetUpdate.asp?action=Client&sType=InfoEdit&cId="&rs("cID")&"', {title: '编辑', width: 900,height: 480, fixed: true}); };</script>" & VBCrlf
    If i >= intPageSize Then Exit Do
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/modify.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>zDialog/zDrag.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>zDialog/zDialog.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>

<script language="javascript" src="Js/Ajax.js"></script>
<script type="text/javascript">
    function SetArticleId(o, i) {
      if (o.checked) {
        AddCookie(i)
      }
      else {
        RemoveCookie(i)
      }
    }
    function SetCookie(name, value) {
      document.cookie = name + "=" + escape(value);
    }
    function GetCookie(name) {
      if (document.cookie.length > 0) {
        c_start = document.cookie.indexOf(name + "=");
        if (c_start != -1) {
          c_start = c_start + name.length + 1;
          c_end = document.cookie.indexOf(";", c_start);
          if (c_end == -1) c_end = document.cookie.length;
          return unescape(document.cookie.substring(c_start, c_end));
        }
      }
      return "";
    }
    function AddCookie(i) {
      d = GetCookie("Plugin_msg");
      if (d == "") d = ",";
      if (d.indexOf("," + i + ",") == -1) {
        d += i + ",";
        SetCookie("Plugin_msg", d);
      }
    }

    function RemoveCookie(i) {
      d = GetCookie("Plugin_msg");
      var reg = new RegExp("\\," + i + "\\,");
      if (reg.test(d)) {
        d = d.replace(reg, ",");  
        SetCookie("Plugin_msg", d);
      }   	  
    }

	function ClearCookie() {
      d = GetCookie("Plugin_msg");
      d = d.replace(d, "");  
        SetCookie("Plugin_msg", d);  	  
    }
  </script>
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

<body style="padding-top:35px;">

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：功能插件 > 短信群发</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<%if inStr(Plugin_Messages_manage,session("CRM_name"))>0 or Session("CRM_level") = 9 then%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10 pdb10">   
            <div class="MenuboxS">
              <ul>
                <li <%if otype="Main" then%>class="hover"<%end if%>><span><a href="?action=List&otype=Main">1. 选择</a></span></li>
                <li <%if otype="Write" then%>class="hover"<%end if%>><span><a href="?action=Write&otype=Write">2. 编辑</a></span></li>
                <li <%if otype="Send" then%>class="hover"<%end if%>><span><a href="?action=Send&otype=Send">3. 发送</a></span></li>
                <li <%if otype="Report" then%>class="hover"<%end if%>><span><a href="?action=Report&otype=Report">反馈报告</a></span></li>
				<%if Session("CRM_level") = 9 then%>
                <li <%if otype="Manage" then%>class="hover"<%end if%>><span><a href="?action=Manage&otype=Manage">高级管理</a></span></li>
				<%end if%>
				<li class="" id="CheckA"><span><a href="javascript:void(0)" onclick="Showhiden(this,'boxMessages',false,'筛选条件','筛选条件')" style="cursor:pointer;">筛选条件</a></span></li>
              </ul>
            </div>
		</td>
	</tr>
	<tr>
		<td valign="top" class="td_n pdl10 pdr10">
<%
action = Trim(Request("action"))
Select Case action
Case "Write"
    Call infoWrite()
Case "CheckAll"
    Call CheckAll()
Case "Send"
    Call infoSend()
Case "Report"
    Call infoReport()
Case "delReport"
    Call delinfoReport()
Case "Manage"
    Call infoManage()
Case "Managesave"
    Call infoManagesave()
Case Else
    Call infolist()
End Select
%>

<%
Sub infolist()
%>
						<form name="searchForm" action="?subAction=searchItem" method="post">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" id="boxMessages" style="display:none;margin:0 0 10px 0;">
							<col width="100" /><col width="120" /><col width="100" /><col width="120" /><col width="100" /><col width="120" /><col width="100" />
							<tr>
								<td class="td_l_r title"><%=L_Client_cCompany%></td>
								<td class="td_r_l" colspan=3><input name="cCompany" type="text" id="cCompany" class="int" size="40" value="<%=Session("Search_Plugin_Messages_cCompany")%>" ></td>
								<td class="td_l_r title"><%=L_Client_cLastUpdated%></td>
								<td class="td_r_l" colspan=3><input name="ETimeBegin" type="text" maxlength="10" id="ETimeBegin" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Client_cETimeBegin")%>" /> ~ <input name="ETimeEnd" type="text" maxlength="10" id="ETimeEnd" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("Search_Client_cETimeEnd")%>" /> </td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Client_cType%></td>
								<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Type","Type","") %></td>
								<td class="td_l_r title"><%=L_Client_cStart%></td>
								<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Star","Start","") %></td>
								<td class="td_l_r title"><%=L_Client_cSource%></td>
								<td class="td_r_l"><% = EasyCrm.getSelect("SelectData","Select_Source","Source","") %></td>
								<td class="td_l_r title"><%=L_Client_cUser%></td>
								<td class="td_r_l"><% If Session("CRM_level") = 9 Then %><% = EasyCrm.UserList(2,"User","") %><%else%><% = EasyCrm.UserList(1,"User","") %><%end if%></td>
							</tr>
							<tr>
								<td class="td_l_r title"><%=L_Client_cArea%><%=L_Client_cSquare%></td>
								<td class="td_r_l" colspan=3>
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
								<td class="td_l_r title"><%=L_Client_cTrade%></td>
								<td class="td_r_l" colspan=3>
									<select name="Trade" class="int" onchange="getTrade(this.options[this.selectedIndex].id);">
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
										<select name="Strades" class="int">
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
								<td class="td_r_l" colspan="8">
									<input type="submit" name="Submit" class="button42" value=" <%=L_Search%> ">　
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?SubAction=killSession" /></td>
								</tr>
						</table>   
						</form>

<script language="JavaScript">
<!--

for(var i=0;i<document.all.User.options.length;i++){
    if(document.all.User.options[i].value == "<% = Session("Search_Plugin_Messages_cUser") %>"){
    document.all.User.options[i].selected = true;}}

for(var i=0;i<document.all.Type.options.length;i++){
    if(document.all.Type.options[i].value == "<% = Session("Search_Plugin_Messages_cType") %>"){
    document.all.Type.options[i].selected = true;}}

for(var i=0;i<document.all.Start.options.length;i++){
    if(document.all.Start.options[i].value == "<% = Session("Search_Plugin_Messages_cStart") %>"){
    document.all.Start.options[i].selected = true;}}

for(var i=0;i<document.all.Source.options.length;i++){
    if(document.all.Source.options[i].value == "<% = Session("Search_Plugin_Messages_cSource") %>"){
    document.all.Source.options[i].selected = true;}}
	
for(var i=0;i<document.all.Area.options.length;i++){
    if(document.all.Area.options[i].value == "<% = Session("Search_Plugin_Messages_cArea") %>"){
    document.all.Area.options[i].selected = true;}}

for(var i=0;i<document.all.Squares.options.length;i++){
    if(document.all.Squares.options[i].value == "<% = Session("Search_Plugin_Messages_cSquare") %>"){
    document.all.Squares.options[i].selected = true;}}
	
for(var i=0;i<document.all.Trade.options.length;i++){
    if(document.all.Trade.options[i].value == "<% = Session("Search_Plugin_Messages_cTrade") %>"){
    document.all.Trade.options[i].selected = true;}}

for(var i=0;i<document.all.Strades.options.length;i++){
    if(document.all.Strades.options[i].value == "<% = Session("Search_Plugin_Messages_cStrade") %>"){
    document.all.Strades.options[i].selected = true;}}
-->
</script>
		<form id="ListAll" action="?action=CheckAll&PN=<%=PNN%>" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1" style="margin:0px 0;border-bottom:0px;">
				<tr class="tr_t"style="border-bottom:0px;"> 
					<td class="td_l_l">
					<span  style="float:left;padding:0 10px;height:34px;text-align:left;position:fixed;right:10px;top:43px;color:#000;" id="CheckSub">
						<input type="button" class="button41" onclick="javascript:selectall('cId')" value="全选" />
						<input type="submit" name="Submit" class="button45" value="选中" />
						<input type="button" name="button" class="button47" value=" 清空 " onClick=window.location.href="?SubAction=killSession" />
					</span>
					<B>信息列表</B></td>
				</tr>
			</table> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin:0 0 10px 0;">
				<tr class="tr_f">
					<td width="40" class="td_l_c"></td>
					<td width="80" class="td_l_c">编号</td>
					<td class="td_l_l">公司名称</td>
					<td width="100" class="td_l_c">联系人</td>
					<td width="120" class="td_l_c">手机号码</td>
					<td width="100" class="td_l_c">最后发信时间</td>
					<td width="100" class="td_l_c">客户类型</td>
					<td width="100" class="td_l_c">客户等级</td>
				</tr>
				<% = strToPrint %>
			</table>
</table>

<script language=javascript> 
//全选/反选
function selectall(id){ //用id区分  
var tform=document.forms['ListAll'];  
for(var i=0;i<tform.length;i++){  
var e=tform.elements[i];  
if(e.type=="checkbox" && e.id==id) e.checked=!e.checked;  } }
</script> 
			
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd ">
		
			<% = strCounter %>
		</td>
	</tr>
</table>
</div>
			
		</form>

<%
end sub 'onClick=window.location.href="?action=Write&otype=Write" 
%>

<%
Sub CheckAll()

	PN = CLng(ABS(Request("PN")))
	cId=Trim(Request("cId"))
	If cId="" Then
		Response.Write "<script>alert(""没有选择！"");</script>"
		Response.Write "<script>location.href='?action=infolist';</script>"
		Response.End
	else
	cidarr = split(cid,",")
	for i = 0 to Ubound(cidarr)
	if inStr(Session("CRM_Message_Mobile"),EasyCrm.getNewItem("Client","cID",""&cidarr(i)&"","cMobile"))=0 then
	Session("CRM_Message_Mobile") = Session("CRM_Message_Mobile") & EasyCrm.getNewItem("Client","cID",""&cidarr(i)&"","cMobile")&","
	end if
	next
	'Response.Write "<script>alert("""&Session("CRM_Message_Mobile")&""");</script>"
	
	Response.Write "<script>location.href='?action=infolist&PN="&PN+1&"';</script>"
	
	end if
	

end sub
%>

<%
Sub infoReport()
Subaction = Trim(Request("Subaction"))

If Subaction = "Search" Then
    Dim TimeBegin,TimeEnd
	TimeBegin = Trim(Request("TimeBegin"))
	TimeEnd = Trim(Request("TimeEnd"))
	Session("Search_Plugin_Messages_TimeBegin") = Trim(Request("TimeBegin"))
	Session("Search_Plugin_Messages_TimeEnd") = Trim(Request("TimeEnd"))
	Dim Searchsql
    Searchsql = ""	
	if Accsql =1 then
	If TimeBegin <> "" Then
        Searchsql = Searchsql & " And mTime > '" & TimeBegin & "' "
	End If
	If TimeEnd <> "" Then
        Searchsql = Searchsql & " And mTime <= '" & TimeEnd & "' "
	End If
	else
	If TimeBegin <> "" Then
        Searchsql = Searchsql & " And mTime > #" & TimeBegin & "# "
	End If
	If TimeEnd <> "" Then
        Searchsql = Searchsql & " And mTime <= #" & TimeEnd & "# "
	End If	
	End If
End If
If Subaction = "killSession" Then
	Session("Search_Plugin_Messages_TimeBegin") = ""
	Session("Search_Plugin_Messages_TimeEnd") = ""
End If
%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1" style="margin:0px 0;border-bottom:0px;">
				<tr class="tr_t"style="border-bottom:0px;"> 
					<td class="td_l_l"><B>信息列表</B></td>
				</tr>
			</table> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin:0 0 10px 0;">
				<tr class="tr_f">
					<td class="td_l_c" width="50">编号</td>
					<td class="td_l_c">短信内容</td>
					<td class="td_l_c" width="80">手机号码</td>
					<td class="td_l_c" width="80">发信人</td>
					<td class="td_l_c" width="130">发送时间</td>
					<td class="td_l_c" width="60">管理</td>
				</tr>
				<%
				Dim rs
				Dim intTotalRecords,intTotalPages,PN,intPageSize'记录总数，总页数，当前页，分页数量
				PN = CLng(ABS(Request("PN")))

				If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
				intPageSize = DataPageSize
				pagenum = intPageSize*(PN-1)
				
				If Session("CRM_level") < 9 Then
					Searchsql = Searchsql & " And mUser = '"&Session("CRM_name")&"' "
				End If
				
				Set rs = Server.CreateObject("ADODB.Recordset")
					IF PN=1 THEN
					rs.Open "Select top "&intPageSize&" * From [Plugin_Messages] where 1=1 "&Searchsql&" Order By mId desc ",conn,1,1 
					ELSE
					rs.Open "Select top "&intPageSize&" * From [Plugin_Messages] where 1=1 "&Searchsql&" and mId < ( SELECT Min(mId) FROM ( SELECT TOP "&pagenum&" mId FROM [Plugin_Messages] where 1=1 "&Searchsql&" ORDER BY mId desc ) AS T ) Order By mId desc ",conn,1,1
					END IF
					SQLstr="Select count(mId) As RecordSum From [Plugin_Messages] where 1=1 "&Searchsql&" " '统计页码
							
					Dim TotalRecords,TotalPages
					Set Rsstr=conn.Execute(SQLstr,1,1) 
					TotalRecords=Rsstr("RecordSum") 
					if Int(TotalRecords/DataPageSize)=TotalRecords/DataPageSize then
					TotalPages=TotalRecords/DataPageSize
					else
					TotalPages=Int(TotalRecords/DataPageSize)+1
					end if
					Rsstr.Close 
					Set Rsstr=Nothing

					If PN > TotalPages Then PN = TotalPages
											
				Do While Not rs.BOF And Not rs.EOF
				%>
				<Tr>
					<TD class="td_l_c"><%=rs("mId")%></TD>
					<TD class="td_l_l"><%=rs("mContent")%></TD>
					<TD class="td_l_c"><a onclick="Showhiden(this,'box<%=rs("mId")%>',false,'收起','查看')" style="cursor:pointer">查看</a>(<%=ubound(split(""&rs("mPhonenum")&"",","))+1%>)</TD>
					<TD class="td_l_c"><%=rs("mUser")%></TD>
					<TD class="td_l_c"><%=rs("mTime")%></TD>
					<TD class="td_l_c"><input type="button" class="button_info_del" value="" title="删除" onClick=" if(confirm('是否确认删除？'))window.location.href='?action=delReport&mId=<%=rs("mId")%>&PN=<%=PNN%>';else return false;" /></TD>
				</TR>
				<tr class="tr" style="display:none;" id="box<%=rs("mId")%>"><td class="td_l_l" colspan=6 style="padding:10px;background-color:#ffffff;Word-break: break-all; word-wrap:break-word;"><%=rs("mPhonenum")%></td></tr>
						<%
							rs.MoveNext
						Loop
						rs.Close
						Set rs = Nothing
						%>
			</table>
</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd ">
		<span class="r"><form name="searchForm" method="post" action="?action=Report&otype=Report&Subaction=Search">
							<input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" value="<%=Session("Search_Plugin_Messages_TimeBegin")%>" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" value="<%=Session("Search_Plugin_Messages_TimeEnd")%>" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" />&nbsp;<input type="submit" name="Submit" class="button222" value=" <%=L_Search%> "> <input type="button" name="button" class="button223" value=" <%=L_Clear%> " onClick=window.location.href="?action=Report&otype=Report&Subaction=killSession" />
							</form></span>
			<%=EasyCrm.pagelist("?action=Report&otype=Report", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<%
end sub
%>
<%
Sub infoWrite()
%>
<script language="javascript"> 
function countChar(textareaName,spanName){ 
document.getElementById(spanName).innerHTML = document.getElementById(textareaName).value.length;} 
</script> 
	<script language="JavaScript">
	<!--
	function CheckInput()
	{
		if(document.all.mMobile.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '收件人<%=alert04%>'});document.all.mMobile.focus();return false;}
		if(document.all.mInfo.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '短信内容<%=alert04%>'});document.all.mInfo.focus();return false;}
	}
	-->
	</script>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>编辑短信</B></td>
				</tr>
			</table>
			<form name="infoSend" id="infoSend" action="?action=Send&otype=Send" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" />
				<tr>
					<td class="td_l_c title" style="border-top:0;width:100px;"><font color="#color:#CC0000">*</font> 收件人</td>
					<td class="td_r_l" style="border-top:0;padding-right:10px;">
						<textarea name="mMobile" id="mMobile" style="width:100%;height:80px;margin:5px 0;word-break:break-all; table-layout:fixed;"><%=mMobile%></textarea>
					</td>
				</tr>
				<tr>
					<td class="td_l_c title"><font color="#color:#CC0000">*</font> 短信内容</td>
					<td class="td_r_l" style="padding-right:10px;"><textarea name="mInfo" id="mInfo" style="width:100%;height:80px;margin:5px 0;" onkeydown='countChar("mInfo","counter");' onkeyup='countChar("mInfo","counter");'></textarea></td>

				</tr>
					<input name="mUser" type="hidden" id="mUser" value="<%=Session("CRM_name")%>">
			</table>  
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr> 
					<td class="td_l_l" style="border-top:0">
						<input type="submit" name="Submit" class="button45" value=" 发送 ">　
						<input name="Back" type="button" id="Back" class="button43" value=" <%=L_Back%> " onClick="history.back();">　
						已经输入：<span id="counter" style="color:#f00;font-weight:bold;">0</span> 字　(支持60个字，长短信325个字，65个字一条计费)
					</td>
				</tr>
			</table> 
			</form>
				
<%
End Sub

function getHTTPPage(url)
	dim Http
	set Http=server.createobject("MSXML2.XMLHTTP")
	Http.open "GET",url,false
	Http.send()
	if Http.readystate<>4 then 
	exit function
	end if
	getHTTPPage=bytesToBSTR(Http.responseBody,"GB2312")
	set http=nothing
	if err.number<>0 then err.Clear 
end function

Sub infoSend()
	Dim mobiles,message
	mobiles = replacemobile(Trim(Request("mMobile")))
	message = Trim(Request("mInfo"))
	if mobiles="" then 
		Response.Write("<script>alert('手机号码不能为空！');</script>")
		Response.Write("<script>location.href='?action=List&otype=Main' ;</script>")
		Response.end
	end if
	if message="" then 
		Response.Write("<script>alert('短信内容不能为空！');history.back(1);</script>")
		Response.end
	end if
	if message <> "" then
		mRs.Open "select * From [blackdict] ",mconn,1,1
		Do While Not mRs.BOF And Not mRs.EOF
			If InStr(message,""&mRs("content")&"")>0 Then
			Response.Write("<script>alert('有敏感词汇【"&mRs("content")&"】，请重新输入');history.back(1);</script>")
			Response.end
			End if
		mRs.MoveNext
		Loop
		mRs.Close
	End if
	
	sms_url="http://gbk.sms.webchinese.cn/?Uid="&Plugin_Messages_uid&"&Key="&Plugin_Messages_pwd&"&smsMob="&mobiles&"&smsText="&Server.URLEncode(message)&Server.URLEncode(Plugin_Messages_company)&""
	
	status=EasyCrm.getHTTPPage(sms_url)
	
	IF status > "0" THEN
	conn.execute "insert into Plugin_Messages(mState,mContent,mPhonenum,mUser,mTime)values('"&status&"','"&message&"','"&mobiles&"','"&Session("CRM_name")&"','"&now()&"')"
	END IF 
%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>发送报告<font color="#color:#CC0000">(*)</font></B></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" />
				<tr>
					<td class="td_l_c title" style="border-top:0;">反馈信息</td>
					<td class="td_r_l" style="border-top:0;color:#f00;">
						<%if status > "0"  then%>
						发送成功
						<%elseif status="-1" then%>
						没有该账户
						<%elseif status="-2" then%>
						秘钥错误
						<%elseif status="-3" then%>
						短信数量不足
						<%elseif status="-11" then%>
						该用户被禁用
						<%elseif status="-6" then%>
						 IP限制
						<%elseif status="-51" then%>
						签名错误
						<%elseif status="-41" then%>
						手机号为空
						<%elseif status="-42" then%>
						短信内容为空
						<%elseif status="-14" then%>
						内容非法字符
						<%end if%>
					</td>
				</tr>
				<tr>
					<td class="td_r_l" colspan="2">
					<input name="Back" type="button" id="Back" class="button_back" value=" <%=L_Back%> " onClick="history.back();">
					</td>
				</tr>
			</table>   

<%
End Sub

Sub infoManage()
%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>高级配置 <font color="#color:#CC0000">(*)</font></B></td>
				</tr>
			</table>
			<form name="Managesave" action="?action=Managesave" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" />
				<tr>
					<td class="td_l_c title" style="border-top:0;">短信剩余量</td>
					<td class="td_r_l" style="border-top:0;">
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
					<td width="50">
					<%
					Messagesyue = EasyCrm.getHTTPPage("http://sms.webchinese.cn/web_api/SMS/GBK/?Action=SMS_Num&Uid="&Plugin_Messages_uid&"&Key="&Plugin_Messages_pwd&"")
					Response.Write Messagesyue
					%>
					</td>
					<td><span class="info_help help01">查询有延迟，请耐心等待一会！</span></td>
					</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">短信签名</td>
					<td class="td_r_l">
						<input name="Plugin_Messages_company" type="text" class="int" id="Plugin_Messages_company" size="20" value="<%=Plugin_Messages_company%>">
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">账户</td>
					<td class="td_r_l">
						<input name="Plugin_Messages_uid" type="text" class="int" id="Plugin_Messages_uid" size="20" value="<%=Plugin_Messages_uid%>">
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">秘钥</td>
					<td class="td_r_l">
						<input name="Plugin_Messages_pwd" type="text" class="int" id="Plugin_Messages_pwd" size="20" value="<%=Plugin_Messages_pwd%>">
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">权限</td>
					<td class="td_r_l" style="padding:10px;">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="100" />
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
								<input type="checkbox" name="Plugin_Messages_manage" value="<%=rsm("uName")%>" <%if inStr(Plugin_Messages_manage,rsm("uName"))>0 then%>checked<%end if%>><%=rsm("uName")%>　
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
				<tr>
					<td class="td_r_l" colspan="4">
					<input type="submit" name="Submit" class="button45" value=" <%=L_Edit%> ">　
					<input name="Back" type="button" id="Back" class="button43" value=" <%=L_Back%> " onClick="history.back();">
					</td>
				</tr>
			</table>   
			</form>

<%
End Sub

Sub infoManagesave()
	Plugin_Messages_company = replace(Trim(Request.Form("Plugin_Messages_company")),CHR(34),"'")
	Plugin_Messages_uid = replace(Trim(Request.Form("Plugin_Messages_uid")),CHR(34),"'")
	Plugin_Messages_pwd = replace(Trim(Request.Form("Plugin_Messages_pwd")),CHR(34),"'")
	Plugin_Messages_manage = replace(Trim(Request.Form("Plugin_Messages_manage")),CHR(34),"'")
	Dim TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "Dim Plugin_Messages_company,Plugin_Messages_uid,Plugin_Messages_pwd,Plugin_Messages_manage" & VbCrLf
	
	TempStr = TempStr & "'短信账户配置" & VbCrLf
	TempStr = TempStr & "Plugin_Messages_company="& Chr(34) & Plugin_Messages_company & Chr(34) &" '企业简称" & VbCrLf
	TempStr = TempStr & "Plugin_Messages_uid="& Chr(34) & Plugin_Messages_uid & Chr(34) &" '帐号" & VbCrLf
	TempStr = TempStr & "Plugin_Messages_pwd="& Chr(34) & Plugin_Messages_pwd & Chr(34) &" '秘钥" & VbCrLf
	TempStr = TempStr & "Plugin_Messages_manage="& Chr(34) & Plugin_Messages_manage & Chr(34) &" '权限" & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf
	ADODB_SaveToFile TempStr,"Config.asp"
	Response.Write("<script>alert(""修改成功！"");</script>")
	Response.Write "<script>location.href='?action=List&otype=Main';</script>"
End Sub

Sub delinfoReport()
    Dim mId,cId,PNN
	mId = Trim(Request("mId"))
	PNN = Trim(Request("PN"))
	If mId = "" Then
	Exit Sub
	End If
	conn.execute ("delete from Plugin_Messages where mId="&mId&" ")	
	Response.Write("<script>location.href='?action=Report&otype=Report&PN="&PNN&"' ;</script>")
End Sub

Sub clearchoose()
	Response.Cookies("Plugin_msg") = ""
	Response.Redirect("index.asp")
	Response.End
End Sub

Sub ADODB_SaveToFile(ByVal strBody,ByVal File)
	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If
	
	If FSFlag = 1 Then
		Set WriteFile = fs.CreateTextFile(Server.MapPath(File),True)
		WriteFile.Write strBody
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请使用FTP等功能，将<font color=Red >data/config.asp</font>文件内容替换成框中内容"
			Err.Clear
			Set objStream = Noting
			Exit Sub
		End If
		With objStream
			.Type = 2
			.Open
			.Charset = "GB2312"
			.Position = objStream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	End If
End Sub

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>
		</td>
	</tr>
</table>
<script src="../../data/calendar/WdatePicker.js"></script>
<%else%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10">   
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>错误提示</B></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr>
					<td class="td_r_l" style="border-top:0;">
						您无权使用该插件！
					</td>
				</tr>
			</table>   
		</td>
	</tr>
</table>

<%end if%>
</body>
</html>
<% Set EasyCrm = nothing %>