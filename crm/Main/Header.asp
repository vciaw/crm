<!--#include file="../data/conn.asp" -->
<!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
if action="getmessage" then 
	
	'调用站内信的未读信息数量
	set rs=conn.execute("select count(id) As newmsg from OA_mms_Receive where ( oIsread = 0 and oReceiver = '"&Session("CRM_name")&"' and oAttime is null ) or ( oAttime is not null and oAttime < getdate()+ 0.007 and oIsread = 0 and oReceiver = '"&Session("CRM_name")&"' ) ")
	if rs("newmsg")<>0 then 
		Response.Write "<a href='../OA/Receive.asp' target='main' title='站内信'><input class='button_top_mail' type='button' value=' '  onClick=window.parent.main.location.href='../OA/Receive.asp' /><span class='notification'>"&rs("newmsg")&"</span></a>"
	end if
	
	'调用工作报告的未读信息数量
	'If mid(Session("CRM_qx"), 46, 1) = "1" then
	If Session("CRM_level") = 9 Then									'管理员管理所有的
		sql = sql & " "
	elseIf Session("CRM_level") < 9 and Session("CRM_level") > 1 Then 	'部门经理可以看到别人提交给自己的，和自己提交的
		sql = sql & " and ( oSendto like '%"&Session("CRM_name")&"%' or oUser = '"&Session("CRM_name")&"' ) "
	else								
		sql = sql & " and oUser = '"&Session("CRM_name")&"' " 			'普通员工只能看到自己提交的报告
	end if
	set rs=conn.execute("select count(id) As newreport from OA_Report where oIsread = 0 "&sql&"  ")
	if rs("newreport")<>0 then 
		Response.Write "<a href='../OA/Report.asp' target='main' title='工作报告'><input class='button_top_reprot' type='button' value=' ' onClick=window.parent.main.location.href='../OA/Report.asp' /><span class='notification'>"&rs("newreport")&"</span></a>"
	end if
	'end if
	
	'调用内部聊天的未读信息数量
	Dim wconn,wMDBPath
	set wrs=server.CreateObject("adodb.recordset")
	Set wconn = Server.CreateObject("ADODB.Connection")
	wMDBPath = Server.MapPath("../Plugin/WEBIM/DataBase/#WebIMdata.mdb")
	wconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wMDBPath
	Set mRs=wconn.Execute("select UserID From [User] Where UserName = '"&Session("CRM_name")&"' ",1,1)
	
	set wrs=wconn.Execute("select count(id) As newmsg from [UserMsg] where IsRead = 2 and ToID = "&mRs("UserID")&" ",1,1)
	if wrs("newmsg")<>0 then 
		Response.Write "<a href='../Plugin/WEBIM/' target='main' title='内部聊天'><input class='button_top_im' type='button' value=' ' onClick=window.parent.main.location.href='../Plugin/WEBIM/' /><span class='notification'>"&wrs("newmsg")&"</span></a>"
	end if
	
else
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><%=title%></title>
<link href="<%=SiteUrl&skinurl%>Style/Common.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script type="text/javascript" language="javascript">
var timeout = window.setInterval("sendRequest()", 5000);
var xmlHttp = false;
</script>
</head>
<body onLoad="sendRequest();">
    <div id="header">
      <div class="con">
        <div id="logo"></div>
        <div id="menu">
            <ul>
              <li class="index"><a href="menu.asp" id="item0" target="menu" class="active" onclick="Tabmenu(this,0);"><%=L_Header_title%></a></li>
              <% If mid(Session("CRM_qx"), 2, 1) = 1 Then %>
              <li><a href="menu.asp?action=Listall" id="item1" target="menu" onclick="Tabmenu(this,1);"><%=L_Header_company%></a></li>
			  <%end if%>
              <% If mid(Session("CRM_qx"), 3, 1) = "1" Then %>
              <li><a href="menu.asp?action=Company" target="menu" id="item2" onclick="Tabmenu(this,2);"><%=L_Header_oa%></a></li>
			  <%end if%>
              <% If mid(Session("CRM_qx"), 4, 1) = "1" Then %>
              <li><a href="menu.asp?action=Plugin" target="menu" id="item4" onclick="Tabmenu(this,4);"><%=L_Header_plugin%></a></li>
			  <%end if%>
              <% If mid(Session("CRM_qx"), 5, 1) = "1" Then %>
              <li><a href="menu.asp?action=System" target="menu" id="item5" onclick="Tabmenu(this,5);"><%=L_Header_manage%></a></li>
              <%end if%>
            </ul>
        </div><!--/ menu-->
		<div id="sendRequestContent"></div>
        <div id="info">
		  <%if Session("CRM_account")<>"" then%>帐号：<a href="../system/User_info.asp?uid=<% = Session("CRM_uid") %>" target='main' title="<%if Session("CRM_level") <> "" then%><%=EasyCrm.getNewItem("system_group","gId",Session("CRM_group"),"gName")%>&nbsp;&nbsp;<%=EasyCrm.getNewItem("system_level","lId",Session("CRM_level"),"lName")%><%else%><%=L_Header_no_login%><%end if%>"><%=Session("CRM_name")%></a><%else%><%=L_Header_no_login%><%end if%>　<a href="logout.asp" target="main"><%=L_Header_logout%></a>
        </div>
      </div><!--/ con-->
    </div><!--/ header-->
<script type="text/javascript">
eval(function(p,a,c,k,e,r){e=function(c){return c.toString(36)};if('0'.replace(0,e)==0){while(c--)r[e(c)]=k[c];k=[function(e){return r[e]||e}];e=function(){return'[1-9b-e]'};c=1};while(c--)if(k[c])p=p.replace(new RegExp('\\b'+e(c)+'\\b','g'),k[c]);return p}('4 5(3,n){1 2=6.7("menu").getElementsByTagName("a");for(1 i=0,8=2.length;i<8;++i){9(2[i].clssName!==""){2[i].b=""}3.b="active";3.blur();c.d=n}};(4(){1 n=c.d.replace("#","");9(!n){n=0}1 e=6.7("item"+n);5(e,n)})();',[],15,'|var|Items|obj|function|Tabmenu|document|getElementById|len|if||className|location|hash|curitem'.split('|'),0,{}))
</script>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html>
<%end if%><% Set EasyCrm = nothing %>