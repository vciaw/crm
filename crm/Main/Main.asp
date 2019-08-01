<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
if Accsql = 1 then 
	Nowdate = "Getdate"
else
	Nowdate = "Now"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<title><%=title%></title>
<link href="<%=SiteUrl&skinurl%>Style/Common.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<link href="<%=SiteUrl&skinurl%>Style/inettuts.css" rel="stylesheet" type="text/css" />
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
<style> li {behavior: url("../skin/default/style/ie-css3.htc");}</style>
<script language="JavaScript">
<!--
function killerrors() { return true; } 
window.onerror = killerrors;
-->
</script>
</head>
<body>
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><span class=info_main>欢迎使用 <%=title%>，祝您工作顺利，有个好心情！</span></td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
        </td>
	</tr>
</table>

    <div id="columns">
        
        <ul id="column1" class="column">
            <li class="widget color-1">
                <div class="widget-head">
                    <h3><a href="../OA/Notice.asp">内部公文</a></h3>
                </div>
                <div class="widget-content">
					<ul>
						<%
						Set rs1 = Server.CreateObject("ADODB.Recordset")
						rs1.Open "Select Top 5 * From OA_Notice Order By ONId Desc",conn,3,1
						i=0
						If rs1.RecordCount > 0 Then
						If Not rs1.Eof Then 
						do while not (rs1.eof or err)
						i=i+1
						%>
						<li <%if i=rs1.recordcount then%>class='none'<%end if%>><span class="r"><%=EasyCrm.FormatDate(rs1("ONaddtime"),10)%></span>[<%=rs1("ONclass")%>] <a onclick='Notice_InfoView<%=rs1("ONId")%>()' style="cursor:pointer" ><%=rs1("ONtitle")%></a></li>
						<script>function Notice_InfoView<%=rs1("ONId")%>() {$.dialog.open('../OA/GetUpdate.asp?action=Notice&sType=View&Id=<%=rs1("ONId")%>', {title: '查看', width: 800,height: 480, fixed: true}); };</script>
						<%
						  rs1.movenext
						  loop
						  end if
						else
						%>
						<li class='none'><%=L_Notfound%></li>
						<%
						end if  
						  rs1.close
						  set rs1=nothing
						%>
					</ul> 
                </div>
            </li>
            <li class="widget color-2">  
                <div class="widget-head">
                    <h3>站内信通知</h3>
                </div>
                <div class="widget-content">
					<ul>
						<%
						Set rs1 = Server.CreateObject("ADODB.Recordset")
						rs1.Open "Select Top 5 * From [OA_mms_Receive] WHERE ( oIsread = 0 and oReceiver = '"&Session("CRM_name")&"' and oAttime is null ) or ( oAttime is not null and oAttime < "&Nowdate&"()+ 0.007 and oIsread = 0 and oReceiver = '"&Session("CRM_name")&"' ) Order By Id Desc",conn,3,1
						i=0
						If rs1.RecordCount > 0 Then
						If Not rs1.Eof Then 
						do while not (rs1.eof or err)
						i=i+1
						%>
						<li <%if i=rs1.recordcount then%>class='none'<%end if%>><span class="r" style="padding-left:10px;"><%=EasyCrm.FormatDate(rs1("oTime"),10)%></span><a onclick='Receive_InfoEdit<%=rs1("Id")%>()' style="cursor:pointer" ><%=rs1("oTitle")%></a></li>
						<script>function Receive_InfoEdit<%=rs1("Id")%>() {$.dialog.open('../OA/Receive.asp?action=Reply&id=<%=rs1("Id")%>', {title: '回复', width: 900,height: 470, fixed: true}); };</script>
						<%
						  rs1.movenext
						  loop
						  end if
						else
						%>
						<li class='none'><%=L_Notfound%></li>
						<%
						end if  
						  rs1.close
						  set rs1=nothing
						%>
					</ul>
                </div>
            </li>
            <li class="widget color-3">  
                <div class="widget-head">
                    <h3><a href="../OA/Calendar.asp">个人日历</a></h3>
                </div>
                <div class="widget-content">
					<ul>
					<%
					Set rs1 = Server.CreateObject("ADODB.Recordset")
					rs1.Open "Select Top 5 * From [Calendar] Where calendaruser = '"&Session("CRM_name")&"' Order By calendarDate Desc",conn,3,1
					i=0
					If rs1.RecordCount > 0 Then
					If Not rs1.Eof Then
					do while not (rs1.eof or err)
					i=i+1
					%>
						<li <%if i=rs1.recordcount then%>class='none'<%end if%>><span class="r"><%=EasyCrm.FormatDate(rs1("calendarDate"),10)%></span><a onclick='Calendar_InfoEdit<%=rs1("Id")%>()' style="cursor:pointer" ><%=left(rs1("calendarText"),20)%></a></li>
						<script>function Calendar_InfoEdit<%=rs1("id")%>() {$.dialog.open('../OA/GetUpdate.asp?action=Calendar&sType=Edit&id=<%=rs1("id")%>', {title: '编辑', width: 400, height: 280,fixed: true}); };</script>
					<%
					  rs1.movenext
					  loop
					end if
					else
					%>
						<li class='none'><%=L_Notfound%></li>
					<%
					end if
					  rs1.close
					  set rs1=nothing
					%>
					</ul>
                </div>
            </li>
        </ul>

        <ul id="column2" class="column">
            <li class="widget color-4">  
                <div class="widget-head">
                    <span style="float:right;padding-right:10px;padding-top:5px;"><span class="info_help help01" onmouseover="tip.start(this)" tips="三天内需跟单的记录">&nbsp;</span></span><h3><a href="Records.asp">跟单提醒</a></h3>
                </div>
                <div class="widget-content">
					<ul>
						<%
						Set rs1 = Server.CreateObject("ADODB.Recordset")
						
							if Session("CRM_level") = 9 Then
							rs1.Open "Select Top 5 * From Records where 1=1 and rNextTime + 3 >= "&Nowdate&"() and rNextTime - 3 <= "&Nowdate&"() Order By rNextTime Desc",conn,3,1
							else
							rs1.Open "Select Top 5 * From Records where rUser='" & Session("CRM_name") & "' and rNextTime + 3 >= "&Nowdate&"() and rNextTime - 3 <= "&Nowdate&"() Order By rNextTime Desc",conn,3,1
							end if
						i=0
						If rs1.RecordCount > 0 Then
						If Not rs1.Eof Then 
						do while not (rs1.eof or err)
						i=i+1
						%>
						<li <%if i=rs1.recordcount then%>class='none'<%end if%>><span class="r" style="<%if EasyCrm.FormatDate(rs1("rNextTime"),2) = EasyCrm.FormatDate(date(),2) then %>color:#ff0000;<% end if %>"><%=EasyCrm.FormatDate(rs1("rNextTime"),10)%></span>[<%=rs1("rState")%>] <a onclick='Records_InfoView<%=rs1("rId")%>()' style="cursor:pointer" > <%=EasyCrm.getNewItem("Client","cid",rs1("cid"),"cCompany")%></a></li>
						<script>function Records_InfoView<%=rs1("rId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&oType=Records&cId=<%=rs1("cId")%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
						<%
						  rs1.movenext
						  loop
						  end if
						else
						%>
						<li class='none'><%=L_Notfound%></li>
						<%
						end if
						  rs1.close
						  set rs1=nothing
						%>
					</ul>
                </div>
            </li>
            <li class="widget color-5">  
                <div class="widget-head">
                    <span style="float:right;padding-right:10px;padding-top:5px;"><span class="info_help help01" onmouseover="tip.start(this)" tips="未处理和处理中的订单">&nbsp;</span></span><h3><a href="Order.asp">未完成订单</a></h3>
                </div>
                <div class="widget-content">
					<ul>
						<%
						Set rs1 = Server.CreateObject("ADODB.Recordset")
						
							if Session("CRM_level") = 9 Then
							rs1.Open "Select Top 5 * From [Order] where 1=1 and oState in (0,1) Order By oid Desc",conn,3,1
							else
							rs1.Open "Select Top 5 * From [Order] where oUser='" & Session("CRM_name") & "' and oState in (0,1) Order By oid Desc",conn,3,1
							end if
						i=0
						If rs1.RecordCount > 0 Then
						If Not rs1.Eof Then 
						do while not (rs1.eof or err)
						i=i+1
						%>
						<li <%if i=rs1.recordcount then%>class='none'<%end if%>><span class="r" style="<%if EasyCrm.FormatDate(rs1("oSDate"),2) = EasyCrm.FormatDate(date(),2) then %>color:#ff0000;<% end if %>"><%=EasyCrm.FormatDate(rs1("oSDate"),10)%></span>[<%if rs1("oState") = 0 then%>未处理<%elseif rs1("oState") = 1 then%>处理中<%elseif rs1("oState") = 2 then%>已完成<%elseif rs1("oState") = 3 then%>已取消<%end if%>] <a onclick='Order_InfoView<%=rs1("oId")%>()' style="cursor:pointer" > <%=EasyCrm.getNewItem("Client","cid",rs1("cid"),"cCompany")%></a></li>
						<script>function Order_InfoView<%=rs1("oId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&oType=Order&cId=<%=rs1("cId")%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
						<%
						  rs1.movenext
						  loop
						  end if
						else
						%>
						<li class='none'><%=L_Notfound%></li>
						<%
						end if
						  rs1.close
						  set rs1=nothing
						%>
					</ul>
                </div>
            </li>
            <li class="widget color-6">  
                <div class="widget-head">
                    <span style="float:right;padding-right:10px;padding-top:5px;"><span class="info_help help01" onmouseover="tip.start(this)" tips="已过期不显示">&nbsp;</span></span><h3><a href="Hetong.asp">合同提醒</a></h3>
                </div>
                <div class="widget-content">
					<ul>
						<%
						Set rs1 = Server.CreateObject("ADODB.Recordset")
							if Session("CRM_level") = 9 Then
							rs1.Open "Select Top 5 * From Hetong Where hedate + 1 >= "&Nowdate&"() Order By hedate asc",conn,3,1
							else
							rs1.Open "Select Top 5 * From Hetong Where hedate + 1 >= "&Nowdate&"() and hUser In (" & arrUser & ") Order By hedate asc",conn,3,1
							end if
						i=0
						If rs1.RecordCount > 0 Then
						If Not rs1.Eof Then
						do while not (rs1.eof or err)
						i=i+1
						%>
						<li <%if i=rs1.recordcount then%>class='none'<%end if%>><span class="r" style="<%if rs1("hEdate") <= now() then %>color:#ff0000;<% end if %>"><%=EasyCrm.FormatDate(rs1("hEdate"),10)%></span><a onclick='Hetong_InfoView<%=rs1("hId")%>()' style="cursor:pointer" ><%=EasyCrm.getNewItem("Client","cid",rs1("cid"),"cCompany")%></a></li>
						<script>function Hetong_InfoView<%=rs1("hId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&oType=Hetong&cId=<%=rs1("cId")%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
						<%
						  rs1.movenext
						  loop
						end if
						else
						%>
						<li class='none'><%=L_Notfound%></li>
						<%
						end if
						  rs1.close
						  set rs1=nothing
						%>
					</ul>
                </div>
            </li>
        </ul>
        <ul id="column3" class="column">
            <li class="widget color-7" id="intro">  
                <div class="widget-head">
                    <h3>概要统计</h3>
                </div>
                <div class="widget-content">
					<ul style="padding-top:10px;">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1" style="border:0;margin:5px 0;">
						<col width="20%"><col width="20%"><col width="20%"><col width="20%">
						<style>.td_l_c {border:0;border-bottom:1px solid #ddd;}
						</style>
							<tr>
								<td class="td_l_c"></td>
								<td class="td_l_c">7天内</td>
								<td class="td_l_c">7-30</td>
								<td class="td_l_c">>30</td>
								<td class="td_l_c">总量</td>
							</tr>
							<%if Session("CRM_level") < 9 then %>
							<tr>
								<td class="td_l_c">客户</td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&Session("CRM_name")&"' and cDate + 7 >= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&Session("CRM_name")&"' and cDate + 30 >= "&Nowdate&"() and cDate + 7 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&Session("CRM_name")&"' and cDate + 30 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&Session("CRM_name")&"' ")%></td>
							</tr>
							<tr>
								<td class="td_l_c">跟单</td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&Session("CRM_name")&"' and rTime + 7 >= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&Session("CRM_name")&"' and rTime + 30 >= "&Nowdate&"() and rTime + 7 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&Session("CRM_name")&"' and rTime + 30 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&Session("CRM_name")&"' ")%></td>
							</tr>
							<tr>
								<td class="td_l_c">订单</td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&Session("CRM_name")&"' and oTime + 7 >= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&Session("CRM_name")&"' and oTime + 30 >= "&Nowdate&"() and oTime + 7 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&Session("CRM_name")&"' and oTime + 30 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&Session("CRM_name")&"' ")%></td>
							</tr>
							<tr>
								<td class="td_l_c">合同</td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&Session("CRM_name")&"' and hTime + 7 >= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&Session("CRM_name")&"' and hTime + 30 >= "&Nowdate&"() and hTime + 7 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&Session("CRM_name")&"' and hTime + 30 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&Session("CRM_name")&"' ")%></td>
							</tr>
							<tr>
								<td class="td_l_c" style="border:0;">售后</td>
								<td class="td_l_c" style="border:0;"><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&Session("CRM_name")&"' and sTime + 7 >= "&Nowdate&"() ")%></td>
								<td class="td_l_c" style="border:0;"><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&Session("CRM_name")&"' and sTime + 30 >= "&Nowdate&"() and sTime + 7 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c" style="border:0;"><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&Session("CRM_name")&"' and sTime + 30 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c" style="border:0;"><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&Session("CRM_name")&"' ")%></td>
							</tr>
							<%else%>
							<tr>
								<td class="td_l_c">客户</td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cDate + 7 >= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cDate + 30 >= "&Nowdate&"() and cDate + 7 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cDate + 30 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 ")%></td>
							</tr>
							<tr>
								<td class="td_l_c">跟单</td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Records","rid","rid"," and rTime + 7 >= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Records","rid","rid"," and rTime + 30 >= "&Nowdate&"() and rTime + 7 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Records","rid","rid"," and rTime + 30 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Records","rid","rid"," ")%></td>
							</tr>
							<tr>
								<td class="td_l_c">订单</td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Order","oid","oid"," and oTime + 7 >= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Order","oid","oid"," and oTime + 30 >= "&Nowdate&"() and oTime + 7 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Order","oid","oid"," and oTime + 30 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Order","oid","oid"," ")%></td>
							</tr>
							<tr>
								<td class="td_l_c">合同</td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hTime + 7 >= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hTime + 30 >= "&Nowdate&"() and hTime + 7 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hTime + 30 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c"><%=EasyCrm.getCountItem("Hetong","hid","hid"," ")%></td>
							</tr>
							<tr>
								<td class="td_l_c" style="border:0;" >售后</td>
								<td class="td_l_c" style="border:0;"><%=EasyCrm.getCountItem("Service","sid","sid"," and sTime + 7 >= "&Nowdate&"() ")%></td>
								<td class="td_l_c" style="border:0;"><%=EasyCrm.getCountItem("Service","sid","sid"," and sTime + 30 >= "&Nowdate&"() and sTime + 7 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c" style="border:0;"><%=EasyCrm.getCountItem("Service","sid","sid"," and sTime + 30 <= "&Nowdate&"() ")%></td>
								<td class="td_l_c" style="border:0;"><%=EasyCrm.getCountItem("Service","sid","sid"," ")%></td>
							</tr>
							<%end if%>
						</table> 
					</ul>
                </div>
            </li>
            <li class="widget color-8">  
                <div class="widget-head">
                    <h3>生日提醒 <span class='remind'><%if Accsql = 1 then%><%=EasyCrm.getCountItem("Linkmans","lid","lid"," and lBirthday = "&date()&" ")%><%else%><%=EasyCrm.getCountItem("Linkmans","lid","lid"," and lBirthday = #"&date()&"# ")%><%end if%></span>　</h3>
                </div>
                <div class="widget-content">
					<ul>
					<%
					Set rs = Server.CreateObject("ADODB.Recordset")
					if Accsql = 1 then
					rs.Open "Select * From Linkmans Where lUser In (" & arrUser & ") and dateadd(year,year(getdate())-year(lBirthday),lBirthday) between getdate()-1 and dateadd(day,10,getdate())  Order By lid Desc",conn,3,1
					else
					rs.Open "Select * From Linkmans Where lBirthday = #"&date()&"# Order By lid Desc",conn,3,1
					end if
					i=0
					If rs.RecordCount > 0 Then
					Do While Not rs.BOF And Not rs.EOF
					i=i+1
					%>
						<li <%if i=rs.recordcount then%>class='none'<%end if%> ><span class="r"><%=rs("lName")%>&nbsp; <%=EasyCrm.FormatDate(rs("lBirthday"),2)%></span><img src="<%=SiteUrl&skinurl%>images/ico/sr.gif">&nbsp; <a onclick='Client_InfoView<%=rs("cId")%>()' style="cursor:pointer" ><%=EasyCrm.getNewItem("Client","cId",rs("cId"),"cCompany")%></a></li>
						<script>function Client_InfoView<%=rs("cId")%>() {$.dialog.open('GetUpdate.asp?action=Client&sType=InfoView&oType=Linkmans&cId=<%=rs("cId")%>', {title: '查看', width: 900,height: 480, fixed: true}); };</script>
					<%
					rs.MoveNext
					Loop
					else
					%>
						<li class='none'><%=L_Notfound%></li>
					<%
					end if
					  rs.close
					  set rs=nothing
					%>
					</ul>
                </div>
            </li>
        </ul>
        
    </div>
	
<div class="PX10"></div>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
</body>
</html><% Set EasyCrm = nothing %>