<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
	'��ȡgetֵ
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	uID 	= 	request("uID")
%><%=Header%>
<!-- start header -->
    <div id="header">
         <a href="index.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
         <a onClick=window.location.href="javascript:history.back();" class="button back"><img src="img/back-button.png" width="15" height="16" alt="icon" /></a>
         <a onClick=window.location.href="javascript:window.location.reload();" class="button create"><img src="img/reload-button.png" width="15" height="16" alt="icon" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start page -->
    <div class="page">
<%
Select Case action
Case "InfoShow" '���
    Call InfoShow()
Case else
    Call Main()
End Select

Sub Main()
%>
	<div class="simplebox">

            	<h1 class="titleh">��Ҫͳ��</h1>
					<table class="tabledata"> 
						<tbody> 
						<thead> 
                        <tr> 
							<td >Ա��</td> 
							<td>�ͻ���</td> 
							<td>����</td> 
							<td>����</td> 
							<td>��ͬ</td> 
							<td>�ۺ�</td> 
                        </tr> 
                        </thead> 
						<%
							if Session("CRM_level")=9 then
							sql = sql &""
							else
							sql = sql &"and uName in ("&arrUser&")"
							end if
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.Open "Select * From [user] where 1=1 "&sql&" order by uId asc ",conn,1,1
							Do While Not rs.BOF And Not rs.EOF
						%>
								<tr onclick="location.href='?action=InfoShow&uId=<%=rs("uId")%>'" style="cursor:pointer">
									<td>[<%=rs("uName")%>]</td>
									<td><%=getCountnum("cId","clientnum","Client","cDate",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","cYn",1,"cUser","'"&rs("uName")&"'","cUser")%></td>
									<td><%=getCountnum("cId","Recordsnum","Records","rTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","","","rUser","'"&rs("uName")&"'","rUser")%></td>
									<td><%=getCountnum("cId","Ordernum","Order","oTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","","","oUser","'"&rs("uName")&"'","oUser")%></td>
									<td><%=getCountnum("cId","Hetongnum","Hetong","hTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","","","hUser","'"&rs("uName")&"'","hUser")%></td>
									<td><%=getCountnum("cId","Servicenum","Service","sTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","","","sUser","'"&rs("uName")&"'","sUser")%></td>
								</tr>
						<%
							rs.MoveNext
							Loop
							rs.Close
							Set rs = Nothing
						%>
                    </table> 
					<blockquote>�ݲ��ṩ����ͼЧ��</blockquote>
			</div>
		<%=Footer%>
            
<%
End Sub

Sub InfoShow() '���
	sType = request("sType")
	uName = EasyCrm.getNewItem("User","uID",""&uID&"","uName")
	if Accsql = 1 then 
		Nowdate = "Getdate"
	else
		Nowdate = "Now"
	end if
%>

		<div class="simplebox">
            <h1 class="titleh"><%=uName%>����ϸ����</h1>
            <div class="content">
			<article>�ͻ�����</article>
					<table class="tabledata"> 
						<col width="80">
						<tbody> 
                        <thead>
                        <tr> 
							<td>������״̬</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>�ܼ�¼: <font color=red><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&uName&"' ")%></font> ��</td> 
                        </tr> 
						
                        <tr> 
							<td>7����: <font color=red><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&uName&"' and cDate + 7 >= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>7-30��: <font color=red><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&uName&"' and cDate + 30 >= "&Nowdate&"() and cDate + 7 <= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>30������: <font color=red><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&uName&"' and cDate + 30 <= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <thead>
                        <tr> 
							<td>���ͻ�����</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>
									<% Set rsp = Server.CreateObject("ADODB.Recordset")
									rsp.Open "Select * From [SelectData] where Select_Type<>'' and Select_Type<>'Null' ",conn,1,1
									Do While Not rsp.BOF And Not rsp.EOF
									%><%=rsp("Select_Type")%>��<font color=red><%=getCountnum("cId","clientnum","Client","","","","cType","'"&rsp("Select_Type")&"'","cYn",1,"cUser","'"&uName&"'","cUser")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>
									<% 
									rsp.MoveNext
									Loop
									%>
									δ֪��<font color=red><%=getCountnum("cId","clientnum","Client","","","","cType","''","cYn",1,"cUser","'"&uName&"'","cUser")%></font> ��
									<%
									rsp.Close
									Set rsp = Nothing %>
							</td> 
                        </tr>
                        <thead>
                        <tr> 
							<td>���ͻ�����</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>
									<% Set rss = Server.CreateObject("ADODB.Recordset")
									rss.Open "Select * From [SelectData] where Select_Star<>'' and Select_Star<>'Null' ",conn,1,1
									Do While Not rss.BOF And Not rss.EOF
									%><%=rss("Select_Star")%>��<font color=red><%=getCountnum("cId","clientnum","Client","","","","cStart","'"&rss("Select_Star")&"'","cYn",1,"cUser","'"&uName&"'","cUser")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>
									<% 
									rss.MoveNext
									Loop
									%>
									δ֪��<font color=red><%=getCountnum("cId","clientnum","Client","","","","cStart","''","cYn",1,"cUser","'"&uName&"'","cUser")%></font> ��
									<%
									rss.Close
									Set rss = Nothing %>
							</td> 
                        </tr> 
                        <thead>
                        <tr> 
							<td>���ͻ���Դ</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>
									<% Set rss = Server.CreateObject("ADODB.Recordset")
									rss.Open "Select * From [SelectData] where Select_Source<>'' and Select_Source<>'Null' ",conn,1,1
									Do While Not rss.BOF And Not rss.EOF
									%><%=rss("Select_Source")%>��<font color=red><%=getCountnum("cId","clientnum","Client","","","","cSource","'"&rss("Select_Source")&"'","cYn",1,"cUser","'"&uName&"'","cUser")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>
									<% 
									rss.MoveNext
									Loop
									%>
									δ֪��<font color=red><%=getCountnum("cId","clientnum","Client","","","","cSource","''","cYn",1,"cUser","'"&uName&"'","cUser")%></font> ��
									<%
									rss.Close
									Set rss = Nothing %>
							</td> 
                        </tr> 
                    </table> 
				<article>������¼</article>
					<table class="tabledata"> 
						<col width="80">
						<tbody> 
                        <thead>
                        <tr> 
							<td>������״̬</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>�ܼ�¼: <font color=red><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&uName&"' ")%></font> ��</td> 
                        </tr> 
						
                        <tr> 
							<td>7����: <font color=red><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&uName&"' and rTime + 7 >= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>7-30��: <font color=red><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&uName&"' and rTime + 30 >= "&Nowdate&"() and rTime + 7 <= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>30������: <font color=red><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&uName&"' and rTime + 30 <= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <thead>
                        <tr> 
							<td>����������</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>
									<% Set rsp = Server.CreateObject("ADODB.Recordset")
									rsp.Open "Select * From [SelectData] where Select_Records<>'' and Select_Records<>'Null' ",conn,1,1
									Do While Not rsp.BOF And Not rsp.EOF
									%><%=rsp("Select_Records")%>��<font color=red><%=getCountnum("rId","Recordsnum","Records","","","","rType","'"&rsp("Select_Records")&"'","","","rUser","'"&uName&"'","rUser")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>
									<% 
									rsp.MoveNext
									Loop
									%>
									δ֪��<font color=red><%=getCountnum("rId","Recordsnum","Records","","","","rType","''","","","rUser","'"&uName&"'","rUser")%></font> ��
									<%
									rsp.Close
									Set rsp = Nothing %>
							</td> 
                        </tr>
                        <thead>
                        <tr> 
							<td>����������</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>
									<% Set rsp = Server.CreateObject("ADODB.Recordset")
									rsp.Open "Select * From [SelectData] where Select_Type<>'' and Select_Type<>'Null' ",conn,1,1
									Do While Not rsp.BOF And Not rsp.EOF
									%><%=rsp("Select_Type")%>��<font color=red><%=getCountnum("rId","Recordsnum","Records","","","","rState","'"&rsp("Select_Type")&"'","","","rUser","'"&uName&"'","rUser")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>
									<% 
									rsp.MoveNext
									Loop
									%>
									δ֪��<font color=red><%=getCountnum("rId","Recordsnum","Records","","","","rState","''","","","rUser","'"&uName&"'","rUser")%></font> ��
									<%
									rsp.Close
									Set rsp = Nothing %>
							</td> 
                        </tr> 
                    </table> 
				<article>������¼</article>
					<table class="tabledata"> 
						<col width="80">
						<tbody> 
                        <thead>
                        <tr> 
							<td>������״̬</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>�ܼ�¼: <font color=red><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&uName&"' ")%></font> ��</td> 
                        </tr> 
						
                        <tr> 
							<td>7����: <font color=red><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&uName&"' and oTime + 7 >= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>7-30��: <font color=red><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&uName&"' and oTime + 30 >= "&Nowdate&"() and oTime + 7 <= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>30������: <font color=red><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&uName&"' and oTime + 30 <= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <thead>
                        <tr> 
							<td>������״̬</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>
									δ����<font color=red><%=getCountnum("oId","Ordernum","Order","","","","oState",0,"","","oUser","'"&uName&"'","oUser")%></font> ��
                        </tr> 
                        <tr> 
							<td>
									�����У�<font color=red><%=getCountnum("oId","Ordernum","Order","","","","oState",1,"","","oUser","'"&uName&"'","oUser")%></font> ��
							</td> 
                        </tr>
                        <tr> 
							<td>
									����ɣ�<font color=red><%=getCountnum("oId","Ordernum","Order","","","","oState",2,"","","oUser","'"&uName&"'","oUser")%></font> ��
							</td> 
                        </tr>
                    </table> 
				<article>��ͬ��¼</article>
					<table class="tabledata"> 
						<col width="80">
						<tbody> 
                        <thead>
                        <tr> 
							<td>������״̬</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>�ܼ�¼: <font color=red><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&uName&"' ")%></font> ��</td> 
                        </tr> 
						
                        <tr> 
							<td>7����: <font color=red><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&uName&"' and hTime + 7 >= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>7-30��: <font color=red><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&uName&"' and hTime + 30 >= "&Nowdate&"() and hTime + 7 <= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>30������: <font color=red><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&uName&"' and hTime + 30 <= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <thead>
                        <tr> 
							<td>����ͬ����</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>
									<% Set rsp = Server.CreateObject("ADODB.Recordset")
									rsp.Open "Select * From [SelectData] where Select_Hetong<>'' and Select_Hetong<>'Null' ",conn,1,1
									Do While Not rsp.BOF And Not rsp.EOF
									%><%=rsp("Select_Hetong")%>��<font color=red><%=getCountnum("hId","Hetongnum","Hetong","","","","hType","'"&rsp("Select_Hetong")&"'","","","hUser","'"&uName&"'","hUser")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>
									<% 
									rsp.MoveNext
									Loop
									%>
									δ֪��<font color=red><%=getCountnum("hId","Hetongnum","Hetong","","","","hType","''","","","hUser","'"&uName&"'","hUser")%></font> ��
									<%
									rsp.Close
									Set rsp = Nothing %>
							</td> 
                        </tr>
						<% Contrs = conn.execute ("select sum(hMoney) as AllMoney,sum(hRevenue) as AllRevenue,sum(hOwed) as AllOwed from [Hetong] where hUser='"&uName&"' ") %><%if Contrs("AllMoney")<>"" then%>
                        <thead>
                        <tr> 
							<td>��ͬ���</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>�ܽ�<font color=red><%=Contrs("AllMoney")%></font> Ԫ</td> 
                        </tr> 
                        <tr> 
							<td>���գ�<font color=red><%=Contrs("AllRevenue")%></font> Ԫ</td> 
                        </tr> 
                        <tr> 
							<td>Ƿ�<font color=red><%=Contrs("AllOwed")%></font> Ԫ</td> 
                        </tr> 
						<%end if%>
                    </table> 
				<article>�ۺ��¼</article>
					<table class="tabledata"> 
						<col width="80">
						<tbody> 
                        <thead>
                        <tr> 
							<td>������״̬</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>�ܼ�¼: <font color=red><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&uName&"' ")%></font> ��</td> 
                        </tr> 
						
                        <tr> 
							<td>7����: <font color=red><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&uName&"' and sTime + 7 >= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>7-30��: <font color=red><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&uName&"' and sTime + 30 >= "&Nowdate&"() and sTime + 7 <= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>30������: <font color=red><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&uName&"' and sTime + 30 <= "&Nowdate&"() ")%></font> ��</td> 
                        </tr> 
                        <thead>
                        <tr> 
							<td>����ͬ����</td> 
                        </tr> 
                        </thead> 
                        <tr> 
							<td>
									<% Set rsp = Server.CreateObject("ADODB.Recordset")
									rsp.Open "Select * From [SelectData] where Select_Service<>'' and Select_Service<>'Null' ",conn,1,1
									Do While Not rsp.BOF And Not rsp.EOF
									%><%=rsp("Select_Service")%>��<font color=red><%=getCountnum("sId","Servicenum","Service","","","","sType","'"&rsp("Select_Service")&"'","","","sUser","'"&uName&"'","sUser")%></font> ��</td> 
                        </tr> 
                        <tr> 
							<td>
									<% 
									rsp.MoveNext
									Loop
									%>
									δ֪��<font color=red><%=getCountnum("sId","Servicenum","Service","","","","sType","''","","","sUser","'"&uName&"'","hUser")%></font> ��
									<%
									rsp.Close
									Set rsp = Nothing %>
							</td> 
                        </tr>
                    </table> 
			</div>
		</div>
<%
End Sub

Function getCountnum(Item0,Item1,Item2,Item3,Item4,Item5,Item6,Item7,Item8,Item9,Item10,Item11,Item12)
    Dim rs,itemValue,sql
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = ""
	if Accsql =1 then
	if Item4<>"" then
	sql = sql & " and "&Item3&" >= '"&Item4&"'"
	end if
	if Item5<>"" then
	sql = sql & " and "&Item3&" <= '"&Item5&"'"
	end if
	else
	if Item4<>"" then
	sql = sql & " and "&Item3&" >= #"&Item4&"#"
	end if
	if Item5<>"" then
	sql = sql & " and "&Item3&" <= #"&Item5&"#"
	end if
	end if
	if Item6<>"" and Item7<>"" then
	sql = sql & " and "&Item6&" = " & Item7&""
	end if
	if Item8<>"" and Item9<>"" then
	sql = sql & " and "&Item8&" = " & Item9&""
	end if
	if Item10<>"" and Item11<>"" then
	sql = sql & " and "&Item10&" = " & Item11&""
	end if
	
	If Session("CRM_level") < 9 Then
		sql = sql & " And "&Item12&" In (" & arrUser & ")"
	else
		sql = sql & ""
	end if
	
	rs.Open "Select count("&Item0&") As "&Item1&" From ["&Item2&"] Where 1=1 "&sql&" " ,conn,1,1
	    itemValue = rs(Item1)
	rs.Close
	Set rs = Nothing
	getCountnum = itemValue
End Function
%>
    <div class="clear"></div>
    </div>
    <!-- end page -->
<script type="text/javascript" src="js/frame.js"></script>
</body>
</html>
