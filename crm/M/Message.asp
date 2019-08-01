<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
	'获取get值
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	Sendname=Request.QueryString("username")
	oIsread	=Request.QueryString("Isread")
	id		=Request.QueryString("id")
	if oIsread = "" then oIsread = 0
	if Accsql = 1 then 
		Nowdate = "Getdate"
	else
		Nowdate = "Now"
	end if
%><%=Header%>
<!-- start header -->
    <div id="header">
         <a href="index.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
		 <%if id<>"" then%>
         <a onClick=window.location.href="javascript:history.back();" class="button back"><img src="img/back-button.png" width="15" height="16" alt="icon" /></a>
		 <%else%>
         <a href="#" class="button list"><img src="img/create.png" width="16" height="16" alt="icon"/></a>
		 <%end if%>
         <a onClick=window.location.href="javascript:window.location.reload();" class="button create"><img src="img/reload-button.png" width="15" height="16" alt="icon" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start page -->

<%
Select Case action
Case "InfoShow" '查看
    Call InfoShow()
Case "SaveAdd" '添加
    Call SaveAdd()
Case "SaveReply" '回复
    Call SaveReply()
Case Else
    Call Main()
End Select

Sub Main()
%>
    <div class="page">
	<div class="simplebox">
	
	<script language="JavaScript">
	<!-- 必填项提示
	function CheckInput()
	{
		if(document.getElementById('oReceiver').value == ""){alert("<%=L_Mms_oReceiver%>不能为空！"); document.all.oReceiver.focus();return false;}
		if(document.getElementById('oTitle').value == ""){alert("<%=L_Mms_oTitle%>不能为空！"); document.all.oTitle.focus();return false;}
	}
	-->
	</script>
                <div class="content listbox" style="margin-bottom:10px;">
				<form name="Save" action="?action=SaveAdd" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> <%=L_Mms_oReceiver%></label>
						<% = EasyCrm.UserList(2,"oReceiver","") %>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> <%=L_Mms_oTitle%></label>
						<input name="oTitle" type="text" class="int" id="oTitle" style="width:80%;">
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><%=L_Mms_oContent%></label>
						<textarea name="oContent" id="oContent" class="int" style="width:80%;"></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="oSender" type="hidden" value="<%=Session("CRM_name")%>">
                    <input type="submit" name="button" id="button" value="&nbsp;保 存&nbsp;" class="submit-button" />
                    </div>
                </form>
                </div>
				
            	<h1 class="titleh">站内信</h1>
					<table class="tabledata"> 
						<col width="70" >
						<tbody> 
                        <tr> 
							<td>发件人</td> 
							<td>标题</td> 
                        </tr> 
						<%
					Dim sql
					sql = ""
					if Sendname <>"" then
					sql = sql & " and oSender ='"&Sendname&"'"
					end if
					if oIsread = 1 then
					sql = sql & " and oIsread ="&oIsread&" and oReceiver = '"&Session("CRM_name")&"'"
					else
					sql = sql & " and ( oIsread = 0 and oReceiver = '"&Session("CRM_name")&"' and oAttime is null ) or ( oAttime is not null and oAttime < "&Nowdate&"()+ 0.007 and oIsread = 0 and oReceiver = '"&Session("CRM_name")&"' ) "
					end if
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.Open "Select * From [OA_mms_Receive] where 1=1 "&sql&" order by Id desc ",conn,1,1
							Do While Not rs.BOF And Not rs.EOF
						%>
								<tr>
									<td>[<%=rs("oSender")%>]</td>
									<td><a href="?action=InfoShow&Id=<%=rs("Id")%>"><%=rs("oTitle")%></a></td>
								</tr>
						<%
							rs.MoveNext
							Loop
							rs.Close
							Set rs = Nothing
						%>
                    </table> 
					<blockquote>可查看、回复，暂不提供删除权限</blockquote>
			</div>
		<%=Footer%>
            
<%
end Sub

Sub SaveAdd() '添加
    Dim oReceiver,oSender,oTitle,oContent,oIsread
	oReceiver = Request.Form("oReceiver")
	oSender = Request.Form("oSender")
	oTitle = Request.Form("oTitle")
	oContent = Request.Form("oContent")
	conn.execute("insert into [OA_mms_send] (oReceiver,oSender,oTitle,oContent,oIsread,oTime) values('"&oReceiver&"','"&oSender&"','"&oTitle&"','"&oContent&"',0,'"&Now()&"')")
	conn.execute("insert into [OA_mms_Receive] (oReceiver,oSender,oTitle,oContent,oIsread,oTime) values('"&oReceiver&"','"&oSender&"','"&oTitle&"','"&oContent&"',0,'"&Now()&"')")
	Response.Redirect("?")
end Sub

Sub SaveReply() '回复
	Dim oReceiver,oSender,oTitle,oContent,oIsread
	oReceiver = Request.Form("oReceiver")
	oSender = Request.Form("oSender")
	oTitle = Request.Form("oTitle")
	oContent = Request.Form("oContent")
	conn.execute("insert into [OA_mms_send] (oReceiver,oSender,oTitle,oContent,oIsread,oTime) values('"&oReceiver&"','"&oSender&"','"&oTitle&"','"&oContent&"',0,'"&Now()&"')")
	conn.execute("insert into [OA_mms_Receive] (oReceiver,oSender,oTitle,oContent,oIsread,oTime) values('"&oReceiver&"','"&oSender&"','"&oTitle&"','"&oContent&"',0,'"&Now()&"')")
	Response.Redirect("?")
end Sub

Sub InfoShow() '查看
Id = Trim(Request("Id"))
  Dim ONtitle,ONcontent,ONedittime
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From OA_mms_Receive where id="&id,conn,1,1
	oSender = rs("oSender")
	oReceiver = rs("oReceiver")
	oContent = rs("oContent")
	oTime = rs("oTime")
	oTitle = Replace(rs("oTitle"),""&L_Reply&"：","")
	rs.Close
	Set rs = Nothing
	conn.execute "UPDATE OA_mms_Receive SET oIsread='1' Where id="&id
%>
    <div class="page">
	<div class="simplebox">
            	<h1 class="titleh"><%=oTitle%> (<%=oSender%>)</h1>
					<table class="tabledata"> 
						<tbody> 
						<tr>
							<td style="padding:10px;"><%=oContent%> <BR><%=EasyCrm.FormatDate(oTime,1)%></td>
						</tr>
                    </table> 
					
				<% If mid(Session("CRM_qx"), 63, 1) = 1 Then %><BR>
				<form name="Save" action="?action=SaveReply" method="post" onSubmit="return CheckInput();">
            	<h1 class="titleh"><%=L_Top_Mms_reply%></h1>
					<table class="tabledata"> 
						<tbody> 
						<tr>
							<td style="padding:10px;"><textarea name="oContent" id="oContent" class="int" style="width:80%;"></textarea>
							<BR>
							<input name="oTitle" type="hidden" value="<%=L_Reply%>：<%=oTitle%>">
							<input name="oReceiver" type="hidden" value="<%=oSender%>">
							<input name="oSender" type="hidden" value="<%=Session("CRM_name")%>">
							<input type="submit" name="button" id="button" value="&nbsp;保 存&nbsp;" class="submit-button" />
							</td>
						</tr>
                    </table> 
				</form>
				<%end if%>
			</div>

<%
End Sub
%>
    <div class="clear"></div>
    </div>
    <!-- end page -->
<script type="text/javascript" src="js/frame.js"></script>
</body>
</html>
