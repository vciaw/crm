<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
	'获取get值
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	oClass=Request.QueryString("oClass")
	oIsread=Request.QueryString("oIsread")
	Id = Trim(Request("Id"))
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
                   	  <label class="st-label"><font color="#FF0000">*</font> 上级领导</label>
						<select name='oSendto'>
							<%
								Set rsm = Server.CreateObject("ADODB.Recordset")
								rsm.Open "Select * From [user] where uGroup="&Session("CRM_group")&" and ulevel > "&Session("CRM_level")&" or ulevel=9 ",conn,1,1
								Do While Not rsm.BOF And Not rsm.EOF
							%>
							<option value="<%=rsm("uName")%>" <%if rsm("ulevel")=9 then%>selected<%end if%> > <%=rsm("uName")%></option>
							<%
								rsm.MoveNext
								Loop
								rsm.Close
								Set rsm = Nothing
							%>
						</select>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><font color="#FF0000">*</font> <%=L_Report_oClass%></label>
						<input name="oClass" type="radio" class="noborder" value="<%=L_Ribao%>" checked> <%=L_Ribao%>
						<input name="oClass" type="radio" class="noborder" value="<%=L_Zhoubao%>"> <%=L_Zhoubao%>
						<input name="oClass" type="radio" class="noborder" value="<%=L_Yuebao%>"> <%=L_Yuebao%>
						<input name="oClass" type="radio" class="noborder" value="<%=L_Jibao%>"> <%=L_Jibao%>
						<input name="oClass" type="radio" class="noborder" value="<%=L_Nianbao%>"> <%=L_Nianbao%>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><%=L_Report_oReport%></label>
						<textarea name="oReport" id="oReport" class="int" style="width:80%;"></textarea>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label"><%=L_Report_oPlan%></label>
						<textarea name="oPlan" id="oPlan" class="int" style="width:80%;"></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="oUser" type="hidden" value="<%=Session("CRM_name")%>">
                    <input type="submit" name="button" id="button" value="&nbsp;保 存&nbsp;" class="submit-button" />
                    </div>
                </form>
                </div>
				
            	<h1 class="titleh">工作报告</h1>
					<table class="tabledata"> 
						<col width="70" >
						<tbody> 
                        <tr> 
							<td>分类</td> 
							<td>标题</td> 
                        </tr> 
						<%
								If Session("CRM_level") = 9 Then									'管理员管理所有的
								sql = sql & " "
								elseIf Session("CRM_level") < 9 and Session("CRM_level") > 1 Then 	'部门经理可以看到别人提交给自己的，和自己提交的
								sql = sql & " and ( oSendto like '%"&Session("CRM_name")&"%' or oUser = '"&Session("CRM_name")&"' ) "
								else								
								sql = sql & " and oUser = '"&Session("CRM_name")&"' " 				'普通员工只能看到自己提交的报告
								end if
								
								if oClass <>"" then
								sql = sql & " and oClass ='"&oClass&"'"
								end if
								
								if oIsread <>"" then
								sql = sql & " and oIsread ="&oIsread&""
								end if
								
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.Open "Select * From [OA_Report] where 1=1 "&sql&" order by Id desc ",conn,1,1
							Do While Not rs.BOF And Not rs.EOF
						%>
								<tr>
									<td>[<%=rs("oClass")%>]</td>
									<td><a href="?action=InfoShow&Id=<%=rs("Id")%>"><%=rs("oTitle")%></a> <%if rs("oIsread") = 1 then%>[已读]<%end if%></td>
								</tr>
						<%
							rs.MoveNext
							Loop
							rs.Close
							Set rs = Nothing
						%>
                    </table> 
					<blockquote>可查看、批注，暂不提供删除权限</blockquote>
			</div>
		<%=Footer%>
            
<%
end Sub

Sub SaveAdd() '添加
	oSendto = Request.Form("oSendto")
	oClass = Request.Form("oClass")
	oReport = EasyCrm.htmlEncode2(Request.Form("oReport"))
	oPlan = EasyCrm.htmlEncode2(Request.Form("oPlan"))
	oUser = Request.Form("oUser")
		
	conn.execute("insert into [OA_Report] (oSendto,oClass,oTitle,oReport,oPlan,oUser,oIsread,oTime) values('"&oSendto&"','"&oClass&"','"&oUser&" "&L_whoswork&oClass&"','"&oReport&"','"&oPlan&"','"&oUser&"',0,'"&Now()&"')")
	Response.Redirect("?")
end Sub

Sub SaveReply() '回复
	id = Request("id")
	oReply = Request.Form("oReply")
	oReplyOld = Request.Form("oReplyOld")
	conn.execute "UPDATE [OA_Report] SET oReply='"&oReply&"' Where id="&id
	
	'如果批复内容改变，则站内信通知提交者
	if oReply<>oReplyOld then
	conn.execute("insert into [OA_mms_Receive] (oReceiver,oSender,oTitle,oContent,oIsread,oTime) values('"&EasyCrm.getNewItem("OA_Report","ID",""&Id&"","oUser")&"','系统消息','您提交的工作报告已批阅！','请进入【工作报告】栏目查看详情！',0,'"&Now()&"')")
	end if
	Response.Redirect("?")
end Sub

Sub InfoShow() '查看
	'接收报告者阅读后，工作报告设为已读
	if Session("CRM_level") = 9 then
	conn.execute "UPDATE OA_Report SET oIsread='1' Where id="&id
	else
	conn.execute "UPDATE OA_Report SET oIsread='1' Where oSendto like '%"&Session("CRM_name")&"%' and id="&id
	end if
	
  Dim ONtitle,ONcontent,ONedittime
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From OA_Report where id="&id,conn,1,1
	oReport = rs("oReport")
	oPlan = rs("oPlan")
	oReply = rs("oReply")
	oUser = rs("oUser")
	oTime = rs("oTime")
	rs.Close
	Set rs = Nothing
%>
    <div class="page">
	<div class="simplebox">
            	<h1 class="titleh"><%=oUser%> 的 <%=L_Report_oReport%></h1>
					<table class="tabledata"> 
						<tbody> 
						<tr>
							<td style="padding:10px;"><%=oReport%></td>
						</tr>
                    </table> 
				<BR>
            	<h1 class="titleh"><%=L_Report_oPlan%></h1>
					<table class="tabledata"> 
						<tbody> 
						<tr>
							<td style="padding:10px;"><%=oPlan%> <BR><%=EasyCrm.FormatDate(oTime,1)%></td>
						</tr>
                    </table> 
					
				<% If mid(Session("CRM_qx"), 68, 1) = 1 Then %><BR>
				<form name="Save" action="?action=SaveReply" method="post" onSubmit="return CheckInput();">
            	<h1 class="titleh"><%=L_Report_oReply%></h1>
					<table class="tabledata"> 
						<tbody> 
						<tr>
							<td style="padding:10px;"><textarea name="oReply" id="oReply" style="width:80%;"><%=oReply%></textarea> 
							<BR>
							<input name="id" type="hidden" value="<%=id%>">
							<input name="oReplyOld" type="hidden" value="<%=oReply%>">
							<input type="submit" name="button" id="button" value="&nbsp;保 存&nbsp;" class="submit-button" />
							</td>
						</tr>
                    </table> 
				</form>
				<%else%>
				<%if oReply<>"" then%>
				<BR>
            	<h1 class="titleh"><%=L_Report_oPlan%></h1>
					<table class="tabledata"> 
						<tbody> 
						<tr>
							<td style="padding:10px;"><%=oReply%></td>
						</tr>
                    </table> 
				<%end if%>
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
<% Set EasyCrm = nothing %>