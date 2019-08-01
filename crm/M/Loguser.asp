<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
If Session("CRM_level") < 9 Then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
end if

action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Loguser.asp"
Session("CRM_pagenum") = PNN

%>
<%=Header%>
    
<!-- start header -->
    <div id="header">
         <a href="System.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    
    
    <!-- start page -->
    <div class="page">
    
    
            <!-- start list menu -->
                		
            <div class="simplebox">
            	<h1 class="titleh">操作记录</h1>
                                <table class="tabledata"> 
                                	<thead> 
                                		<tr> 
                                			<th>编号</th> 
                                			<th>账户</th> 
                                			<th>时间</th> 
                                			<th>登录IP</th> 
                                		</tr> 
                                	</thead> 
                                    <tbody> 
	<%
	PN = CLng(ABS(Request("PN")))
	If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
	intPageSize = DataPageSize
	pagenum = intPageSize*(PN-1)
	Set rs = Server.CreateObject("ADODB.Recordset")
	IF PN=1 THEN
	rs.Open "Select top "&intPageSize&" * From [userlog] where 1=1 "&sql&" Order By id desc ",conn,1,1 
	ELSE
	rs.Open "Select top "&intPageSize&" * From [userlog] where 1=1 "&sql&" and id < ( SELECT Min(id) FROM ( SELECT TOP "&pagenum&" id FROM [userlog] where 1=1 "&sql&" ORDER BY id desc ) AS T ) Order By id desc ",conn,1,1
	END IF
	Set Rsstr=conn.Execute("Select count(id) As RecordSum From [userlog] where 1=1 "&sql&" ",1,1)
	TotalRecords=Rsstr("RecordSum") 
	if Int(TotalRecords/intPageSize)=TotalRecords/intPageSize then
	TotalPages=TotalRecords/intPageSize
	else
	TotalPages=Int(TotalRecords/intPageSize)+1
	end if
	Rsstr.Close:Set Rsstr=Nothing
	If PN > TotalPages Then PN = TotalPages
	If rs.RecordCount > 0 Then
	Do While Not rs.BOF And Not rs.EOF
	%>
                                        <tr> 
                                        	<td><%=rs("id")%></td> 
                                        	<td><%=rs("olname")%></td> 
                                        	<td><%=rs("olstarttime")%></td> 
                                        	<td><%=rs("olip")%></td> 
                                        </tr> 
	<%
	rs.MoveNext
	Loop
	else
	%>
							<tr><td class="td_l_l" colspan="4">无记录！</td></tr>
	<%
	end if
	rs.Close
	Set rs = Nothing
	%>


                                    </tbody> 
                                </table>
             
             </div>
             
            <!-- end list menu -->

            <!-- start top button -->
            <%=pagelist("Loguser.asp", PN,TotalPages,TotalRecords)%>
            <!-- end top button -->
            
	<%=Footer%>
    
    <div class="clear"></div>
    </div>
    <!-- end page -->  
</body>
</html>
<script type="text/javascript" src="js/frame.js"></script>
<% Set EasyCrm = nothing %>