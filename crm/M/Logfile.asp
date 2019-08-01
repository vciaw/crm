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
Session("CRM_thispage") = "Logfile.asp"
Session("CRM_pagenum") = PNN

	Keyword = EasyCrm.Searchcode(Request("Keyword"))
	
    If Keyword <> "" Then
	End If
%>
<%=Header%>
    
<!-- start header -->
    <div id="header">
         <a href="System.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start searchbox -->
    <div class="searchbox">
   	  <form id="form1" name="form1" method="get" action="?subAction=searchItem">
      	<input type="text" name="Keyword" id="Keyword" class="txtbox" />
   	  </form>
    </div>
    <!-- end searchbox -->
    
    
    <!-- start page -->
    <div class="page">
    
    
            <!-- start list menu -->
                		
            <div class="simplebox">
            	<h1 class="titleh">操作记录</h1>
                                <table class="tabledata"> 
                                	<thead> 
                                		<tr> 
                                			<th>时间</th> 
                                			<th>客户编号</th> 
                                			<th>行为</th> 
                                			<th>数据表</th> 
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
	rs.Open "Select top "&intPageSize&" * From [Logfile] where 1=1 "&sql&" Order By lid desc ",conn,1,1 
	ELSE
	rs.Open "Select top "&intPageSize&" * From [Logfile] where 1=1 "&sql&" and lid < ( SELECT Min(lid) FROM ( SELECT TOP "&pagenum&" lid FROM [Logfile] where 1=1 "&sql&" ORDER BY lid desc ) AS T ) Order By lid desc ",conn,1,1
	END IF
	Set Rsstr=conn.Execute("Select count(lid) As RecordSum From [Logfile] where 1=1 "&sql&" ",1,1)
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
                                        <tr onClick=window.location.href="GetUpdate.asp?action=Client&sType=View&cid=<%=rs("lcId")%>"> 
                                        	<td><%=rs("lTime")%></td> 
                                        	<td><%=rs("lcId")%></td> 
                                        	<td><%=rs("lAction")%></td> 
                                        	<td><%=rs("lClass")%></td> 
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
            <%=pagelist("Logfile.asp", PN,TotalPages,TotalRecords)%>
            <!-- end top button -->
            
	<%=Footer%>
    
    <div class="clear"></div>
    </div>
    <!-- end page -->  
</body>
</html>
<script type="text/javascript" src="js/frame.js"></script>
<% Set EasyCrm = nothing %>