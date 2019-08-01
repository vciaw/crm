<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Records.asp"
Session("CRM_pagenum") = PNN

	Keyword = EasyCrm.Searchcode(Request("Keyword"))
	
    If Keyword <> "" Then
		sql = sql & " and ( cid in ( select cid from client where cCompany  like '%"&Keyword&"%' ) "
		sql = sql & " or rState like '%"&Keyword&"%' "
		sql = sql & " or rType like '%"&Keyword&"%' "
		sql = sql & " or rUser like '%"&Keyword&"%' "
		sql = sql & " ) "
	End If

If Session("CRM_level") < 9 Then
	sql = sql & " And rUser In (" & arrUser & ")"
end if
%>
<%=Header%>
    
<!-- start header -->
    <div id="header">
         <a href="Index.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
         	<a href="Index.asp" class="button back"><img src="img/back-button.png" width="15" height="16" alt="icon" /></a>
         	<a href="#" class="button rightbox"><img src="img/search.png" width="16" height="16" alt="icon"/></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start searchbox -->
    <div class="searchbox">
   	  <form id="form1" name="form1" method="get" action="?subAction=searchItem">
      	<input type="text" name="Keyword" id="Keyword" class="txtbox" value="<%=Keyword%>" />
		<blockquote>查询条件: 公司,类型,进度,员工</blockquote>
   	  </form>
    </div>
    <!-- end searchbox -->
    
    
    <!-- start page -->
    <div class="page">
    
    
            <!-- start list menu -->
                		
            <div class="simplebox">
            	<h1 class="titleh">跟单记录</h1>
                		
           	 <ul class="list-menu">
			 
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [Records] where 1 = 1 "&sql&" Order By rId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Records] where 1 = 1 "&sql&" and rId < ( SELECT Min(rId) FROM ( SELECT TOP "&pagenum&" rId FROM [Records]  where 1 = 1 "&sql&" ORDER BY rId desc ) AS T ) Order By rId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(rId) As RecordSum From [Records] where 1 = 1 "&sql&" ",1,1)
						
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
             	<li><a href="GetUpdate.asp?action=Client&sType=View&otype=Records&cid=<%=rs("cId")%>"><b><%=EasyCrm.getNewItem("Client","cID",rs("cId"),"cCompany")%></b> <%=EasyCrm.FormatDate(rs("rTime"),2)%>　<%if rs("rUser")<>"" then%><%=rs("rUser")%>　<%end if%> <%if rs("rType")<>"" then%><%=rs("rType")%>　<%end if%> <%if rs("rState")<>"" then%><%=rs("rState")%>　<%end if%> <%if rs("rContent")<>"" then%><BR><%=rs("rContent")%><%end if%> </a> </li>
							<%
							rs.MoveNext
							Loop
							else
							%>
             	<li><i>无记录！</i></li>
							<%
							end if
							rs.Close
							Set rs = Nothing
							%>
             </ul>
             
             </div>
             
            <!-- end list menu -->

            <%=pagelist("Records.asp?Keyword="&Keyword&"", PN,TotalPages,TotalRecords)%>
            
    <div class="clear"></div>
    </div>
    <!-- end page -->  
    
    
    
</body>
</html>
<script type="text/javascript" src="js/frame.js"></script>
<% Set EasyCrm = nothing %>