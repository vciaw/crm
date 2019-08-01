<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Recycler.asp"
Session("CRM_pagenum") = PNN

	Keyword = EasyCrm.Searchcode(Request("Keyword"))
	
    If Keyword <> "" Then
		sql = sql & " and ( cCompany like '%"&Keyword&"%' "
		sql = sql & " or cLinkman like '%"&Keyword&"%' "
		sql = sql & " or cMobile like '%"&Keyword&"%' "
		sql = sql & " or cType like '%"&Keyword&"%' "
		sql = sql & " or cSource like '%"&Keyword&"%' "
		sql = sql & " or cStart like '%"&Keyword&"%' "
		sql = sql & " or cTrade like '%"&Keyword&"%' "
		sql = sql & " ) "
	End If

If Session("CRM_level") < 9 Then
	sql = sql & " And cUser In (" & arrUser & ")"
end if
%>
<%=Header%>
    
<!-- start header -->
    <div id="header">
         <a href="Index.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
         	<a href="Index.asp" class="button back"><img src="img/back-button.png" width="15" height="16" alt="icon" /></a>
         	<a href="#" class="button search"><img src="img/search.png" width="16" height="16" alt="icon"/></a>
         	<a href="GetUpdate.asp?action=Client&sType=Add" class="button create"><img src="img/create.png" width="16" height="16" alt="icon"/></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start searchbox -->
    <div class="searchbox">
   	  <form id="form1" name="form1" method="get" action="?subAction=searchItem">
      	<input type="text" name="Keyword" id="Keyword" class="txtbox" value="<%=Keyword%>" />
   	  </form>
    </div>
    <!-- end searchbox -->
    
    
    <!-- start page -->
    <div class="page">
    
    
            <!-- start list menu -->
                		
            <div class="simplebox">
            	<h1 class="titleh">系统公海</h1>
                		
           	 <ul class="list-menu">
			 
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [client] where cYn = 0 "&sql&" Order By cId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [client] where cYn = 0 "&sql&" and cid < ( SELECT Min(cid) FROM ( SELECT TOP "&pagenum&" cid FROM [client]  where cYn = 0 "&sql&" ORDER BY cId desc ) AS T ) Order By cId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(cid) As RecordSum From [client] where cYn = 0 "&sql&" ",1,1)
						
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
             	<li><a href="GetUpdate.asp?action=Client&sType=View&cid=<%=rs("cId")%>"><b><%=rs("cCompany")%></b></a></li>
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

            <%=pagelist("Recycler.asp?Keyword="&Keyword&"", PN,TotalPages,TotalRecords)%>
            
    <div class="clear"></div>
    </div>
    <!-- end page -->  
    
    
    
</body>
</html>
<script type="text/javascript" src="js/frame.js"></script>
<% Set EasyCrm = nothing %>