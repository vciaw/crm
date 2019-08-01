<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Hetong.asp"
Session("CRM_pagenum") = PNN

	Keyword = EasyCrm.Searchcode(Request("Keyword"))
	
    If Keyword <> "" Then
		sql = sql & " and ( cid in ( select cid from client where cCompany  like '%"&Keyword&"%' ) "
		sql = sql & " or hType like '%"&Keyword&"%' "
		sql = sql & " or hState like '%"&Keyword&"%' "
		sql = sql & " or hUser like '%"&Keyword&"%' "
		sql = sql & " ) "
	End If

If Session("CRM_level") < 9 Then
	sql = sql & " And hUser In (" & arrUser & ")"
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
		<blockquote>查询条件: 公司,类型,状态(有效,无效,待审),员工</blockquote>
   	  </form>
    </div>
    <!-- end searchbox -->
    
    
    <!-- start page -->
    <div class="page">
    
    
            <!-- start list menu -->
                		
            <div class="simplebox">
            	<h1 class="titleh">合同记录</h1>
                		
           	 <ul class="list-menu">
			 
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [Hetong] where 1 = 1 "&sql&" Order By hId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Hetong] where 1 = 1 "&sql&" and hId < ( SELECT Min(hId) FROM ( SELECT TOP "&pagenum&" hId FROM [Hetong]  where 1 = 1 "&sql&" ORDER BY hId desc ) AS T ) Order By hId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(hId) As RecordSum From [Hetong] where 1 = 1 "&sql&" ",1,1)
						
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
             	<li><a href="GetUpdate.asp?action=Client&sType=View&otype=Hetong&cid=<%=rs("cId")%>"><b><%=EasyCrm.getNewItem("Client","cID",rs("cId"),"cCompany")%></b> 
				<%if rs("hUser")<>"" then%>业务员 : <%=rs("hUser")%>　<%end if%> <%if rs("hType")<>"" then%>合同分类 : <%=rs("hType")%>　<%end if%> <BR>
				<%if rs("hSdate")<>"" then%>合同期限 : <%=EasyCrm.FormatDate(rs("hSdate"),2)%>&nbsp;<%end if%> <%if rs("hEDate")<>"" then%>～&nbsp;<%=EasyCrm.FormatDate(rs("hEDate"),2)%>　<%end if%> <BR>
				<%if rs("hMoney")<>"" then%>总金额 : <%=rs("hMoney")%>　<%end if%> <%if rs("hRevenue")<>"" then%>已付款 : <%=rs("hRevenue")%>　<%end if%> <%if rs("hOwed")<>"" then%>欠款 : <%=rs("hOwed")%>　<%end if%> <%if rs("hState")<>"" then%>状态 : <%=rs("hState")%><%end if%> <%if rs("hContent")<>"" then%><BR><%=rs("hContent")%><%end if%> </a> </li>
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

            <%=pagelist("Hetong.asp?Keyword="&Keyword&"", PN,TotalPages,TotalRecords)%>
            
    <div class="clear"></div>
    </div>
    <!-- end page -->  
    
    
    
</body>
</html>
<script type="text/javascript" src="js/frame.js"></script>
<% Set EasyCrm = nothing %>