<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Expense.asp"
Session("CRM_pagenum") = PNN

	Keyword = EasyCrm.Searchcode(Request("Keyword"))
	
	if Keyword = "支出" then
		eOutIn = "0" 
	elseif Keyword = "收入" then
		eOutIn = "1" 
	end if

    If Keyword <> "" Then
		sql = sql & " and ( cid in ( select cid from client where cCompany  like '%"&Keyword&"%' ) "
		sql = sql & " or eType like '%"&Keyword&"%' "
		sql = sql & " or eOutIn = '"&eOutIn&"' "
		sql = sql & " or eUser like '%"&Keyword&"%' "
		sql = sql & " ) "
	End If

If Session("CRM_level") < 9 Then
	sql = sql & " And eUser In (" & arrUser & ")"
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
		<blockquote>查询条件: 公司,类型(收入,支出),费用类别,员工</blockquote>
   	  </form>
    </div>
    <!-- end searchbox -->
    
    
    <!-- start page -->
    <div class="page">
    
    
            <!-- start list menu -->
                		
            <div class="simplebox">
            	<h1 class="titleh">费用记录</h1>
                		
           	 <ul class="list-menu">
			 
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [Expense] where 1 = 1 "&sql&" Order By eId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Expense] where 1 = 1 "&sql&" and eId < ( SELECT Min(eId) FROM ( SELECT TOP "&pagenum&" eId FROM [Expense]  where 1 = 1 "&sql&" ORDER BY eId desc ) AS T ) Order By eId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(eId) As RecordSum From [Expense] where 1 = 1 "&sql&" ",1,1)
						
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
             	<li><a href="GetUpdate.asp?action=Client&sType=View&otype=Expense&cid=<%=rs("cId")%>"><b><%=EasyCrm.getNewItem("Client","cID",rs("cId"),"cCompany")%></b>
				<%if rs("eUser")<>"" then%>业务员 : <%=rs("eUser")%>　<%end if%> <%if rs("eDate")<>"" then%>收支日期 : <%=EasyCrm.FormatDate(rs("eDate"),2)%><%end if%> <BR>
				<%if rs("eOutIn")<>"" then%>收支类型 : <%if rs("eOutIn") = 1 then %><font color=green>收入</font><%else%><font color=red>支出</font><%end if%>　<%end if%> <%if rs("eType")<>"" then%>费用类别 : <%=rs("eType")%><%end if%> <BR>
				<%if rs("eMoney")<>"" then%>总金额 : <%=rs("eMoney")%>　<%end if%> <%if rs("eTime")<>"" then%>录入日期 : <%=EasyCrm.FormatDate(rs("eTime"),2)%><%end if%> <%if rs("eContent")<>"" then%><BR>详情备注 : <%=rs("eContent")%><%end if%> </a></li>
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

            <%=pagelist("Expense.asp?Keyword="&Keyword&"", PN,TotalPages,TotalRecords)%>
            
    <div class="clear"></div>
    </div>
    <!-- end page -->  
    
    
    
</body>
</html>
<script type="text/javascript" src="js/frame.js"></script>
<% Set EasyCrm = nothing %>