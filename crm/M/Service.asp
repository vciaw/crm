<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Service.asp"
Session("CRM_pagenum") = PNN

	Keyword = EasyCrm.Searchcode(Request("Keyword"))
	
	if Keyword = "δ����" then
		sSolve = "0" 
	elseif Keyword = "�Ѵ���" then
		sSolve = "1" 
	end if

    If Keyword <> "" Then
		sql = sql & " and ( cid in ( select cid from client where cCompany  like '%"&Keyword&"%' ) "
		sql = sql & " or sTitle like '%"&Keyword&"%' "
		sql = sql & " or sLinkman like '%"&Keyword&"%' "
		sql = sql & " or sType like '%"&Keyword&"%' "
		sql = sql & " or sSolve = '"&sSolve&"' "
		sql = sql & " or sUser like '%"&Keyword&"%' "
		sql = sql & " ) "
	End If

If Session("CRM_level") < 9 Then
	sql = sql & " And sUser In (" & arrUser & ")"
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
		<blockquote>��ѯ����: ��˾,����,����,״̬(δ����,�Ѵ���),Ա��</blockquote>
   	  </form>
    </div>
    <!-- end searchbox -->
    
    
    <!-- start page -->
    <div class="page">
    
    
            <!-- start list menu -->
                		
            <div class="simplebox">
            	<h1 class="titleh">�ۺ��¼</h1>
                		
           	 <ul class="list-menu">
			 
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [Service] where 1 = 1 "&sql&" Order By sId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Service] where 1 = 1 "&sql&" and sId < ( SELECT Min(sId) FROM ( SELECT TOP "&pagenum&" sId FROM [Service]  where 1 = 1 "&sql&" ORDER BY sId desc ) AS T ) Order By sId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(sId) As RecordSum From [Service] where 1 = 1 "&sql&" ",1,1)
						
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
             	<li><a href="GetUpdate.asp?action=Client&sType=View&otype=Service&cid=<%=rs("cId")%>"><b><%=EasyCrm.getNewItem("Client","cID",rs("cId"),"cCompany")%></b> 
				<%if rs("sUser")<>"" then%>ҵ��Ա : <%=rs("sUser")%>��<%end if%> <%if rs("sType")<>"" then%>�������� : <%=rs("sType")%>��<%end if%> <BR>
				<%if rs("sTitle")<>"" then%>�������� : <%=rs("sTitle")%>��<%end if%> <%if rs("sLinkman")<>"" then%>��ϵ�� : <%=rs("sLinkman")%>��<%end if%> <BR>
				<%if rs("sSDate")<>"" then%>�������� : <%=EasyCrm.FormatDate(rs("sSDate"),2)%>��<%end if%> <%if rs("sEDate")<>"" then%>�������� : <%=EasyCrm.FormatDate(rs("sEDate"),2)%>��<%end if%> 
				<%if rs("sContent")<>"" then%><BR>�������� : <%=rs("sContent")%><%end if%><BR>
				 <%if rs("sSolve")<>"" then%>״̬ : <%if rs("sSolve") = 0 then%><font color=red>δ����</font><%elseif rs("sSolve") = 1 then%><font color=green>�Ѵ���</font><%end if%><%end if%> <%if rs("sInfo")<>"" then%><BR>������ : <%=rs("sInfo")%><%end if%> </a> </li>
							<%
							rs.MoveNext
							Loop
							else
							%>
             	<li><i>�޼�¼��</i></li>
							<%
							end if
							rs.Close
							Set rs = Nothing
							%>
             </ul>
             
             </div>
             
            <!-- end list menu -->

            <%=pagelist("Service.asp?Keyword="&Keyword&"", PN,TotalPages,TotalRecords)%>
            
    <div class="clear"></div>
    </div>
    <!-- end page -->  
    
    
    
</body>
</html>
<script type="text/javascript" src="js/frame.js"></script>
<% Set EasyCrm = nothing %>