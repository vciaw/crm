<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
action = Trim(Request("action"))
oType = Trim(Request("oType"))
subAction = Trim(Request("subAction"))
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
Session("CRM_thispage") = "Order.asp"
Session("CRM_pagenum") = PNN

	Keyword = EasyCrm.Searchcode(Request("Keyword"))
	
	if Keyword = "δ����" then
		oState = "0" 
	elseif Keyword = "������" then
		oState = "1" 
	elseif Keyword = "�����" then
		oState = "2" 
	elseif Keyword = "��ȡ��" then
		oState = "3" 
	end if
	
    If Keyword <> "" Then
		sql = sql & " and ( cid in ( select cid from client where cCompany  like '%"&Keyword&"%' ) "
		sql = sql & " or oLinkman like '%"&Keyword&"%' "
		sql = sql & " or oState = '"&oState&"' "
		sql = sql & " or oUser like '%"&Keyword&"%' "
		sql = sql & " ) "
	End If

If Session("CRM_level") < 9 Then
	sql = sql & " And oUser In (" & arrUser & ")"
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
		<blockquote>��ѯ����: ��˾,״̬(δ����,������,�����),Ա��</blockquote>
   	  </form>
    </div>
    <!-- end searchbox -->
    
    
    <!-- start page -->
    <div class="page">
    
    
            <!-- start list menu -->
                		
            <div class="simplebox">
            	<h1 class="titleh">������¼</h1>
                		
           	 <ul class="list-menu">
			 
							<%
							PN = CLng(ABS(Request("PN")))
							If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
							intPageSize = DataPageSize
							pagenum = intPageSize*(PN-1)
							Set rs = Server.CreateObject("ADODB.Recordset")
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [Order] where 1 = 1 "&sql&" Order By oId Desc",conn,1,1
							ELSE
							rs.Open "Select top "&intPageSize&" * From [Order] where 1 = 1 "&sql&" and oId < ( SELECT Min(oId) FROM ( SELECT TOP "&pagenum&" oId FROM [Order]  where 1 = 1 "&sql&" ORDER BY oId desc ) AS T ) Order By oId Desc ",conn,1,1
							END IF
							Set Rsstr=conn.Execute("Select count(oId) As RecordSum From [Order] where 1 = 1 "&sql&" ",1,1)
						
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
             	<li><a href="GetUpdate.asp?action=Client&sType=View&otype=Order&cid=<%=rs("cId")%>"><b><%=EasyCrm.getNewItem("Client","cID",rs("cId"),"cCompany")%></b> 
				<%if rs("oUser")<>"" then%>ҵ��Ա : <%=rs("oUser")%>��<%end if%> <%if rs("oLinkman")<>"" then%>��ϵ�� : <%=rs("oLinkman")%>��<%end if%> <BR>
				<%if rs("oSDate")<>"" then%>�������� : <%=EasyCrm.FormatDate(rs("oSDate"),2)%>&nbsp;<%end if%> <%if rs("oEDate")<>"" then%>��&nbsp;<%=EasyCrm.FormatDate(rs("oEDate"),2)%>��<%end if%> <BR>
				<%if rs("oDeposit")<>"" then%>Ԥ���� : <%=rs("oDeposit")%>��<%end if%> <%if rs("oMoney")<>"" then%>�ܽ�� : <%=rs("oMoney")%>��<%end if%> <%if rs("oState")<>"" then%>״̬ : <%if rs("oState") = 0 then%><font color=red>δ����</font><%elseif rs("oState") = 1 then%><font color=blue>������</font><%elseif rs("oState") = 2 then%><font color=green>�����</font><%elseif rs("oState") = 3 then%><font color=DoderBlue>��ȡ��</font><%end if%><%end if%> <%if rs("oContent")<>"" then%><BR><%=rs("oContent")%><%end if%> </a> </li>
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

            <%=pagelist("Order.asp?Keyword="&Keyword&"", PN,TotalPages,TotalRecords)%>
            
    <div class="clear"></div>
    </div>
    <!-- end page -->  
    
    
    
</body>
</html>
<script type="text/javascript" src="js/frame.js"></script>
<% Set EasyCrm = nothing %>