<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
	'��ȡgetֵ
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
%><%=Header%>
<!-- start header -->
    <div id="header">
         <a href="index.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
         <a onClick=window.location.href="javascript:history.back();" class="button back"><img src="img/back-button.png" width="15" height="16" alt="icon" /></a>
         <a onClick=window.location.href="javascript:window.location.reload();" class="button create"><img src="img/reload-button.png" width="15" height="16" alt="icon" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start page -->

<%
Select Case action
Case "InfoShow" '�鿴
    Call InfoShow()
Case Else
    Call Main()
End Select

Sub Main()
%>
    <div class="page">
	<div class="simplebox">
            	<h1 class="titleh">�ڲ�����</h1>
					<table class="tabledata"> 
						<tbody> 
						<thead> 
                        <tr> 
							<th>���</th> 
							<th>����</th> 
                        </tr> 
						</thead> 
						<%
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.Open "Select * From [OA_Notice] order by ONid desc ",conn,1,1
							Do While Not rs.BOF And Not rs.EOF
						%>
								<tr>
									<td>[<%=rs("ONclass")%>]</td>
									<td><a href="?action=InfoShow&Id=<%=rs("ONId")%>"><%=rs("ONtitle")%></a></td>
								</tr>
						<%
							rs.MoveNext
							Loop
							rs.Close
							Set rs = Nothing
						%>
                    </table> 
					<blockquote>�ݲ��ṩ�༭��ɾ������</blockquote>
			</div>
		<%=Footer%>
            
<%
end Sub

Sub InfoShow() '�鿴
Id = Trim(Request("Id"))
%>
    <div class="page">
	<div class="simplebox">
            	<h1 class="titleh"><%=EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONclass")%>:<%=EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONtitle")%></h1>
					<table class="tabledata"> 
						<tbody> 
						<tr>
							<td style="padding:10px;"><%=EasyCrm.htmlEncode3(EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONcontent"))%></td>
						</tr>
                    </table> 
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
