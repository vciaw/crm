<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
	'获取get值
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
%><%=Header%>
<!-- start header -->
    <div id="header">
         <a href="Index.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
         <a href="Index.asp" class="button back"><img src="img/back-button.png" width="15" height="16" alt="icon" /></a>
         <a onClick=window.location.href="javascript:window.location.reload();" class="button create"><img src="img/reload-button.png" width="15" height="16" alt="icon" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start page -->
    <div class="page">
	<div class="simplebox">

            	<h1 class="titleh">通讯录</h1>
					<table class="tabledata"> 
						<tbody> 
                        <tr> 
							<td >编号</td> 
							<td>姓名</td> 
							<td>电话</td> 
							<td>生日</td> 
                        </tr> 
						<%
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.Open "Select * From [user] order by uId asc ",conn,1,1
							Do While Not rs.BOF And Not rs.EOF
						%>
								<tr>
									<td>[<%=rs("uId")%>]</td>
									<td><%=rs("uName")%></td>
									<td><a href="tel:<%=rs("uMobile")%>"><%=rs("uMobile")%></a></td>
									<td><%=EasyCrm.FormatDate(rs("uBirthday"),2)%></td>
								</tr>
						<%
							rs.MoveNext
							Loop
							rs.Close
							Set rs = Nothing
						%>
                    </table> 
					<blockquote>提示：长按电话号码，可拨打或发短信</blockquote>
			</div>
		<%=Footer%>
          
    <div class="clear"></div>
    </div>
    <!-- end page -->
<script type="text/javascript" src="js/frame.js"></script>
</body>
</html>
