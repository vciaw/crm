<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%=Header%>
    
<!-- start header -->
    <div id="header">
         <a href="#"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
         	<a href="GetUpdate.asp?action=Client&sType=Add" class="button create"><img src="img/create.png" width="16" height="16" alt="icon"/></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    
    
    <!-- start page -->
    <div class="page">
    
    		
            <!-- start profile box -->
            <div class="profilebox">
            	<img src="img/avatar.png" width="19" height="20" alt="avatar" class="avatar"/> 欢迎 <b><%=Session("CRM_name")%></b> 登录系统
                <a href="Logout.asp" class="logout" title="退出">退出</a>
                <div class="clear"></div>
            </div>
            <!-- end profile box -->
            <!-- start menu -->
			<%
			CountMessages = EasyCrm.getCountItem("OA_mms_Receive","id","idstr"," and ( oIsread = 0 and oReceiver = '"&Session("CRM_name")&"' and oAttime is null ) or ( oAttime is not null and oAttime < getdate()+ 0.007 and oIsread = 0 and oReceiver = '"&Session("CRM_name")&"' ) ")
			If Session("CRM_level") = 9 Then
			CountReport = EasyCrm.getCountItem("OA_Report","id","idstr"," and oIsread = 0 ")
			elseIf Session("CRM_level") < 9 and Session("CRM_level") > 1 Then
			CountReport = EasyCrm.getCountItem("OA_Report","id","idstr"," and oIsread = 0 and ( oSendto like '%"&Session("CRM_name")&"%' or oUser = '"&Session("CRM_name")&"' ) ")
			else
			CountReport = EasyCrm.getCountItem("OA_Report","id","idstr"," and oIsread = 0 and oUser = '"&Session("CRM_name")&"' ")
			end if
			
			%>
           	 <ul id="menu">
             	<li><a href="Listall.asp"><img src="img/icons/listall.png" width="24" height="24" alt="icon" class="m-icon"/><b>客户列表</b></a></li>
             	<li><a href="Records.asp"><img src="img/icons/Records.png" width="24" height="24" alt="icon" class="m-icon"/><b>跟单管理</b></a></li>
             	<li><a href="Order.asp"><img src="img/icons/Order.png" width="24" height="24" alt="icon" class="m-icon"/><b>订单管理</b></a></li>
             	<li><a href="Hetong.asp"><img src="img/icons/Hetong.png" width="24" height="24" alt="icon" class="m-icon"/><b>合同管理</b></a></li>
             	<li><a href="Service.asp"><img src="img/icons/Service.png" width="24" height="24" alt="icon" class="m-icon"/><b>售后管理</b></a></li>
             	<li><a href="Expense.asp"><img src="img/icons/Expense.png" width="24" height="24" alt="icon" class="m-icon"/><b>费用管理</b></a></li>
             	<li><a href="Recycler.asp"><img src="img/icons/delete.png" width="24" height="24" alt="icon" class="m-icon"/><b>系统公海</b></a></li>
             	<li><a href="Notice.asp"><img src="img/icons/Notice.png" width="24" height="24" alt="icon" class="m-icon"/><b>内部公文</b></a></li>
             	<li><a href="Message.asp"><img src="img/icons/messages.png" width="24" height="24" alt="icon" class="m-icon"/><b>站内信</b></a> 
				<%if CountMessages > 0 then%><p><span class="red"><b> <%=CountMessages%></b></span></p><%end if%></li>
             	<li><a href="Report.asp"><img src="img/icons/Report.png" width="24" height="24" alt="icon" class="m-icon"/><b>工作报告</b></a>
				<%if CountReport > 0 then%><p><span class="red"><b> <%=CountReport%></b></span></p><%end if%></li>
             	<li><a href="Contact.asp"><img src="img/icons/content.png" width="24" height="24" alt="icon" class="m-icon"/><b>通讯录</b></a></li>
             	<li><a href="Statistics.asp"><img src="img/icons/graph.png" width="24" height="24" alt="icon" class="m-icon"/><b>数据统计</b></a></li>
             	<li><a href="System.asp"><img src="img/icons/setting.png" width="24" height="24" alt="icon" class="m-icon"/><b>系统设置</b></a></li>
             </ul>
            <!-- end menu -->
    <div class="clear"></div>
    </div>
    <!-- end page -->
<script type="text/javascript" src="js/frame.js"></script>
</body>
</html>
<% Set EasyCrm = nothing %>