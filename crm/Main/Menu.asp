<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
Response.Buffer = true
Response.Expires = -10000
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<title><%=title%></title>
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="<%=SiteUrl&skinurl%>Js/common.js"></script>
<script type="text/javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script type="text/javascript" src="<%=SiteUrl&skinurl%>Js/sidebar.js"></script>
<!--[if IE 6]>
<script type="text/javascript" src="<%=SiteUrl&skinurl%>Js/fixpng.js"></script>
<script>DD_belatedPNG.fix('#sidebar_main,#sidebar_listall,#sidebar_Company,#sidebar_Statistics,#sidebar_drawer,#sidebar_setting,#sidebar_user,#sidebar_log');</script>
<![endif]-->
</head>
<body id="sidebar_page">
<div class="wrap">
    <div class="cotainer">
        <div id="sidebar">
        <div class="con">        
<%
select case request("action")
case "":
call Main()
case "index":
call index()
case "Listall":
call Listall()
case "Company":
call Company()
case "Plugin":
call Plugin()
case "System":
call System()
case "Help":
call Help()
end select
%>

<%sub Main() %>
<%if Session("CRM_account")<>"" then%>
        <h2 id="sidebar_main"><%=lmquick%></h2>
        <ul>
          <li><a href='../main/Main.asp' target='main' >系统首页</a></li>
		<%
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [QuickMenu] where QuickYN = 1 Order By Sort asc,Id asc",conn,1,1
		Do While Not rs.BOF And Not rs.EOF
		%>
          <li><a href='<%=rs("Url")%>' target='main' ><%=rs("Title")%></a></li>
		<%
		rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		%>
        </ul>
		
		<%IF EasyCrm.getCountItem("Plugin","Id","Idstr"," and pYn=1 and QuickYN = 1")>0 then %>
        <h2 id="sidebar_plugin"><%=lmgncj%></h2>
        <ul>
		<%
		Set rsplugin = Server.CreateObject("ADODB.Recordset")
		rsplugin.Open "Select * From [Plugin] where pYn=1 and QuickYN = 1 Order By pSort asc,Id asc",conn,1,1
		Do While Not rsplugin.BOF And Not rsplugin.EOF
		%>
          <li><a href='../plugin/<%=rsplugin("pUrl")%>/' target='main' ><%=rsplugin("pTitle")%></a></li>
		<%
		rsplugin.MoveNext
		Loop
		rsplugin.Close
		Set rsplugin = Nothing
		%>
        </ul>
		<%end if%>
<%end if%>
<% end sub %>

<%sub Listall() %>
<%if Session("CRM_account")<>"" then%>
        <h2 id="sidebar_listall"><%=lmliall%></h2>
        <ul>
		 <li><a href='../main/Customer.asp' target='main' >所有客户</a></li>
         <li><a href='../main/Listall.asp' target='main' >房产管理</a></li>
		<% If mid(Session("CRM_qx"), 26, 1) = 1 Then %>
          <li><a href='../main/Records.asp' target='main' >跟单管理</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 31, 1) = 1 Then %>
          <li><a href='../main/Order.asp' target='main' >订单管理</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 36, 1) = 1 Then %>
          <li><a href='../main/Hetong.asp' target='main' >合同管理</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 41, 1) = 1 Then %>
          <li><a href='../main/Service.asp' target='main' >售后管理</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 46, 1) = 1 Then %>
          <li><a href='../main/Expense.asp' target='main' >费用管理</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 7, 1) = 1 Then %>
          <li><a href='../main/Export.asp' target='main' >数据导出</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 8, 1) = 1 Then %>
          <li><a href='../main/Import.asp' target='main' >数据导入</a></li>
		<%end if%>
          <li><a href='../main/Recycler.asp' target='main' >系统公海</a></li>

        </ul>
<%end if%>
<% end sub %>

<%sub Company() %>
<%if Session("CRM_account")<>"" then%>
        <h2 id="sidebar_Company">办公OA</h2>
        <ul>
		<% If mid(Session("CRM_qx"), 56, 1) = 1 Then %>
          <li><a href='../OA/Notice.asp' target='main' >内部公文</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 61, 1) = 1 Then %>
          <li><a href='../OA/Receive.asp' target='main' >站内短信</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 66, 1) = 1 Then %>
          <li><a href='../OA/Report.asp' target='main' >工作报告</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 71, 1) = 1 Then %>
          <li><a href='../Soft/index.asp' target='main' >文 件 柜</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 72, 1) = 1 Then %>
          <li><a href='../OA/Contact.asp' target='main' >通 讯 录</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 73, 1) = 1 Then %>
          <li><a href='../OA/Calendar.asp' target='main' >个人日历</a></li>
		<%end if%>
		<% If mid(Session("CRM_qx"), 74, 1) = 1 Then %>
          <li><a href='../Plugin/WebIm' target='main' >内部聊天</a></li>
		<%end if%>
        </ul>
<%end if%>
<% end sub %>

<%sub System() %>
<%If mid(Session("CRM_qx"), 5, 1) = 1 Then%>
        <h2 id="sidebar_setting">系统管理</h2>
        <ul>
          <li><a href='../system/Setting.asp' target='main' >全局设置</a></li>
          <li><a href='../system/QuickMenu.asp' target='main' >快捷菜单</a></li>
          <li><a href='../system/Select.asp' target='main' >下拉框</a></li>
          <li><a href='../system/Products.asp' target='main' >产品管理</a></li>
          <li><a href='../system/AreaData.asp' target='main' >地区管理</a></li>
          <li><a href='../system/CustomField.asp' target='main' >自定义字段</a></li>
          <li><a href='../system/Lang.asp' target='main' >语言包</a></li>
          <li><a href='../system/sql.asp' target='main' >数据库管理</a></li>
        </ul>
        <h2 id="sidebar_user">用户管理</h2>
        <ul>
          <li><a href='../system/User.asp' target='main' >员工管理</a></li>
          <li><a href='../system/Group.asp' target='main' >部门设置</a></li>
          <li><a href='../system/Level.asp' target='main' >角色管理</a></li>
        </ul>
		<% if YnUserLog=1 then %>
        <h2 id="sidebar_log">日志管理</h2>
        <ul>
          <li><a href='../system/Log_user.asp' target='main' >登录日志</a></li>
          <li><a href='../system/Logfile.asp' target='main' >操作记录</a></li>
        </ul>
		<% end if %>
<%end if%>
<% end sub %>

<%sub Plugin() %>
        <h2 id="sidebar_plugin"><%=lmgncj%></h2>
        <ul>
		<%
		Dim rsplugin
		Set rsplugin = Server.CreateObject("ADODB.Recordset")
		rsplugin.Open "Select * From [Plugin] where pYn=1 Order By pSort asc,Id asc",conn,1,1
		Do While Not rsplugin.BOF And Not rsplugin.EOF
		%>
          <li><a href='../plugin/<%=rsplugin("pUrl")%>/' target='main' ><%=rsplugin("pTitle")%></a></li>
		<%
		rsplugin.MoveNext
		Loop
		rsplugin.Close
		Set rsplugin = Nothing
		%>
		<%If Session("CRM_level") = 9 Then%>
          <li><a href='../plugin/index.asp' target='main' >插件管理</a></li>
		<%end if%>
        </ul>
<% end sub %>


        
        </div><!--/ .con-->
        </div><!--/ sidebar-->
    </div>
</div>

<script type="text/javascript">
$(document).ready(function(){
    var aArr = $(".con").find("li:first a");
  if (aArr && aArr.html() == "") or (aarr && aarr.html()=="")
    {
        aArr.addClass("active");
        $('#main', window.parent.document).attr('src', aArr.attr('href'));
    }
})
</script>
</body>
</html><% Set EasyCrm = nothing %>
