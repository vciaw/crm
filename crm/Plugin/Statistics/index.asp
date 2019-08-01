<!--#include file="../../data/conn.asp" --><!--#include file="../../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
otype	=	Request.QueryString("otype")
action = Trim(Request("action"))
if otype="" then otype="Main"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<title><%=title%></title>
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript" src="js/highcharts.js"></script>
<script type="text/javascript" src="js/modules/exporting.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>

<body> 
<%if Action <>"InfoView" then%>
<style>body{padding-top:35px}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：功能插件 > 数据统计</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
            <div class="MenuboxS">
              <ul>
                <li <%if otype="Main" or otype="" then%>class="hover"<%end if%>><span><a href="?action=Main&otype=Main">客户概况</a></span></li>
                <li <%if otype="Records" then%>class="hover"<%end if%>><span><a href="?action=Records&otype=Records">跟单统计</a></span></li>
                <li <%if otype="Order" then%>class="hover"<%end if%>><span><a href="?action=Order&otype=Order">订单统计</a></span></li>
                <li <%if otype="Hetong" then%>class="hover"<%end if%>><span><a href="?action=Hetong&otype=Hetong">合同统计</a></span></li>
                <li <%if otype="Service" then%>class="hover"<%end if%>><span><a href="?action=Service&otype=Service">售后统计</a></span></li>
                <li <%if otype="Yeardata" then%>class="hover"<%end if%>><span><a href="?action=Yeardata&otype=Yeardata">年度曲线</a></span></li>
				<%if Session("CRM_level")>1 then%>
                <li <%if otype="Searchdata" then%>class="hover"<%end if%>><span><a href="?action=Searchdata&otype=Searchdata">详细数据</a></span></li>
				<%end if%>
              </ul>
            </div>
		</td>
	</tr>
</table>
<%end if%>
<%
Select Case action
Case "Records"
    Call Records()
Case "Order"
    Call Order()
Case "Hetong"
    Call Hetong()
Case "Service"
    Call Service()
Case "Yeardata"
    Call Yeardata()
Case "Searchdata"
    Call Searchdata()
Case "InfoView"
    Call InfoView()
Case Else
    Call Main()
	
End Select
%>

<%Sub Main()%>

<script type="text/javascript">
Highcharts.visualize=function(table,options){options.xAxis.categories=[];$('tbody th',table).each(function(i){options.xAxis.categories.push(this.innerHTML)});options.series=[];$('tr',table).each(function(i){var tr=this;$('th, td',tr).each(function(j){if(j>0){if(i==0){options.series[j-1]={name:this.innerHTML,data:[]}}else{options.series[j-1].data.push(parseFloat(this.innerHTML))}}})});var chart=new Highcharts.Chart(options)};$(document).ready(function(){var table=document.getElementById('datatable'),options={chart:{renderTo:'container',defaultSeriesType:'column'},title:{text:''},xAxis:{},yAxis:{title:{text:'HuohoCrm'}},tooltip:{formatter:function(){return'<b>'+this.series.name+'</b><br/>'+this.y+' '+this.x.toLowerCase()}}};Highcharts.visualize(table,options)});
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pd10"> 

					<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="50%" /><col width="50%" />
						<tr class="tr_t">
							<td class="td_l_l" colspan=2>系统概况</td>
						</tr>
						<tr>
							<td class="td_l_c" colspan=2>
							<div id="container" style="width: 100%; height: 400px;"></div>
							</td>
						</tr>
					</table>
        </td>
	</tr>
</table>

<table id="datatable" style="display:none;">
	<thead><tr><th></th><th>24小时内更新</th><th>7天内更新</th><th>30天内更新</th><th>合计</th></tr></thead>
	<tbody>
		<tr>
			<th>客户档案</th>
			<td><%=getCountnum("cId","clientnum","Client","cLastUpdated",""&date()-1&"","","","","cYn",1,"","","cUser")%></td>
			<td><%=getCountnum("cId","clientnum","Client","cLastUpdated",""&date()-7&"","","","","cYn",1,"","","cUser")%></td>
			<td><%=getCountnum("cId","clientnum","Client","cLastUpdated",""&date()-30&"","","","","cYn",1,"","","cUser")%></td>
			<td><%=getCountnum("cId","clientnum","Client","","","","","","cYn",1,"","","cUser")%></td>
		</tr>
		<tr>
			<th>联系人</th>
			<td><%=getCountnum("lId","Linkmansnum","Linkmans","lTime",""&date()-1&"","","","","","","","","lUser")%></td>
			<td><%=getCountnum("lId","Linkmansnum","Linkmans","lTime",""&date()-7&"","","","","","","","","lUser")%></td>
			<td><%=getCountnum("lId","Linkmansnum","Linkmans","lTime",""&date()-30&"","","","","","","","","lUser")%></td>
			<td><%=getCountnum("lId","Linkmansnum","Linkmans","","","","","","","","","","lUser")%></td>
		</tr>
		<tr>
			<th>跟单记录</th>
			<td><%=getCountnum("rId","Recordsnum","Records","rTime",""&date()-1&"","","","","","","","","rUser")%></td>
			<td><%=getCountnum("rId","Recordsnum","Records","rTime",""&date()-7&"","","","","","","","","rUser")%></td>
			<td><%=getCountnum("rId","Recordsnum","Records","rTime",""&date()-30&"","","","","","","","","rUser")%></td>
			<td><%=getCountnum("rId","Recordsnum","Records","","","","","","","","","","rUser")%></td>
		</tr>
		<tr>
			<th>订单记录</th>
			<td><%=getCountnum("oId","Ordernum","Order","oTime",""&date()-1&"","","","","","","","","oUser")%></td>
			<td><%=getCountnum("oId","Ordernum","Order","oTime",""&date()-7&"","","","","","","","","oUser")%></td>
			<td><%=getCountnum("oId","Ordernum","Order","oTime",""&date()-30&"","","","","","","","","oUser")%></td>
			<td><%=getCountnum("oId","Ordernum","Order","","","","","","","","","","oUser")%></td>
		</tr>
		<tr>
			<th>合同记录</th>
			<td><%=getCountnum("hId","Hetongnum","Hetong","hTime",""&date()-1&"","","","","","","","","hUser")%></td>
			<td><%=getCountnum("hId","Hetongnum","Hetong","hTime",""&date()-7&"","","","","","","","","hUser")%></td>
			<td><%=getCountnum("hId","Hetongnum","Hetong","hTime",""&date()-30&"","","","","","","","","hUser")%></td>
			<td><%=getCountnum("hId","Hetongnum","Hetong","","","","","","","","","","hUser")%></td>
		</tr>
		<tr>
			<th>售后记录</th>
			<td><%=getCountnum("sId","Servicenum","Service","sTime",""&date()-1&"","","","","","","","","sUser")%></td>
			<td><%=getCountnum("sId","Servicenum","Service","sTime",""&date()-7&"","","","","","","","","sUser")%></td>
			<td><%=getCountnum("sId","Servicenum","Service","sTime",""&date()-30&"","","","","","","","","sUser")%></td>
			<td><%=getCountnum("sId","Servicenum","Service","","","","","","","","","","sUser")%></td>
		</tr>
	</tbody>
</table>
<script type="text/javascript" src="js/themes/grid.js"></script>
<%
end sub

sub Records()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pd10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td class="td_l_l" colspan=2>
					
					<span class="right">
					<% Set rs = Server.CreateObject("ADODB.Recordset")
					if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
					end if
					rs.pagesize=10
					if request("page")<>"" then
					epage=cint(request("page"))
					if epage<1 then epage=1
					if epage>rs.pagecount then epage=rs.pagecount
					else
					epage=1
					end if
					rs.absolutepage=epage
					for j=1 to rs.pagecount
						if j=epage then
					%>
						<b style="color:#fff">第<%=j%>页</b>　
					<%
						else
					%>
						<a href='?action=Records&otype=Records&page=<%=j%>'><b>第<%=j%>页</b></a>　
					<%
						end if
					next
					rs.Close
					Set rs = Nothing
					%> 
					</span><B>按员工统计</B>
					</td>
				</tr>
				<tr >
					<td class="td_l_tj" colspan=2>
		<script type="text/javascript">
		$(function () { var chart; $(document).ready(function() {
        chart = new Highcharts.Chart({ chart: { renderTo: 'container11', type: 'line' },title: {text:''},
            xAxis: { categories: [ <% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%>'<%=rs("uName")%>'<%
				rs.MoveNext
				'Loop
				next
				rs.Close
				Set rs = Nothing
				%> ] },
            yAxis: { title: { text: 'HuohoCrm' } }, plotOptions: {line: { dataLabels: { enabled: true }, enableMouseTracking: false }},
            series: [{ name: '7天内记录', data: [<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("rId","Recordsnum","Records","rTime",""&date()-7&"","","","","","","rUser","'"&rs("uName")&"'","rUser")%><%
				rs.MoveNext
				'Loop
				next
				rs.Close
				Set rs = Nothing
				%> ] }, {
                name: '30天内记录', data: [	<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&")",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("rId","Recordsnum","Records","rTime",""&date()-30&"","","","","","","rUser","'"&rs("uName")&"'","rUser")%><%
				rs.MoveNext
				'Loop
				next
				rs.Close
				Set rs = Nothing
				%> ] }, {
                name: '所有记录', data: [<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user] ",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("rId","Recordsnum","Records","","","","","","","","rUser","'"&rs("uName")&"'","rUser")%><%
				rs.MoveNext
				'Loop
				next
				rs.Close
				Set rs = Nothing
				%> ] }]
        }); }); });
		</script>
					<div id="container11" style="width: 100%; height: 400px;"></div></td>
				</tr>
			</table>
        </td>
	</tr>
</table>
<%
end sub

sub Order()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pd10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td class="td_l_l" colspan=2>
					<span class="right">
					<% Set rs = Server.CreateObject("ADODB.Recordset")
					if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&")",conn,1,1
					end if
					rs.pagesize=10
					if request("page")<>"" then
					epage=cint(request("page"))
					if epage<1 then epage=1
					if epage>rs.pagecount then epage=rs.pagecount
					else
					epage=1
					end if
					rs.absolutepage=epage
					for j=1 to rs.pagecount
						if j=epage then
					%>
						<b style="color:#fff">第<%=j%>页</b>　
					<%
						else
					%>
						<a href='?action=Order&otype=Order&page=<%=j%>'><b>第<%=j%>页</b></a>　
					<%
						end if
					next
					rs.Close
					Set rs = Nothing
					%> 
					</span><B>按员工统计</B>
					</td>
				</tr>
				<tr >
					<td class="td_l_tj" colspan=2>
		<script type="text/javascript">
		$(function () { var chart; $(document).ready(function() {
        chart = new Highcharts.Chart({ chart: { renderTo: 'container12', type: 'line' },title: {text:''},
            xAxis: { categories: [ <% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&")",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%>'<%=rs("uName")%>'<%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] },
            yAxis: { title: { text: 'HuohoCrm' } }, plotOptions: {line: { dataLabels: { enabled: true }, enableMouseTracking: false }},
            series: [{ name: '7天内记录', data: [<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&")",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("oId","Ordernum","Order","oTime",""&date()-7&"","","","","","","oUser","'"&rs("uName")&"'","oUser")%><%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] }, {
                name: '30天内记录', data: [	<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&")",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("oId","Ordernum","Order","oTime",""&date()-30&"","","","","","","oUser","'"&rs("uName")&"'","oUser")%><%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] }, {
                name: '所有记录', data: [<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&")",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("oId","Ordernum","Order","","","","","","","","oUser","'"&rs("uName")&"'","oUser")%><%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] }]
        }); }); });
		</script>
					<div id="container12" style="width: 100%; height: 400px;"></div></td>
				</tr>
			</table>
        </td>
	</tr>
</table>
<%
end sub

sub Hetong()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pd10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td class="td_l_l" colspan=2>
					<span class="right">
					<% Set rs = Server.CreateObject("ADODB.Recordset")
					if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
					end if
					rs.pagesize=10
					if request("page")<>"" then
					epage=cint(request("page"))
					if epage<1 then epage=1
					if epage>rs.pagecount then epage=rs.pagecount
					else
					epage=1
					end if
					rs.absolutepage=epage
					for j=1 to rs.pagecount
						if j=epage then
					%>
						<b style="color:#ff0">第<%=j%>页</b>　
					<%
						else
					%>
						<a href='?action=Hetong&otype=Hetong&page=<%=j%>' style="color:#fff"><b>第<%=j%>页</b></a>　
					<%
						end if
					next
					rs.Close
					Set rs = Nothing
					%></span> <B>按员工统计</B>
					</td>
				</tr>
				<tr >
					<td class="td_l_tj" colspan=2>
		<script type="text/javascript">
		$(function () { var chart; $(document).ready(function() {
        chart = new Highcharts.Chart({ chart: { renderTo: 'container13', type: 'line' },title: {text:''},
            xAxis: { categories: [ <% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%>'<%=rs("uName")%>'<%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] },
            yAxis: { title: { text: 'HuohoCrm' } }, plotOptions: {line: { dataLabels: { enabled: true }, enableMouseTracking: false }},
            series: [{ name: '7天内记录', data: [<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("hId","Hetongsnum","Hetong","hTime",""&date()-7&"","","","","","","hUser","'"&rs("uName")&"'","hUser")%><%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] }, {
                name: '30天内记录', data: [	<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("hId","Hetongsnum","Hetong","hTime",""&date()-30&"","","","","","","hUser","'"&rs("uName")&"'","hUser")%><%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] }, {
                name: '所有记录', data: [<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("hId","Hetongsnum","Hetong","","","","","","","","hUser","'"&rs("uName")&"'","hUser")%><%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] }]
        }); }); });
		</script>
					<div id="container13" style="width: 100%; height: 400px;"></div></td>
				</tr>
			</table>
        </td>
	</tr>
</table>
<%
end sub

sub Service()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pd10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td class="td_l_l" colspan=2>
					<span class="right">
					<% Set rs = Server.CreateObject("ADODB.Recordset")
					if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
					end if
					rs.pagesize=10
					if request("page")<>"" then
					epage=cint(request("page"))
					if epage<1 then epage=1
					if epage>rs.pagecount then epage=rs.pagecount
					else
					epage=1
					end if
					rs.absolutepage=epage
					for j=1 to rs.pagecount
						if j=epage then
					%>
						<b style="color:#ff0">第<%=j%>页</b>　
					<%
						else
					%>
						<a href='?action=Service&otype=Service&page=<%=j%>' style="color:#fff"><b>第<%=j%>页</b></a>　
					<%
						end if
					next
					rs.Close
					Set rs = Nothing
					%></span> <B>按员工统计</B>
					</td>
				</tr>
				<tr >
					<td class="td_l_tj" colspan=2>
		<script type="text/javascript">
		$(function () { var chart; $(document).ready(function() {
        chart = new Highcharts.Chart({ chart: { renderTo: 'container14', type: 'line' },title: {text:''},
            xAxis: { categories: [ <% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%>'<%=rs("uName")%>'<%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] },
            yAxis: { title: { text: 'HuohoCrm' } }, plotOptions: {line: { dataLabels: { enabled: true }, enableMouseTracking: false }},
            series: [{ name: '7天内记录', data: [<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("sId","Servicenum","Service","sTime",""&date()-7&"","","","","","","sUser","'"&rs("uName")&"'","sUser")%><%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] }, {
                name: '30天内记录', data: [	<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("sId","Servicenum","Service","sTime",""&date()-30&"","","","","","","sUser","'"&rs("uName")&"'","sUser")%><%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] }, {
                name: '所有记录', data: [<% Set rs = Server.CreateObject("ADODB.Recordset")
				if Session("CRM_level")=9 then
					rs.Open "Select * From [user]",conn,1,1
					else
					rs.Open "Select * From [user] where uName in ("&arrUser&") ",conn,1,1
				end if
				
				rs.pagesize=10
				if request("page")<>"" then
				epage=cint(request("page"))
				if epage<1 then epage=1
				if epage>rs.pagecount then epage=rs.pagecount
				else
				epage=1
				end if
				rs.absolutepage=epage
				
				i=0
				'Do While Not rs.BOF And Not rs.EOF
				i=i+1
				for i=0 to rs.pagesize-1
				if rs.bof or rs.eof then exit for
				%><%if i>0 then%>,<%end if%><%=getCountnum("sId","Servicenum","Service","","","","","","","","sUser","'"&rs("uName")&"'","sUser")%><%
				rs.MoveNext
				next
				rs.Close
				Set rs = Nothing
				%> ] }]
        }); }); });
		</script>
					<div id="container14" style="width: 100%; height: 400px;"></div></td>
				</tr>
			</table>
        </td>
	</tr>
</table>
<%
end sub

Function Clientnum(m)
if Session("CRM_level")<9 then sqlc=" and cUser In ("&arrUser&")"
if Accsql =1 then
Set Rsstr=conn.Execute("Select count(cid) As Clientnum From client where cYn = 1 and Year(cdate)=Year(getdate()) and Month(cdate)="&m&" "&sqlc&" ",1,1)
else
Set Rsstr=conn.Execute("Select count(cid) As Clientnum From client where cYn = 1 and Year(cdate)=Year(now()) and Month(cdate)="&m&" "&sqlc&" ",1,1)
end if
Clientnum=Rsstr("Clientnum") 
Rsstr.Close 
Set Rsstr=Nothing
end Function

Function Recordsnum(m)
if Session("CRM_level")<9 then sqlr=" and rUser In ("&arrUser&")"
if Accsql =1 then
Set Rsstr=conn.Execute("Select count(cid) As Recordsnum From Records where Year(rTime)=Year(getdate()) and Month(rTime)="&m&" "&sqlr&" ",1,1)
else
Set Rsstr=conn.Execute("Select count(cid) As Recordsnum From Records where Year(rTime)=Year(now()) and Month(rTime)="&m&" "&sqlr&" ",1,1)
end if
Recordsnum=Rsstr("Recordsnum") 
Rsstr.Close 
Set Rsstr=Nothing
end Function

Function RecordsPlannum(m)
if Session("CRM_level")<9 then sqlr=" and rUser In ("&arrUser&")"
if Accsql =1 then
Set Rsstr=conn.Execute("Select count(cid) As RecordsPlannum From RecordsPlan where Year(rTime)=Year(getdate()) and Month(rTime)="&m&" "&sqlr&" ",1,1)
else
Set Rsstr=conn.Execute("Select count(cid) As RecordsPlannum From RecordsPlan where Year(rTime)=Year(now()) and Month(rTime)="&m&" "&sqlr&" ",1,1)
end if
RecordsPlannum=Rsstr("RecordsPlannum") 
Rsstr.Close 
Set Rsstr=Nothing
end Function

Function Hetongnum(m)
if Session("CRM_level")<9 then sqlh=" and hUser In ("&arrUser&")"
if Accsql =1 then
Set Rsstr=conn.Execute("Select count(cid) As Hetongnum From Hetong where Year(hTime)=Year(getdate()) and Month(hTime)="&m&" "&sqlh&" ",1,1)
else
Set Rsstr=conn.Execute("Select count(cid) As Hetongnum From Hetong where Year(hTime)=Year(now()) and Month(hTime)="&m&" "&sqlh&" ",1,1)
end if
Hetongnum=Rsstr("Hetongnum") 
Rsstr.Close 
Set Rsstr=Nothing
end Function

Sub Yeardata()
'Set Rsstr=conn.Execute("Select count(cid) As Client01 From client where cYn = 1 and Year(cdate)=Year(getdate()) and Month(cdate)=1 "&sql&" ",1,1)

%>
<script type="text/javascript">
theme = 'grid';
var chart;$(document).ready(function(){chart=new Highcharts.Chart({chart:{renderTo:'container',defaultSeriesType:'line'},title: {text:''},xAxis:{categories:['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']},yAxis:{title:{text:'HuohoCrm'}},tooltip:{enabled:false,formatter:function(){return'<b>'+this.series.name+'</b><br/>'+this.x+': '+this.y+'°C'}},plotOptions:{line:{dataLabels:{enabled:true},enableMouseTracking:false}},series:[
{name:'新增客户',data:[<%
for i=1 to 12 Step 1
if i = 12 then
Response.Write(getCountnum("cid","Clientnum","client","","","","Year(cdate)",""&Year(date())&"","Month(cdate)","'"&i&"'","cYn",1,"cUser"))
else
Response.Write(getCountnum("cid","Clientnum","client","","","","Year(cdate)",""&Year(date())&"","Month(cdate)","'"&i&"'","cYn",1,"cUser")&",")
end if
Next
%>]},
{name:'跟单次数',data:[<%
for i=1 to 12 Step 1
if i = 12 then
Response.Write(getCountnum("cid","Recordsnum","Records","","","","Year(rTime)",""&Year(date())&"","Month(rTime)","'"&i&"'","","","rUser"))
else
Response.Write(getCountnum("cid","Recordsnum","Records","","","","Year(rTime)",""&Year(date())&"","Month(rTime)","'"&i&"'","","","rUser")&",")
end if
Next
%>]},
{name:'订单数量',data:[<%
for i=1 to 12 Step 1
if i = 12 then
Response.Write(getCountnum("cid","Ordernum","Order","","","","Year(oTime)",""&Year(date())&"","Month(oTime)","'"&i&"'","","","oUser"))
else
Response.Write(getCountnum("cid","Ordernum","Order","","","","Year(oTime)",""&Year(date())&"","Month(oTime)","'"&i&"'","","","oUser")&",")
end if
Next
%>]},
{name:'合同数量',data:[<%
for i=1 to 12 Step 1
if i = 12 then
Response.Write(getCountnum("cid","Hetongnum","Hetong","","","","Year(hTime)",""&Year(date())&"","Month(hTime)","'"&i&"'","","","hUser"))
else
Response.Write(getCountnum("cid","Hetongnum","Hetong","","","","Year(hTime)",""&Year(date())&"","Month(hTime)","'"&i&"'","","","hUser")&",")
end if
Next
%>]},
{name:'售后次数',data:[<%
for i=1 to 12 Step 1
if i = 12 then
Response.Write(getCountnum("cid","Servicenum","Service","","","","Year(sTime)",""&Year(date())&"","Month(sTime)","'"&i&"'","","","sUser"))
else
Response.Write(getCountnum("cid","Servicenum","Service","","","","Year(sTime)",""&Year(date())&"","Month(sTime)","'"&i&"'","","","sUser")&",")
end if
Next
%>]}
]})});
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pd10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td class="td_l_l"><B><%=Year(now())%> 年度客户统计报表</B>
					</td>
				</tr>
				<tr >
					<td class="td_l_tj"><div id="container" style="width: 100%; height: 400px;"></div></td>
				</tr>
			</table>
        </td>
	</tr>
</table>
<%
end sub

Sub InfoView()	
	uID = request("uID")
	sType = request("sType")
	uName = EasyCrm.getNewItem("User","uID",""&uID&"","uName")
	if Accsql = 1 then 
		Nowdate = "Getdate"
	else
		Nowdate = "Now"
	end if
%>
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="100" />
							<tr class="tr_t"><td class="td_l_l" colspan=2>客户档案</td></tr>
							<tr>
								<td class="td_l_r title">按更新状态</td>
								<td class="td_l_l">
									总记录：<font color=red><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&uName&"' ")%></font> 条　
									7天内：<font color=red><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&uName&"' and cDate + 7 >= "&Nowdate&"() ")%></font> 条　
									7-30天：<font color=red><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&uName&"' and cDate + 30 >= "&Nowdate&"() and cDate + 7 <= "&Nowdate&"() ")%></font> 条　
									30天以上：<font color=red><%=EasyCrm.getCountItem("Client","cid","cid"," and cYn=1 and cUser = '"&uName&"' and cDate + 30 <= "&Nowdate&"() ")%></font> 条　
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">按客户类型</td>
								<td class="td_l_l">
									<% Set rsp = Server.CreateObject("ADODB.Recordset")
									rsp.Open "Select * From [SelectData] where Select_Type<>'' and Select_Type<>'Null' ",conn,1,1
									Do While Not rsp.BOF And Not rsp.EOF
									%><%=rsp("Select_Type")%>：<font color=red><%=getCountnum("cId","clientnum","Client","","","","cType","'"&rsp("Select_Type")&"'","cYn",1,"cUser","'"&uName&"'","cUser")%></font> 条　
									<% 
									rsp.MoveNext
									Loop
									%>
									未知：<font color=red><%=getCountnum("cId","clientnum","Client","","","","cType","''","cYn",1,"cUser","'"&uName&"'","cUser")%></font> 条
									<%
									rsp.Close
									Set rsp = Nothing %>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">按客户级别</td>
								<td class="td_l_l">
									<% Set rss = Server.CreateObject("ADODB.Recordset")
									rss.Open "Select * From [SelectData] where Select_Star<>'' and Select_Star<>'Null' ",conn,1,1
									Do While Not rss.BOF And Not rss.EOF
									%><%=rss("Select_Star")%>：<font color=red><%=getCountnum("cId","clientnum","Client","","","","cStart","'"&rss("Select_Star")&"'","cYn",1,"cUser","'"&uName&"'","cUser")%></font> 条　
									<% 
									rss.MoveNext
									Loop
									%>
									未知：<font color=red><%=getCountnum("cId","clientnum","Client","","","","cStart","''","cYn",1,"cUser","'"&uName&"'","cUser")%></font> 条
									<%
									rss.Close
									Set rss = Nothing %>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">按客户来源</td>
								<td class="td_l_l">
									<% Set rss = Server.CreateObject("ADODB.Recordset")
									rss.Open "Select * From [SelectData] where Select_Source<>'' and Select_Source<>'Null' ",conn,1,1
									Do While Not rss.BOF And Not rss.EOF
									%><%=rss("Select_Source")%>：<font color=red><%=getCountnum("cId","clientnum","Client","","","","cSource","'"&rss("Select_Source")&"'","cYn",1,"cUser","'"&uName&"'","cUser")%></font> 条　
									<% 
									rss.MoveNext
									Loop
									%>
									未知：<font color=red><%=getCountnum("cId","clientnum","Client","","","","cSource","''","cYn",1,"cUser","'"&uName&"'","cUser")%></font> 条
									<%
									rss.Close
									Set rss = Nothing %>
								</td>
							</tr>
						</table>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin-top:10px;">
						<col width="100" />
							<tr class="tr_t"><td class="td_l_l" colspan=2>跟单记录</td></tr>
							<tr>
								<td class="td_l_r title">按更新状态</td>
								<td class="td_l_l">
									总记录：<font color=red><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&uName&"' ")%></font> 条　
									7天内：<font color=red><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&uName&"' and rTime + 7 >= "&Nowdate&"() ")%></font> 条　
									7-30天：<font color=red><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&uName&"' and rTime + 30 >= "&Nowdate&"() and rTime + 7 <= "&Nowdate&"() ")%></font> 条　
									30天以上：<font color=red><%=EasyCrm.getCountItem("Records","rid","rid"," and rUser = '"&uName&"' and rTime + 30 <= "&Nowdate&"() ")%></font> 条　
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">按跟单类型</td>
								<td class="td_l_l">
									<% Set rsp = Server.CreateObject("ADODB.Recordset")
									rsp.Open "Select * From [SelectData] where Select_Records<>'' and Select_Records<>'Null' ",conn,1,1
									Do While Not rsp.BOF And Not rsp.EOF
									%><%=rsp("Select_Records")%>：<font color=red><%=getCountnum("rId","Recordsnum","Records","","","","rType","'"&rsp("Select_Records")&"'","","","rUser","'"&uName&"'","rUser")%></font> 条　
									<% 
									rsp.MoveNext
									Loop
									%>
									未知：<font color=red><%=getCountnum("rId","Recordsnum","Records","","","","rType","''","","","rUser","'"&uName&"'","rUser")%></font> 条
									<%
									rsp.Close
									Set rsp = Nothing %>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">按跟单进度</td>
								<td class="td_l_l">
									<% Set rsp = Server.CreateObject("ADODB.Recordset")
									rsp.Open "Select * From [SelectData] where Select_Type<>'' and Select_Type<>'Null' ",conn,1,1
									Do While Not rsp.BOF And Not rsp.EOF
									%><%=rsp("Select_Type")%>：<font color=red><%=getCountnum("rId","Recordsnum","Records","","","","rState","'"&rsp("Select_Type")&"'","","","rUser","'"&uName&"'","rUser")%></font> 条　
									<% 
									rsp.MoveNext
									Loop
									%>
									未知：<font color=red><%=getCountnum("rId","Recordsnum","Records","","","","rState","''","","","rUser","'"&uName&"'","rUser")%></font> 条
									<%
									rsp.Close
									Set rsp = Nothing %>
								</td>
							</tr>
						</table>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin-top:10px;">
						<col width="100" />
							<tr class="tr_t"><td class="td_l_l" colspan=2>订单记录</td></tr>
							<tr>
								<td class="td_l_r title">按更新状态</td>
								<td class="td_l_l">
									总记录：<font color=red><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&uName&"' ")%></font> 条　
									7天内：<font color=red><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&uName&"' and oTime + 7 >= "&Nowdate&"() ")%></font> 条　
									7-30天：<font color=red><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&uName&"' and oTime + 30 >= "&Nowdate&"() and oTime + 7 <= "&Nowdate&"() ")%></font> 条　
									30天以上：<font color=red><%=EasyCrm.getCountItem("Order","oid","oid"," and oUser = '"&uName&"' and oTime + 30 <= "&Nowdate&"() ")%></font> 条　
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">按订单状态</td>
								<td class="td_l_l">
									未处理：<font color=red><%=getCountnum("oId","Ordernum","Order","","","","oState",0,"","","oUser","'"&uName&"'","oUser")%></font> 条　
									处理中：<font color=red><%=getCountnum("oId","Ordernum","Order","","","","oState",1,"","","oUser","'"&uName&"'","oUser")%></font> 条　
									已完成：<font color=red><%=getCountnum("oId","Ordernum","Order","","","","oState",2,"","","oUser","'"&uName&"'","oUser")%></font> 条　
								</td>
							</tr>
						</table>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin-top:10px;">
						<col width="100" />
							<tr class="tr_t"><td class="td_l_l" colspan=2>合同记录</td></tr>
							<tr>
								<td class="td_l_r title">按更新状态</td>
								<td class="td_l_l">
									总记录：<font color=red><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&uName&"' ")%></font> 条　
									7天内：<font color=red><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&uName&"' and hTime + 7 >= "&Nowdate&"() ")%></font> 条　
									7-30天：<font color=red><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&uName&"' and hTime + 30 >= "&Nowdate&"() and hTime + 7 <= "&Nowdate&"() ")%></font> 条　
									30天以上：<font color=red><%=EasyCrm.getCountItem("Hetong","hid","hid"," and hUser = '"&uName&"' and hTime + 30 <= "&Nowdate&"() ")%></font> 条　
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">按合同分类</td>
								<td class="td_l_l">
									<% Set rsp = Server.CreateObject("ADODB.Recordset")
									rsp.Open "Select * From [SelectData] where Select_Hetong<>'' and Select_Hetong<>'Null' ",conn,1,1
									Do While Not rsp.BOF And Not rsp.EOF
									%><%=rsp("Select_Hetong")%>：<font color=red><%=getCountnum("hId","Hetongnum","Hetong","","","","hType","'"&rsp("Select_Hetong")&"'","","","hUser","'"&uName&"'","hUser")%></font> 条　
									<% 
									rsp.MoveNext
									Loop
									%>
									未知：<font color=red><%=getCountnum("hId","Hetongnum","Hetong","","","","hType","''","","","hUser","'"&uName&"'","hUser")%></font> 条
									<%
									rsp.Close
									Set rsp = Nothing %>
								
								</td>
							</tr>
							
							<% Contrs = conn.execute ("select sum(hMoney) as AllMoney,sum(hRevenue) as AllRevenue,sum(hOwed) as AllOwed from [Hetong] where hUser='"&uName&"' ") %><%if Contrs("AllMoney")<>"" then%>
							<tr>
								<td class="td_l_r title">合同金额</td>
								<td class="td_l_l">
								总金额：<font color=red><%=Contrs("AllMoney")%></font> 元　
								已收：<font color=red><%=Contrs("AllRevenue")%></font> 元　
								欠款：<font color=red><%=Contrs("AllOwed")%></font> 元
								</td>
							</tr>
							<%end if%>
						</table>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin-top:10px;">
						<col width="100" />
							<tr class="tr_t"><td class="td_l_l" colspan=2>售后记录</td></tr>
							<tr>
								<td class="td_l_r title">按更新状态</td>
								<td class="td_l_l">
									总记录：<font color=red><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&uName&"' ")%></font> 条　
									7天内：<font color=red><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&uName&"' and sTime + 7 >= "&Nowdate&"() ")%></font> 条　
									7-30天：<font color=red><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&uName&"' and sTime + 30 >= "&Nowdate&"() and sTime + 7 <= "&Nowdate&"() ")%></font> 条　
									30天以上：<font color=red><%=EasyCrm.getCountItem("Service","sid","sid"," and sUser = '"&uName&"' and sTime + 30 <= "&Nowdate&"() ")%></font> 条　
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">按反馈分类</td>
								<td class="td_l_l">
									<% Set rsp = Server.CreateObject("ADODB.Recordset")
									rsp.Open "Select * From [SelectData] where Select_Service<>'' and Select_Service<>'Null' ",conn,1,1
									Do While Not rsp.BOF And Not rsp.EOF
									%><%=rsp("Select_Service")%>：<font color=red><%=getCountnum("sId","Servicenum","Service","","","","sType","'"&rsp("Select_Service")&"'","","","sUser","'"&uName&"'","sUser")%></font> 条　　
									<% 
									rsp.MoveNext
									Loop
									%>
									未知：<font color=red><%=getCountnum("sId","Servicenum","Service","","","","sType","''","","","sUser","'"&uName&"'","hUser")%></font> 条
									<%
									rsp.Close
									Set rsp = Nothing %>
								
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
<%
end sub

Sub Searchdata()					
	if Request("Subaction")="Searchdata" then
		if Trim(Request("TimeBegin")) <> "" then Session("Search_ST_TimeBegin") = Trim(Request("TimeBegin"))
		if Trim(Request("TimeEnd")) <> "" then Session("Search_ST_TimeEnd") = Trim(Request("TimeEnd"))
	elseif  Request("Subaction")="killSession" then
		Session("Search_ST_TimeBegin") = ""
		Session("Search_ST_TimeEnd") = ""
	End If
%>
<style>body{padding-bottom:48px}</style>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan=2 class="Search_All td_n">
				<form name="searchForm" method="post" action="?Action=Searchdata&otype=Searchdata&Subaction=Searchdata">
				<input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" value="<%=Session("Search_ST_TimeBegin")%>" style="width:100px;" onFocus="WdatePicker()" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" value="<%=Session("Search_ST_TimeEnd")%>" style="width:100px;" onFocus="WdatePicker()" />&nbsp;<input type="submit" name="Submit" class="button222" value=" <%=L_Search%> "> <input type="button" name="button" class="button223" value=" <%=L_Clear%> " onClick=window.location.href="?Action=Searchdata&otype=Searchdata&Subaction=killSession" />
				</form>
		</td>
	</tr>
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdb10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" /><col width="80" /><col width="80" /><col width="80" /><col width="80" /><col width="80" />
				<col width="80" /><col width="80" /><col width="80" />
				<tr class="tr_t">
					<td class="td_l_c" colspan=6>基本情况</td>
					<td class="td_l_c" colspan=3>活跃状态</td>
				</tr>
				<tr class="tr_f">
					<td class="td_l_l">姓名</td>
					<td class="td_l_c">客户</td>
					<td class="td_l_c">跟单</td>
					<td class="td_l_c">订单</td>
					<td class="td_l_c">合同</td>
					<td class="td_l_c">售后</td>
					<td class="td_l_c">新增</td>
					<td class="td_l_c">修改</td>
					<td class="td_l_c">删除</td>
				</tr>
				<%
				Dim rs
				Dim intTotalRecords,intTotalPages,PN,intPageSize'记录总数，总页数，当前页，分页数量
				PN = CLng(ABS(Request("PN")))

				If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
				intPageSize = DataPageSize
				pagenum = intPageSize*(PN-1)
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				
				if Session("CRM_level")=9 then
					IF PN=1 THEN
					rs.Open "Select top "&intPageSize&" * From [user] Order By uId asc ",conn,1,1 
					ELSE
					rs.Open "Select top "&intPageSize&" * From [user] where uId > ( SELECT Max(uId) FROM ( SELECT TOP "&pagenum&" uId FROM [user] ORDER BY uId asc ) AS T ) Order By uId asc ",conn,1,1
					END IF
					SQLstr="Select count(uId) As RecordSum From [user] " '统计页码
				else
					IF PN=1 THEN
					rs.Open "Select top "&intPageSize&" * From [user] where uName in ("&arrUser&") Order By uId asc ",conn,1,1 
					ELSE
					rs.Open "Select top "&intPageSize&" * From [user] where uName in ("&arrUser&") and uId > ( SELECT Max(uId) FROM ( SELECT TOP "&pagenum&" uId FROM [user] where uName in ("&arrUser&") ORDER BY uId asc ) AS T ) Order By uId asc ",conn,1,1
					END IF
					SQLstr="Select count(uId) As RecordSum From [user] where uName in ("&arrUser&") " '统计页码
				end if
				
							
					Dim TotalRecords,TotalPages
					Set Rsstr=conn.Execute(SQLstr,1,1) 
					TotalRecords=Rsstr("RecordSum") 
					if Int(TotalRecords/DataPageSize)=TotalRecords/DataPageSize then
					TotalPages=TotalRecords/DataPageSize
					else
					TotalPages=Int(TotalRecords/DataPageSize)+1
					end if
					Rsstr.Close 
					Set Rsstr=Nothing

					If PN > TotalPages Then PN = TotalPages
											
				Do While Not rs.BOF And Not rs.EOF
				%>
				<Tr class="tr">
					<TD class="td_l_l"><a onclick='InfoView<%=rs("uId")%>()' style="cursor:pointer"> <%=rs("uName")%></a></TD>
					<td class="td_l_c"><%=getCountnum("cId","clientnum","Client","cDate",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","cYn",1,"cUser","'"&rs("uName")&"'","cUser")%></td>
					<td class="td_l_c"><%=getCountnum("cId","Recordsnum","Records","rTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","","","rUser","'"&rs("uName")&"'","rUser")%></td>
					<td class="td_l_c"><%=getCountnum("cId","Ordernum","Order","oTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","","","oUser","'"&rs("uName")&"'","oUser")%></td>
					<td class="td_l_c"><%=getCountnum("cId","Hetongnum","Hetong","hTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","","","hUser","'"&rs("uName")&"'","hUser")%></td>
					<td class="td_l_c"><%=getCountnum("cId","Servicenum","Service","sTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","","","sUser","'"&rs("uName")&"'","sUser")%></td>
					<td class="td_l_c"><%=getCountnum("lId","addnum","Logfile","lTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","lAction","'新增'","lUser","'"&rs("uName")&"'","lUser")%></td>
					<td class="td_l_c"><%=getCountnum("lId","addnum","Logfile","lTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","lAction","'修改'","lUser","'"&rs("uName")&"'","lUser")%></td>
					<td class="td_l_c"><%=getCountnum("lId","addnum","Logfile","lTime",""&Session("Search_ST_TimeBegin")&"",""&Session("Search_ST_TimeEnd")&"","","''","lAction","'删除'","lUser","'"&rs("uName")&"'","lUser")%></td>
				</TR>
				<script>function InfoView<%=rs("uId")%>() {$.dialog.open('index.asp?action=InfoView&uId=<%=rs("uId")%>', {title: '<%=rs("uName")%>的详细统计', width: 800,height: 480, fixed: true}); };</script>
				<tr style="display:none;" id="box<%=rs("uId")%>">
					<td class="td_l_l" colspan="9" style="padding:10px;background-color:#ffffff;Word-break: break-all; word-wrap:break-word;">
					</td>
				</tr>
						<%
							rs.MoveNext
						Loop
						rs.Close
						Set rs = Nothing
						%>
			</table>
        </td>
	</tr>
</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd ">
			<%=EasyCrm.pagelist("?Action=Searchdata&otype=Searchdata&Subaction=Searchdata", PN,TotalPages,TotalRecords)%>
		</td> 
	</tr>
</table>
</div>
<%end sub%>
<div class="PX10"></div>
</body>
</html>
<script src="../../data/calendar/WdatePicker.js"></script>
<%
'统计类
Function getCountnum(Item0,Item1,Item2,Item3,Item4,Item5,Item6,Item7,Item8,Item9,Item10,Item11,Item12)
    Dim rs,itemValue,sql
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = ""
	if Accsql =1 then
	if Item4<>"" then
	sql = sql & " and "&Item3&" >= '"&Item4&"'"
	end if
	if Item5<>"" then
	sql = sql & " and "&Item3&" <= '"&Item5&"'"
	end if
	else
	if Item4<>"" then
	sql = sql & " and "&Item3&" >= #"&Item4&"#"
	end if
	if Item5<>"" then
	sql = sql & " and "&Item3&" <= #"&Item5&"#"
	end if
	end if
	if Item6<>"" and Item7<>"" then
	sql = sql & " and "&Item6&" = " & Item7&""
	end if
	if Item8<>"" and Item9<>"" then
	sql = sql & " and "&Item8&" = " & Item9&""
	end if
	if Item10<>"" and Item11<>"" then
	sql = sql & " and "&Item10&" = " & Item11&""
	end if
	
	If Session("CRM_level") < 9 Then
		sql = sql & " And "&Item12&" In (" & arrUser & ")"
	else
		sql = sql & ""
	end if
	
	rs.Open "Select count("&Item0&") As "&Item1&" From ["&Item2&"] Where 1=1 "&sql&" " ,conn,1,1
	    itemValue = rs(Item1)
	rs.Close
	Set rs = Nothing
	getCountnum = itemValue
End Function
%><% Set EasyCrm = nothing %>