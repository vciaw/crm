<!--#include file="../../data/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>
<body>
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：功能插件 > 任务发布 > 安装</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<%
a = Trim(Request("a"))
Select Case a
Case "setup01"
    Call setup01()
Case "setup02"
    Call setup02()
Case Else
    Call setup01()
End Select
%>
<%
Sub setup01()
%>      
<form name="login" action="?a=setup02" method="post" style="padding:10px;">
	<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="100" />
        <tr> 
			<td class="td_l_c title">适用版本</td>
			<td class="td_r_l"> EasyCrm 2013 V4.0 + 商业版</td>
        </tr>
        <tr> 
			<td class="td_l_c title">插件简介</td>
			<td class="td_r_l">任务发布管理，指定主次负责人，反馈工作进度，领导批注；</td>
        </tr>
        <tr> 
			<td class="td_r_l" colspan=2> <input type="Submit" class="button45" value="安装" /></td>
        </tr>
    </table>
</form>
<%
end Sub
%>

<%
Sub setup02()
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select pTitle From Plugin where pTitle = '任务发布' ",conn,1,1
	If rs.RecordCount = 1 Then
		Response.Write("<script>alert(""该插件已经安装，请勿重复安装！"");location.href='../../index.asp' ;</script>")
	Else
		conn.execute "insert into [Plugin](pTitle,pUrl,pAuthor,pVersion,pContent,pTime,pYn) values('任务发布','Workschedule','易用科技','V2.0','任务发布管理，指定主次负责人，反馈工作进度，领导批注','2012/4/1 10:00:00','1')"
		if accsql=1 then
		conn.execute "create table [Plugin_Workschedule]([wId] [int] IDENTITY(1,1) NOT NULL primary key,[wClass] [nvarchar](50) NULL,[wStar] [nvarchar](50) NULL,[wTitle] [nvarchar](50) NULL,[wCompletiontime] [datetime] NULL,[wContent] [ntext] NULL,[wUserb] [nvarchar](50) NULL,[wUsers] [ntext] NULL,[wState] [nvarchar](50) NULL,[wMsg] [int] NULL,[wSH] [int] NULL,[wYd] [int] NULL,[wUser] [nvarchar](50) NULL,[wTime] [datetime] NULL )"
		conn.execute "create table [Plugin_Workschedule_re]([rID] [int] IDENTITY(1,1) NOT NULL primary key,[wID] [int] NULL,[rState] [nvarchar](50) NULL,[rContent] [ntext] NULL,[rRE] [ntext] NULL,[rUser] [nvarchar](50) NULL,[rTime] [datetime] NULL)" 
		else
		conn.execute "create table [Plugin_Workschedule](wId autoincrement primary key,wClass varchar(50),wStar varchar(50),wTitle varchar(50),wCompletiontime datetime,wContent memo,wUserb varchar(50),wUsers memo,wState varchar(50),wMsg int,wSH int,wYd int,wUser varchar(50),wTime datetime)" 
		conn.execute "create table [Plugin_Workschedule_re](rID autoincrement primary key,wID int,rState varchar(50) ,rContent memo,rRE memo,rUser varchar(50),rTime datetime)" 
		end if
	End If
    rs.Close
	Set rs = Nothing
	Response.Write("<script>alert(""安装成功！"");</script>")
    Response.Write("<script>location.href='../index.asp' ;</script>")
end Sub

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>