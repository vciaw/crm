<!--#include file="../../data/conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<title><%=title%></title>
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script src="<%=SiteUrl&skinurl%>Js/Common.js" type="text/javascript"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>
<body>
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：功能插件 > 外出申请 > 安装</td>
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
			<td class="td_r_l">员工外出申请单</td>
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
	rs.Open "Select * From Plugin Where pUrl = 'GoOut'",conn,1,1
	If rs.RecordCount = 1 Then
		Response.Write("<script>alert(""该插件已经安装，请勿重复安装！"");location.href='../index.asp' ;</script>")
	Else
		conn.execute "insert into [Plugin](pTitle,pUrl,pAuthor,pVersion,pContent,pTime,pYn) values('外出申请','GoOut','易用科技','V2.0','员工外出申请单','"&now()&"','1')"
		if accsql=1 then
		conn.execute "create table [Plugin_GoOut]([ID] [int] IDENTITY(1,1) NOT NULL primary key,[cid] [int] NULL,[gAddress] [nvarchar](255) NULL,[gSTime] [datetime] NULL,[gETime] [datetime] NULL,[gReason] [ntext] NULL,[gLinkman] [nvarchar](50) NULL,[gTel] [nvarchar](50) NULL,[gContent] [ntext] NULL,[gUser] [nvarchar](50) NULL,[gTime] [datetime] NULL,[gAudit] [nvarchar](50) NULL,[gAuditTime] [datetime] NULL,[gState] [int] NULL )"
		else
		conn.execute "create table [Plugin_GoOut](ID autoincrement primary key,cid int,gAddress varchar(255),gSTime datetime,gETime datetime,gReason memo,gLinkman varchar(50),gTel varchar(50),gContent memo,gUser varchar(50),gTime datetime,gAudit varchar(50),gAuditTime datetime,gState int)" 
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