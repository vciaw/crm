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
		<td class="top_left td_t_n td_r_n">当前位置：功能插件 > 财务管理 > 安装</td>
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
        <tr class="tr_t"> 
			<td class="td_l_l" COLSPAN="6"><B>当前位置：功能插件 > 安装</B></td>
        </tr>
        <tr> 
			<td class="td_l_c title">适用版本</td>
			<td class="td_r_l"> EasyCrm 2013 V4.0 + 商业版</td>
        </tr>
        <tr> 
			<td class="td_l_c title">插件简介</td>
			<td class="td_r_l">公司内部流水账</td>
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
	rs.Open "Select * From Plugin Where pUrl = 'Finance'",conn,1,1
	If rs.RecordCount = 1 Then
		Response.Write("<script>alert(""该插件已经安装，请勿重复安装！"");location.href='../index.asp' ;</script>")
	Else
		conn.execute "insert into [Plugin](pTitle,pUrl,pAuthor,pVersion,pContent,pTime,pYn) values('财务管理','Finance','易用科技','V1.0','公司内部流水账','"&now()&"','1')"
		if accsql=1 then
		conn.execute "create table [Plugin_Finance]([ID] [int] IDENTITY(1,1) NOT NULL primary key,[fSubjects] [nvarchar](50) NULL,[fGroup] [int] NULL,[fClass] [nvarchar](50) NULL,[fInvoice] [nvarchar](50) NULL,[fProject] [nvarchar](50) NULL,[fDigest] [nvarchar](255) NULL,[fDebit] [money] NULL,[fCredit] [money] NULL,[fRemark] [ntext] NULL,[fState] [nvarchar](50) NULL,[fUser] [nvarchar](50) NULL,[fAudit] [nvarchar](50) NULL,[fTime] [datetime] NULL )"
		
		conn.execute "create table [Plugin_Finance_Bank]([ID] [int] IDENTITY(1,1) NOT NULL primary key,[bName] [nvarchar](50) NULL,[bOpening] [nvarchar](50) NULL,[bCard] [nvarchar](50) NULL,[bClass] [nvarchar](50) NULL,[bDebit] [money] NULL,[bCredit] [money] NULL,[bDigest] [nvarchar](255) NULL,[bInvoice] [nvarchar](50) NULL,[bAudit] [nvarchar](50) NULL,[bRemark] [ntext] NULL,[bTime] [datetime] NULL )"
		
		conn.execute "create table [Plugin_Finance_Outin]([ID] [int] IDENTITY(1,1) NOT NULL primary key,[oInvoice] [nvarchar](50) NULL,[oCompany] [nvarchar](50) NULL,[oDebit] [money] NULL,[oCredit] [money] NULL,[oDigest] [nvarchar](255) NULL,[oRemark] [ntext] NULL,[oState] [nvarchar](50) NULL,[oAudit] [nvarchar](50) NULL,[oTime] [datetime] NULL )"
	
		else
		conn.execute "create table [Plugin_Finance](ID autoincrement primary key,fSubjects varchar(50),fGroup int,fClass varchar(50),fInvoice varchar(50),fProject varchar(50),fDigest varchar(50),fDebit Currency,fCredit Currency,fRemark memo,fState varchar(50),fUser varchar(50),fAudit varchar(50),fTime datetime)" 
		
		conn.execute "create table [Plugin_Finance_Bank](ID autoincrement primary key,bName varchar(50),bOpening varchar(50),bCard varchar(50),bClass varchar(50),bDebit Currency,bCredit Currency,bDigest varchar(255),bInvoice varchar(50),bAudit varchar(50),bRemark memo,bTime datetime)" 
		
		conn.execute "create table [Plugin_Finance_Outin](ID autoincrement primary key,oInvoice varchar(50),oCompany varchar(50),oDebit Currency,oCredit Currency,oDigest varchar(255),oRemark memo,oState varchar(50),oAudit varchar(50),oTime datetime)" 
		
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