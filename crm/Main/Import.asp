<!--#include file="../data/conn.asp"--><!--#include file="../UpLoad/UpLoad.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 8, 1) = 1 Then %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<link href="<%=SiteUrl&skinurl%>chosen/chosen.css" rel="stylesheet" />
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Float.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>

<body style="padding-top:35px;"> 
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Company%> > 导入数据</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>
<%
Function FileList(FolderUrl,FileExName,FileUrl)
Set fso=Server.CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set folder=fso.GetFolder(Server.MapPath(Trim(FolderUrl)))
Set file=folder.Files
FileList=""
For Each FileName in file
IF Trim(FileUrl)<>"" Then
	If InStr(Trim(FileExName),Trim(Mid(FileName.Name,InStr(FileName.Name,".")+1,len(FileName.Name))))>0 Then
    	FileList=FileList&"下载地址：<a href='../Upload/"&FileName.Name&"' target='_blank' style='color:#090'>"&FileName.Name&"</a>　"
	End If
Else
	If InStr(Trim(FileExName),Trim(Mid(FileName.Name,InStr(FileName.Name,".")+1,len(FileName.Name))))>0 Then
    	FileList=FileList&""&FileName.Name&""
	End If
End If
Next
Set file=Nothing
Set folder=Nothing
Set fso=Nothing
End Function
	
action = Trim(Request("action"))
Select Case action
	Case "import"
		Call import()
	Case "gettemplate"
		Call gettemplate()
	Case Else
		Call main()
End Select

sub main()
%>
	<script language="JavaScript">
	<!-- 必填项提示
	function CheckInput()
	{
		if(document.getElementById('excelfile').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '未选择要导入的文件！'});document.getElementById('excelfile').focus();return false;}
		if(document.getElementById('User').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '未选择业务员！'});document.getElementById('User').focus();return false;}
	}
	-->
	</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top" class="td_n pd10"> 
			<form name="linkmansForm" action="?action=import" enctype="multipart/form-data" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="120" />
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="2"><B>外部 Excel 导入数据库</B></td>
				</tr>
				<tr>
					<td class="td_l_r title">上传文件</td>
					<td class="td_r_l"><input name="excelfile" type="file" id="excelfile" value="" maxlength="200" class="int" size="40">　<span class="info_help help01">只允许上传 <B style="color:#ff0000">.xls</b> 格式
					</td>
				</tr>
				<tr>
					<td class="td_l_r title">数据表</td>
					<td class="td_r_l"><select name="tbname" ><option value="client">client</option><option value="Sheet1">Sheet1</option><option value="Sheet2">Sheet2</option><option value="Sheet3">Sheet3</option></select> <span class="info_help help01">位置：Excel文件左下角 </span><span class="info_exceldb">&nbsp;</span></td>
				</tr>
				<tr>
					<td class="td_l_r title">业务员</td>
					<td class="td_r_l"> <% if Session("CRM_level") = 9 then %><% = EasyCrm.UserList(2,"User","") %><%else%><% = EasyCrm.UserList(1,"User","") %><%end if%> <span class="info_help help01">默认导入到某人的库中</span></td>
				</tr>
				<tr>
					<td class="td_l_r title">导入公海</td>
					<td class="td_r_l"> <input type="checkbox" name="GORecycler" value="1"> 是 <span class="info_help help01">勾选后，业务员选择无效</span></td>
				</tr>
				<tr>
					<td class="td_l_r title">注意事项</td>
					<td class="td_r_l" style="padding:5px 10px;"> 
						1、Excel文件仅支持 <B>Office Excel 2003</B>；<BR>
						2、时间格式标准：<B>1900-01-01</B>，不支持其它格式；<BR>
						3、邮编、手机、电话、传真 限数字：“<B>0 1 2 3 4 5 6 7 8 9</B>” 和 “<B>-</B>” ；<BR>
						4、地址、网址、Email 限 <B>255</B> 字节，主营、备注不限；<BR>
						5、其它所有内容限 <B>50</B> 字节。<BR>
						6、建议每次导入数据量在 <B style="color:red;">500</B> 条以内；<BR>
					</td>
				</tr>
			</table>   
			<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:10px;">
				<tr>
					<td colspan="2">
						<input type="submit" name="Submit" class="button45" value=" 开始导入 ">　<input name="gettemplate" type="button" class="button43" value=" 生成模板 " onClick="location.href='?action=gettemplate';">　<%=FileList("../Upload/","xls","link")%>
					</td>
					</tr>
			</table>   
			</form>
		</td> 
	</tr>
</table>   
<%
	Set fso = CreateObject("Scripting.FileSystemObject")
	if FileList("../main/","xls","")<>"" then
	IF fso.FileExists(server.MapPath(FileList("../main/","xls",""))) Then
	fso.DeleteFile server.MapPath(FileList("../main/","xls",""))
	End IF
	End IF
end sub
%>
<%
sub gettemplate()
	on error resume next'如果有错误继续执行下面的代码 
	dim excelfile,tbname
	Dim ExcelDriver,DBExcelPath
	'这里是为了创建sheet1用的
	'Createtable=Createtable&L_Client_cDate&" text ,"&L_Client_cLastUpdated&" text ,"
	
	Createtable=Createtable&L_Client_cDate&" text ,"&L_Client_cLastUpdated&" text ,"&L_Client_cCompany&" text ,"&L_Client_cArea&" text ,"&L_Client_cSquare&" text ,"&L_Client_cAddress&" text ,"&L_Client_cZip&" text ,"&L_Client_cLinkman&" text ,"&L_Client_cZhiwei&" text ,"&L_Client_cMobile&" text ,"&L_Client_cTel&" text ,"&L_Client_cFax&" text ,"&L_Client_cHomepage&" text ,"&L_Client_cEmail&" text ,"&L_Client_cTrade&" text ,"&L_Client_cStrade&" text ,"&L_Client_cType&" text ,"&L_Client_cStart&" text ,"&L_Client_cSource&" text ,"&L_Client_cInfo&" text ,"&L_Client_cBeizhu&" text ,"

	Createtablesql="Create table client("&left(Createtable,len(Createtable)-1)&")"
	ExcelFile="../Upload/Client.xls"
		
	set fso=Server.CreateObject ("Scripting.FileSystemObject")
	fpath=Server.MapPath(ExcelFile)  
	if fso.FileExists(fpath) then
	whichfile=Server.MapPath(ExcelFile)
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set thisfile = fs.GetFile(whichfile)
	thisfile.delete true
	end if             
	Set conn = Server.CreateObject("ADODB.Connection")
	ExcelDriver = "Driver={Microsoft Excel Driver (*.xls)};Readonly=0;"
	DBExcelPath = "DBQ=" & Server.MapPath(excelfile) 
	conn.Open ExcelDriver & DBExcelPath
	conn.Execute(Createtablesql)'在这个conn上执行就得到一个excel
	
	if ""&YNalert&"" = 1 then
	Response.Write ("<script>alert("""&L_Improt_template_alert&""");</script>")
	end if
	Response.Write ("<script>location.href='Import.asp' ;</script>")
end sub
%>
<%
sub import()
	dim nTime : nTime = Timer()
	dim request,lngUpSize
	Set request=new UpLoadClass
		request.TotalSize= 10485760
		request.MaxSize  = 5000*1024
		request.FileType = "xls"
		request.Savepath = ""
	lngUpSize = request.Open()
	
	dim excelfile,tbname,i,lc
	excelfile = request.Savepath & Request.Form("excelfile")
	if excelfile = request.Savepath then excelfile=""
		
	excelfile=request.form("excelfile")
	Username=request.form("User")
	GORecycler=request.form("GORecycler")
	tbname=request.form("tbname")
	
	if right(excelfile,3)<>"xls" then
    Response.Write("<script>alert("""&alert_excelfile&""");history.back(1);</script>")
	Response.End
	end if
	
	dim Connxls,rsxls
	Dim Driver,DBPath
	Set connxls = Server.CreateObject("ADODB.Connection")
	connxls="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(excelfile)&";Extended Properties='Excel 8.0;HDR=yes;IMEX=1';"
		
	Set rsxls=server.createobject("adodb.recordset")
	rsxls.open "select * from ["&tbname&"$]" ,Connxls,1,1

	Set rs=server.createobject("adodb.recordset")
	rs.open "select * from [client] ",conn,3,2

	do while not rsxls.eof '循环读取Excel

	rs.addnew
	
	if rsxls(0) <> "" then
	  rs("cDate") = rsxls(0)
	  else
	  rs("cDate") = date()
	end if
	
	if rsxls(1) <> "" then
	  rs("cLastUpdated") = rsxls(1)
	  else
	  rs("cLastUpdated") = Now()
	end if
	
	if rsxls(2) <> "" then
		dim rs1
		Set rs1 = Server.CreateObject("ADODB.Recordset")
		rs1.Open "Select * From [client] Where cCompany = '" & rsxls(2) & "' ",conn,1,1
		If rs1.RecordCount = 0 Then rs("cCompany") = rsxls(2)
		rs1.Close
	end if
	
	if rsxls(3) <> "" then rs("cArea") = rsxls(3)
	if rsxls(4) <> "" then rs("cSquare") = rsxls(4)
	if rsxls(5) <> "" then rs("cAddress") = rsxls(5)
	if rsxls(6) <> "" then rs("cZip") = rsxls(6)
	if rsxls(7) <> "" then rs("cLinkman") = rsxls(7)
	if rsxls(8) <> "" then rs("cZhiwei") = rsxls(8)
	if rsxls(9) <> "" then rs("cMobile") = rsxls(9)
	if rsxls(10) <> "" then rs("cTel") = rsxls(10)
	if rsxls(11) <> "" then rs("cFax") = rsxls(11)
	if rsxls(12) <> "" then rs("cHomepage") = rsxls(12)
	if rsxls(13) <> "" then rs("cEmail") = rsxls(13)
	if rsxls(14) <> "" then rs("cTrade") = rsxls(14)
	if rsxls(15) <> "" then rs("cStrade") = rsxls(15)
	if rsxls(16) <> "" then rs("cType") = rsxls(16)
	if rsxls(17) <> "" then rs("cStart") = rsxls(17)
	if rsxls(18) <> "" then rs("cSource") = rsxls(18)
	if rsxls(19) <> "" then rs("cInfo") = rsxls(19)
	if rsxls(20) <> "" then rs("cBeizhu") = rsxls(20)
	
	if GORecycler="1" then
	  rs("cGroup") = 1
	  rs("cUser") = "系统公海"
	  rs("cYn") = 0
	else
	  rs("cGroup") = EasyCrm.getNewItem("User","uName","'"&Username&"'","uGroup")
	  rs("cUser") = Username
	  rs("cYn") = 1
	end if
	
	i=i+1
	rsxls.movenext
	loop
	rs.update
	rsxls.close
	rs.close
	
    conn.execute ("Delete from [client] Where cCompany is NULL or cCompany='' ")
    conn.execute ("Delete from [Linkmans] Where lid in (select max(lid) from linkmans group by cid,lname having count(*) > 1) or lName is NULL or lName=''  ")
	
	if ""&YNalert&"" = 1 then
	Response.Write ("<script>alert("""&L_Improt_alert&""");</script>")
	end if
	Response.Write ("<script>location.href='Import.asp' ;</script>")
end sub
%>
</body>
</html><%else%>无权限<%end if%><% Set EasyCrm = nothing %>