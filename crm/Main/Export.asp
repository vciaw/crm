<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 7, 1) = 1 Then %>
<%
action = Trim(Request("action"))
otype	=	Request.QueryString("otype")
if otype="" then otype="Client"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
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
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Company%> > <%=L_Page_Export%></td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
            <div class="MenuboxS">
              <ul>
                <li <%if otype="Client" then%>class="hover"<%end if%>><span><a href="?action=Client&otype=Client"><%=L_Client%></a></span></li>
                <li <%if otype="Records" then%>class="hover"<%end if%>><span><a href="?action=Records&otype=Records"><%=L_Records%></a></span></li>
                <li <%if otype="Order" then%>class="hover"<%end if%>><span><a href="?action=Order&otype=Order"><%=L_Order%></a></span></li>
                <li <%if otype="Hetong" then%>class="hover"<%end if%>><span><a href="?action=Hetong&otype=Hetong"><%=L_Hetong%></a></span></li>
                <li <%if otype="Service" then%>class="hover"<%end if%>><span><a href="?action=Service&otype=Service"><%=L_Service%></a></span></li>
                <li <%if otype="Expense" then%>class="hover"<%end if%>><span><a href="?action=Expense&otype=Expense"><%=L_Expense%></a></span></li>
              </ul>
            </div>
		</td>
	</tr>

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
	Case "Expense"
		Call Expense()
		
	Case "ClienttoExcel"
		Call ClienttoExcel()
	Case "RecordstoExcel"
		Call RecordstoExcel()
	Case "OrdertoExcel"
		Call OrdertoExcel()
	Case "HetongtoExcel"
		Call HetongtoExcel()
	Case "ServicetoExcel"
		Call ServicetoExcel()
	Case "ExpensetoExcel"
		Call ExpensetoExcel()
		
	Case "killSession"
		Call killSession()
	Case Else
		Call Client()
End Select

Sub Client()%>
	<tr> 
		<td valign="top" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" class="td_n pd10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="6"><B><%=L_Top_Search%></B></td>
							</tr>
						</table>
						<form name="searchForm" action="?Action=ClienttoExcel" method="post">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" /><col width="220" /><col width="100" />
							<tr>
								<td class="td_l_c title" style="border-top:0;"><%=L_Client_cType%></td>
								<td class="td_r_l" style="border-top:0;"><% = EasyCrm.getSelect("SelectData","Select_Type","Type","") %></td>
								<td class="td_l_c title" style="border-top:0;"><%=L_Client_cArea%><%=L_Client_cSquare%></td>
								<td class="td_r_l" style="border-top:0;">
									<select name="Area" onchange="getArea(this.options[this.selectedIndex].id);">
									<option value=""><%=L_Please_choose_01%></option>
									<% 
										Set rsb = Conn.Execute("select * from AreaData where aFId = '0' ")
										If Not rsb.Eof then
										Do While Not rsb.Eof
										aId= rsb("aId")
										aName= rsb("aName")
									%>
										<option value="<%=aName%>" id="<%=aId%>"><%=aName%></option>
									<%
										rsb.Movenext
										Loop
										End If
										rsb.Close
										Set rss = Nothing 
									%>
									</select> 
									<span id="Squarediv"  style="margin-left:10px;padding:0;">
										<select name="Squares">
											<option value=""><%=L_Please_choose_02%></option>
										</select>
										
									</span>　
								<input name="Square" type="hidden" id="Square" class="int">
								</td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Client_cSource%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_Source","Source","") %></td>
								<td class="td_l_c title"><%=L_Client_cTrade%></td>
								<td class="td_r_l">
									<select name="Trade" onchange="getTrade(this.options[this.selectedIndex].id);">
									<option value=""><%=L_Please_choose_01%></option>
									<% 
										Set rsb = Conn.Execute("select * from ProductClass where pClassFid = '0' ")
										If Not rsb.Eof then
										Do While Not rsb.Eof
										pClassid= rsb("pClassid")
										pClassname= rsb("pClassname")
									%>
										<option value="<%=pClassname%>" id="<%=pClassid%>"><%=pClassname%></option>
									<%
										rsb.Movenext
										Loop
										End If
										rsb.Close
										Set rsb = Nothing 
									%>
									</select> 
									<span id="Stradediv"  style="margin-left:10px;padding:0;">
										<select name="Strades">
											<option value=""><%=L_Please_choose_02%></option>
										</select>
									</span>
									<input name="Strade" type="hidden" id="Strade" class="int">
								</td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Client_cStart%></td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_Star","Start","") %></td>
								<td class="td_l_c title"><%=L_Client_cDate%></td>
								<td class="td_r_l"> <input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" /></td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Client_cUser%></td>
								<td class="td_r_l"><% If Session("CRM_level") = 9 Then %><% = EasyCrm.UserList(2,"User","") %><%else%><% = EasyCrm.UserList(1,"User","") %><%end if%></td>
								<td class="td_l_c title"><%=L_Client_cLastUpdated%></td>
								<td class="td_r_l"> <input name="ETimeBegin" type="text" id="ETimeBegin" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" />&nbsp;~&nbsp;<input name="ETimeEnd" type="text" id="ETimeEnd" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" /></td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Export_content%></td>
								<td class="td_r_l" colspan=3 style="padding:5px 10px;">
									<input type="checkbox" name="Exportitem1" value="1" checked> <%=L_Client_cId%>　
									<input type="checkbox" name="Exportitem2" value="1" checked> <%=L_Client_cDate%>　
									<input type="checkbox" name="Exportitem3" value="1" checked> <%=L_Client_cLastUpdated%>　
									<input type="checkbox" name="Exportitem4" value="1" checked> <%=L_Client_cCompany%>　
									<input type="checkbox" name="Exportitem5" value="1" checked> <%=L_Client_cAddress%>　
									<input type="checkbox" name="Exportitem6" value="1" checked> <%=L_Client_cTel%>　
									<input type="checkbox" name="Exportitem7" value="1" checked> <%=L_Client_cFax%>　
									<input type="checkbox" name="Exportitem8" value="1" checked> <%=L_Client_cHomepage%>　<BR>
									<input type="checkbox" name="Exportitem9" value="1" checked> <%=L_Client_cEmail%>　
									<input type="checkbox" name="Exportitem10" value="1" checked> <%=L_Client_cTrade%>　
									<input type="checkbox" name="Exportitem11" value="1" checked> <%=L_Client_cType%>　
									<input type="checkbox" name="Exportitem12" value="1" checked> <%=L_Client_cStart%>　
									<input type="checkbox" name="Exportitem13" value="1" checked> <%=L_Client_cSource%>　
									<input type="checkbox" name="Exportitem14" value="1" checked> <%=L_Client_cInfo%>　
									<input type="checkbox" name="Exportitem15" value="1" checked> <%=L_Client_cBeizhu%>　
									<input type="checkbox" name="Exportitem16" value="1" checked> <%=L_Client_cGroup%>　<BR>
									<input type="checkbox" name="Exportitem17" value="1" checked> <%=L_Client_cUser%>　
									<input type="checkbox" name="Exportitem18" value="1" checked> <%=L_Client_cLinkman%>　
									<input type="checkbox" name="Exportitem19" value="1" checked> <%=L_Client_cZhiwei%>　
									<input type="checkbox" name="Exportitem20" value="1" checked> <%=L_Client_cMobile%>　
									<input type="checkbox" name="Exportitem21" value="1" checked> <%=L_Client_cRNextTime%>　
									<input type="checkbox" name="Exportitem22" value="1" checked> <%=L_Client_cOEDate%>　
									<input type="checkbox" name="Exportitem23" value="1" checked> <%=L_Client_cHEdate%>　
								</td>
							</tr>
						</table> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:10px;">
							<tr>
								<td>
									<input type="submit" name="Submit" class="button45" value=" <%=L_Export%> ">　
									<input type="button" class="button42" value=" 下载文件 " onClick=window.location.href="../Soft/index.asp" />　
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?action=killSession&otype=Client" />　
								</td>
							</tr>
						</table>  
						</form>  
					</td> 
				</tr>
			</table>  
		</td> 
	</tr>
</table>
</body>
</html>
<script src="../data/calendar/WdatePicker.js"></script>
<% end sub

sub ClienttoExcel()

	Dim cType,cArea,cSquare,cSource,cTrade,cStrade,cStart,cUser,arrUser,cTimeBegin,cTimeEnd,ETimeBegin,ETimeEnd	
	cType = Request("type")
	cArea = Request("area")
	cSquare = Request("Square")
	cSource = Request("Source")
	cTrade = Request("trade")
	cStrade = Request("Strade")
	cStart = Request("Start")
	cUser = Request("user")
	cTimeBegin = Request("TimeBegin")
	cTimeEnd = Request("TimeEnd")
	ETimeBegin = Request("ETimeBegin")
	ETimeEnd = Request("ETimeEnd")
	
	arrUser = getUserList(Session("CRM_level"),Session("CRM_group"),Session("CRM_MR"))
	
	Dim sql
    sql = ""
	
    If cType <> "" Then
		sql = sql & " And cType = '" & cType & "'"
	End If
		
    If cArea <> "" Then
		sql = sql & " and cArea = '" & cArea & "'"
	End If
		
    If cSquare <> "" Then
		sql = sql & " And cSquare = '" & cSquare & "'"
	End If
		
    If cSource <> "" Then
		sql = sql & " And cSource = '" & cSource & "'"
	End If
	
    If cTrade <> "" Then
		sql = sql & " And cTrade = '" & cTrade & "'"
	End If
		
    If cStrade <> "" Then
		sql = sql & " And cStrade = '" & cStrade & "'"
	End If
	
    If cStart <> "" Then
		sql = sql & " And cStart = '" & cStart & "'"
	End If
			
	if Accsql =1 then
	If cTimeBegin <> "" Then
		sql = sql & " And cdate >= '" & cTimeBegin & "' "
	End If
			
	If cTimeEnd <> "" Then
		sql = sql & " And cdate <= '" & cTimeEnd & "' "
	End If
			
	If ETimeBegin <> "" Then
		sql = sql & " And cLastUpdated >= '" & ETimeBegin & "' "
	End If
			
	If ETimeEnd <> "" Then
		sql = sql & " And cLastUpdated <= '" & ETimeEnd & "' "
	End If
	else
	If cTimeBegin <> "" Then
		sql = sql & " And cdate >= #" & cTimeBegin & "# "
	End If
			
	If cTimeEnd <> "" Then
		sql = sql & " And cdate <= #" & cTimeEnd & "# "
	End If
			
	If ETimeBegin <> "" Then
		sql = sql & " And cLastUpdated >= #" & ETimeBegin & "# "
	End If
			
	If ETimeEnd <> "" Then
		sql = sql & " And cLastUpdated <= #" & ETimeEnd & "# "
	End If
	End If
			
	If cUser <> "" Then
		sql = sql & " And cUser = '" & cUser & "' "
	End If
	
	If Session("CRM_level") < 9 Then
         sql = sql & " And cUser In ( " & arrUser & " )"
	End If

If cType = "" And cArea = "" And cSquare = "" And cTrade = "" And cStrade = "" And cSource = "" And cStart = "" And cUser = "" And cTimeBegin = "" And cTimeEnd = "" And ETimeBegin = "" And ETimeEnd = "" Then
    If Session("CRM_sql_Export_Client") <> "" Then
        sql = Session("CRM_sql_Export_Client")
	Else
	    If Session("CRM_level") < 9 Then
		    sql = " And cUser In ( " & arrUser & " ) "
		End If
	End If
Else
    Session("CRM_sql_Export_Client") = sql
End If

userfolder = Session("CRM_account") '生成文件夹
filefolder = Server.MapPath("../soft/"&userfolder)
set fso = Server.CreateObject("Scripting.FileSystemObject")
if not fso.FolderExists(filefolder) then '如果文件夹不存在则创建
fso.CreateFolder(filefolder) 
end if

	Exportitem = ""
	for i = 1 to 23
		if Request("Exportitem" & i) = "1" then
			Exportitem = Exportitem & "1"
		else
			Exportitem = Exportitem & "0"
		end if
	next
	
    strLine = strLine & "cId as 编号"
	
	if mid(Exportitem, 2, 1) = "1" then
	strLine = strLine & ", cDate as 录入时间"
	end if
	if mid(Exportitem, 3, 1) = "1" then
	strLine = strLine & ", cLastUpdated as 最后更新"
	end if
	if mid(Exportitem, 4, 1) = "1" then
	strLine = strLine & ", cCompany as 公司名称"
	end if
	if mid(Exportitem, 5, 1) = "1" then
	strLine = strLine & ", cArea as 省份"
	strLine = strLine & ", cSquare as 地区"
	strLine = strLine & ", cAddress as 详细地址"
	strLine = strLine & ", cZip as 邮编"
	end if
	if mid(Exportitem, 18, 1) = "1" then
	strLine = strLine & ", cLinkman as 联系人"
	end if
	if mid(Exportitem, 19, 1) = "1" then
	strLine = strLine & ", cZhiwei as 职位"
	end if
	if mid(Exportitem, 20, 1) = "1" then
	strLine = strLine & ", cMobile as 手机"
	end if
	if mid(Exportitem, 21, 1) = "1" then
	strLine = strLine & ", cRNextTime as 下次跟进时间"
	end if
	if mid(Exportitem, 22, 1) = "1" then
	strLine = strLine & ", cOEDate as 交单时间"
	end if
	if mid(Exportitem, 23, 1) = "1" then
	strLine = strLine & ", cHEdate as 合同到期"
	end if
	if mid(Exportitem, 6, 1) = "1" then
	strLine = strLine & ", cTel as 电话"
	end if
	if mid(Exportitem, 7, 1) = "1" then
	strLine = strLine & ", cFax as 传真"
	end if
	if mid(Exportitem, 8, 1) = "1" then
	strLine = strLine & ", cHomepage as 网址"
	end if
	if mid(Exportitem, 9, 1) = "1" then
	strLine = strLine & ", cEmail as 邮箱"
	end if
	if mid(Exportitem, 10, 1) = "1" then
	strLine = strLine & ", cTrade as 产品大类"
	strLine = strLine & ", cStrade as 产品小类"
	end if
	if mid(Exportitem, 11, 1) = "1" then
	strLine = strLine & ", cType as 客户类型"
	end if
	if mid(Exportitem, 12, 1) = "1" then
	strLine = strLine & ", cStart as 星级"
	end if
	if mid(Exportitem, 13, 1) = "1" then
	strLine = strLine & ", cSource as 来源"
	end if
	if mid(Exportitem, 14, 1) = "1" then
	strLine = strLine & ", cInfo as 主营项目"
	end if
	if mid(Exportitem, 15, 1) = "1" then
	strLine = strLine & ", cBeizhu as 详情备注"
	end if
	strLine = strLine & ", cGroup as 部门"
	strLine = strLine & ", cUser as 业务员"

on error resume next'如果有错误继续执行下面的代码 
dim excelfile,tbname
Dim ExcelDriver,DBExcelPath
tbname="Client"
Server.ScriptTimeOut=360000'防止超时
set rs=server.createobject("adodb.recordset")
sql="select "&strLine&" from ["&tbname&"] where cYn = 1 "&sql&" Order By cId desc "'根据此SQL语句导出至Excel
rs.Open sql,conn,3,3
for Createtablei=0 to rs.Fields.Count-1                 '这里是为了创建sheet1用的
Createtable=Createtable&rs.fields(Createtablei).name&" text ,"
next

Createtablesql="Create table Sheet1("&left(Createtable,len(Createtable)-1)&")"
ExcelFile="../soft/"&userfolder&"/"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls" 
	'同步写入文件柜
	conn.execute "insert into OA_soft(s_class,s_title,s_file,s_user,s_share,s_time) values('"&L_Export_soft&"','"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls','"&ExcelFile&"','"&Session("CRM_name")&"','0','"&now()&"')"
	
set fso=Server.CreateObject ("Scripting.FileSystemObject")
fpath=Server.MapPath(ExcelFile)  
if fso.FileExists(fpath) then
whichfile=Server.MapPath(ExcelFile)
Set fs = CreateObject("Scripting.FileSystemObject")
Set thisfile = fs.GetFile(whichfile)
thisfile.delete true
end if             
Set connXLS = Server.CreateObject("ADODB.Connection")
ExcelDriver = "Driver={Microsoft Excel Driver (*.xls)};Readonly=0;"
DBExcelPath = "DBQ=" & Server.MapPath(excelfile) 
connXLS.Open ExcelDriver & DBExcelPath
connXLS.Execute(Createtablesql)'在这个conn上执行就得到一个excel
for ii=0 to rs.recordcount-1
for i=0 to rs.Fields.Count-1
   Inserttablename=Inserttablename&rs.fields(i).name&","
   if i=rs.Fields.Count-2 then
   Inserttable=Inserttable&"'"&EasyCrm.getNewItem("system_group","gId",""&Rs(i)&"","gName")&"',"
   else
   Inserttable=Inserttable&"'"&Rs(i)&"',"
   end if
Next 
Insertintosql="Insert into Sheet1("&left(Inserttablename,len(Inserttablename)-1)&")values("&left(Inserttable,len(Inserttable)-1)&")"
connXLS.Execute(Insertintosql)
Insertintosql =""
Inserttable=""
Inserttablename=""
rs.MoveNext
Next

Session("CRM_sql_Export_Client") = ""
	
if ""&YNalert&"" = 1 then
	Response.Write ("<script>alert("""&L_Export_alert&""");</script>")
end if
Response.Write ("<script>location.href='Export.asp' ;</script>")

end sub

Sub Records()%>
	<tr> 
		<td valign="top" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="6"><B><%=L_Top_Search%></B></td>
							</tr>
						</table>
						<form name="searchForm" action="?Action=RecordstoExcel" method="post">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" /><col width="220" /><col width="100" />
							<tr>
								<td class="td_l_c title" style="border-top:0;"><%=L_Records_rType%></td>
								<td class="td_r_l" style="border-top:0;"> <% = EasyCrm.getSelect("SelectData","Select_Records","rType","") %></td>
								<td class="td_l_c title" style="border-top:0;"><%=L_Records_rState%></td>
								<td class="td_r_l" style="border-top:0;"> <% = EasyCrm.getSelect("SelectData","Select_Type","rState","") %></td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Records_rUser%></td>
								<td class="td_r_l"> <% if Session("CRM_level") = 9 then %><% = EasyCrm.UserList(2,"User","") %><%else%><% = EasyCrm.UserList(1,"User","") %><%end if%></td>
								<td class="td_l_c title"><%=L_Records_rTime%></td>
								<td class="td_r_l"> <input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" style="width:150px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" style="width:150px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" /></td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Export_content%></td>
								<td class="td_r_l" colspan=3>
									<input type="checkbox" name="Exportitem1" value="1" checked> <%=L_Records_rId%>　
									<input type="checkbox" name="Exportitem2" value="1" checked> <%=L_Records_cId%>　
									<input type="checkbox" name="Exportitem3" value="1" checked> <%=L_Records_rType%>　
									<input type="checkbox" name="Exportitem4" value="1" checked> <%=L_Records_rState%>　
									<input type="checkbox" name="Exportitem5" value="1" checked> <%=L_Records_rLinkman%>　
									<input type="checkbox" name="Exportitem6" value="1" checked> <%=L_Records_rNextTime%>　
									<input type="checkbox" name="Exportitem7" value="1" checked> <%=L_Records_rContent%>　
									<input type="checkbox" name="Exportitem8" value="1" checked> <%=L_Records_rUser%>　
									<input type="checkbox" name="Exportitem9" value="1" checked> <%=L_Records_rTime%>　
								</td>
							</tr>
						</table>  
						<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:10px;">
							<tr>
								<td>
									<input type="submit" name="Submit" class="button45" value=" <%=L_Export%> ">　
									<input type="button" class="button42" value=" 下载文件 " onClick=window.location.href="../Soft/index.asp" />　
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?action=killSession&otype=Records" />　
								</td>
							</tr>
						</table>  
						</form> 
					</td> 
				</tr>
			</table>  
		</td> 
	</tr>
</table>
</body>
</html>
<script src="../data/calendar/WdatePicker.js"></script>
<%
end Sub

sub RecordstoExcel()

	Dim rType,rState,rUser,arrUser,rTimeBegin,rTimeEnd
	rType = Request("rType")
	rState = Request("rState")
	rUser = Request("user")
	rTimeBegin = Request("TimeBegin")
	rTimeEnd = Request("TimeEnd")
	
	arrUser = getUserList(Session("CRM_level"),Session("CRM_group"),Session("CRM_MR"))
	
	Dim sql
    sql = ""
	
    If rType <> "" Then
		sql = sql & " And rType = '" & rType & "'"
	End If
		
    If rState <> "" Then
		sql = sql & " and rState = '" & rState & "'"
	End If
			
	if Accsql =1 then
	If rTimeBegin <> "" Then
		sql = sql & " And rTime >= '" & rTimeBegin & "' "
	End If
			
	If rTimeEnd <> "" Then
		sql = sql & " And rTime <= '" & rTimeEnd & "' "
	End If
	else
	If rTimeBegin <> "" Then
		sql = sql & " And rTime >= #" & rTimeBegin & "# "
	End If
			
	If rTimeEnd <> "" Then
		sql = sql & " And rTime <= #" & rTimeEnd & "# "
	End If
	end if
			
	If rUser <> "" Then
		sql = sql & " And rUser = '" & rUser & "' "
	End If
	
	If Session("CRM_level") < 9 Then
        sql = sql & " And rUser In (" & arrUser & ")"
	End If

If rType = "" And rState = "" And rUser = "" And rTimeBegin = "" And rTimeEnd = "" Then
    If Session("CRM_sql_Export_Records") <> "" Then
        sql = Session("CRM_sql_Export_Records")
	Else
	    If Session("CRM_level") < 9 Then
		    sql = " and rUser In ( " & arrUser & " ) "
		End If
	End If
Else
    Session("CRM_sql_Export_Records") = sql
End If

userfolder = Session("CRM_account") '生成文件夹
filefolder = Server.MapPath("../soft/"&userfolder)
set fso = Server.CreateObject("Scripting.FileSystemObject")
if not fso.FolderExists(filefolder) then '如果文件夹不存在则创建
fso.CreateFolder(filefolder) 
end if

	Exportitem = ""
	for i = 1 to 9
		if Request("Exportitem" & i) = "1" then
			Exportitem = Exportitem & "1"
		else
			Exportitem = Exportitem & "0"
		end if
	next
	
	if mid(Exportitem, 1, 1) = "1" then
    strLine = strLine & "rId as 编号"
	end if
	if mid(Exportitem, 2, 1) = "1" then
	strLine = strLine & ", cId as 公司名称"
	end if
	if mid(Exportitem, 3, 1) = "1" then
	strLine = strLine & ", rType as 跟单类型"
	end if
	if mid(Exportitem, 4, 1) = "1" then
	strLine = strLine & ", rState as 跟单进度"
	end if
	if mid(Exportitem, 5, 1) = "1" then
	strLine = strLine & ", rLinkman as 联系人"
	end if
	if mid(Exportitem, 6, 1) = "1" then
	strLine = strLine & ", rNextTime as 下次跟进时间"
	end if
	if mid(Exportitem, 7, 1) = "1" then
	strLine = strLine & ", rContent as 详情备注"
	end if
	if mid(Exportitem, 8, 1) = "1" then
	strLine = strLine & ", rUser as 业务员"
	end if
	if mid(Exportitem, 9, 1) = "1" then
	strLine = strLine & ", rTime as 录入时间"
	end if

on error resume next'如果有错误继续执行下面的代码 
dim excelfile,tbname
Dim ExcelDriver,DBExcelPath
tbname="Records"
Server.ScriptTimeOut=360000'防止超时
set rs=server.createobject("adodb.recordset")     
sql="select "&strLine&" from ["&tbname&"] where 1 = 1 "&sql&" Order By rId desc "'根据此SQL语句导出至Excel
rs.Open sql,conn,3,3
for Createtablei=0 to rs.Fields.Count-1                 '这里是为了创建sheet1用的
Createtable=Createtable&rs.fields(Createtablei).name&" text ,"
next

Createtablesql="Create table Sheet1("&left(Createtable,len(Createtable)-1)&")"
ExcelFile="../soft/"&userfolder&"/"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls" 
	'同步写入文件柜
	conn.execute "insert into OA_soft(s_class,s_title,s_file,s_user,s_share,s_time) values('"&L_Export_soft&"','"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls','"&ExcelFile&"','"&Session("CRM_name")&"','0','"&now()&"')"
set fso=Server.CreateObject ("Scripting.FileSystemObject")
fpath=Server.MapPath(ExcelFile)  
if fso.FileExists(fpath) then
whichfile=Server.MapPath(ExcelFile)
Set fs = CreateObject("Scripting.FileSystemObject")
Set thisfile = fs.GetFile(whichfile)
thisfile.delete true
end if             
Set connXLS = Server.CreateObject("ADODB.Connection")
ExcelDriver = "Driver={Microsoft Excel Driver (*.xls)};Readonly=0;"
DBExcelPath = "DBQ=" & Server.MapPath(excelfile) 
connXLS.Open ExcelDriver & DBExcelPath
connXLS.Execute(Createtablesql)'在这个conn上执行就得到一个excel
	
for ii=0 to rs.recordcount-1
for i=0 to rs.Fields.Count-1
	'cCompany = EasyCrm.getNewItem("Client","cID","'"&Rs(i)&"'","cCompany")
   Inserttablename=Inserttablename&rs.fields(i).name&","
   if i=1 then
   Inserttable=Inserttable&"'"&EasyCrm.getNewItem("Client","cID",""&Rs(i)&"","cCompany")&"',"
   else
   Inserttable=Inserttable&"'"&Rs(i)&"',"
   end if
Next 
Insertintosql="Insert into Sheet1("&left(Inserttablename,len(Inserttablename)-1)&")values("&left(Inserttable,len(Inserttable)-1)&")"
connXLS.Execute(Insertintosql)
Insertintosql =""
Inserttable=""
Inserttablename=""
rs.MoveNext
Next

Session("CRM_sql_Export_Records") = ""

if ""&YNalert&"" = 1 then
	Response.Write ("<script>alert("""&L_Export_alert&""");</script>")
end if
	Response.Write ("<script>location.href='?action=Records&otype=Records' ;</script>")
end Sub

Sub Order()%>
	<tr> 
		<td valign="top" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="6"><B><%=L_Top_Search%></B></td>
							</tr>
						</table>
						<form name="searchForm" action="?Action=OrdertoExcel" method="post">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" /><col width="220" /><col width="100" />
							<tr>
								<td class="td_l_c title" style="border-top:0;"><%=L_Order_oState%></td>
								<td class="td_r_l" style="border-top:0;"> <select name='oState'><option value="">请选择</option><option value="0">未处理</option><option value="1">处理中</option><option value="2">已完成</option></select></td>
								<td class="td_l_c title" style="border-top:0;"><%=L_Order_oSDate%></td>
								<td class="td_r_l" style="border-top:0;"> <input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" /></td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Order_oUser%></td>
								<td class="td_r_l"> <% if Session("CRM_level") = 9 then %><% = EasyCrm.UserList(2,"User","") %><%else%><% = EasyCrm.UserList(1,"User","") %><%end if%></td>
								<td class="td_l_c title"><%=L_Order_oEDate%></td>
								<td class="td_r_l"> <input name="ETimeBegin" type="text" id="ETimeBegin" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" />&nbsp;~&nbsp;<input name="ETimeEnd" type="text" id="ETimeEnd" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" /></td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Export_content%></td>
								<td class="td_r_l" colspan=3>
									<input type="checkbox" name="Exportitem1" value="1" checked> <%=L_Order_oId%></font>　
									<input type="checkbox" name="Exportitem2" value="1" checked> <%=L_Order_cId%>　
									<input type="checkbox" name="Exportitem9" value="1" checked> <%=L_Order_oState%>
									<input type="checkbox" name="Exportitem3" value="1" checked> <%=L_Order_oCode%>　
									<input type="checkbox" name="Exportitem4" value="1" checked> <%=L_Order_oLinkman%>　
									<input type="checkbox" name="Exportitem5" value="1" checked> <%=L_Order_oSDate%>　
									<input type="checkbox" name="Exportitem6" value="1" checked> <%=L_Order_oEDate%>　<br>
									<input type="checkbox" name="Exportitem7" value="1" checked> <%=L_Order_oDeposit%>　
									<input type="checkbox" name="Exportitem8" value="1" checked> <%=L_Order_oMoney%>　
									<input type="checkbox" name="Exportitem10" value="1" checked> <%=L_Order_oContent%>　
									<input type="checkbox" name="Exportitem11" value="1" checked> <%=L_Order_oUser%>　
									<input type="checkbox" name="Exportitem12" value="1" checked> <%=L_Order_oTime%>　
								</td>
							</tr>
						</table>  
						<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:10px;">
							<tr>
								<td>
									<input type="submit" name="Submit" class="button45" value=" <%=L_Export%> ">　
									<input type="button" class="button42" value=" 下载文件 " onClick=window.location.href="../Soft/index.asp" />　
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?action=killSession&otype=Records" />　
								</td>
							</tr>
						</table>  
						</form> 
					</td> 
				</tr>
			</table>  
		</td> 
	</tr>
</table>
</body>
</html>
<script src="../data/calendar/WdatePicker.js"></script>
<%
end Sub

sub OrdertoExcel()

	Dim oState,oUser,arrUser,oTimeBegin,oTimeEnd,oETimeBegin,oETimeEnd
	oState = Request("oState")
	oUser = Request("User")
	oTimeBegin = Request("TimeBegin")
	oTimeEnd = Request("TimeEnd")
	oETimeBegin = Request("ETimeBegin")
	oETimeEnd = Request("ETimeEnd")
	
	arrUser = getUserList(Session("CRM_level"),Session("CRM_group"),Session("CRM_MR"))
	
	Dim sql
    sql = ""
	
    If oState <> "" Then
		sql = sql & " And oState = '" & oState & "'"
	End If
		
    If oUser <> "" Then
		sql = sql & " and oUser = '" & rState & "'"
	End If

	if Accsql =1 then
	If oTimeBegin <> "" and  oTimeEnd <> "" Then
	    sql = sql & " And oSDate >= '" & oTimeBegin & "' And oSDate <= '" & oTimeEnd & "' "
	End If
	If oTimeBegin <> "" and  oTimeEnd = "" Then
	    sql = sql & " And oSDate = '" & oTimeBegin & "' "
	End If
	else
	If oTimeBegin <> "" and oTimeEnd <> "" Then
	    sql = sql & " And oSDate >= #" & oTimeBegin & "# And oSDate <= #" & oTimeEnd & "# "
	End If
	If oTimeBegin <> "" and  oTimeEnd = "" Then
	    sql = sql & " And oSDate = #" & oTimeBegin & "# "
	End If
	end if

	if Accsql =1 then
	If oETimeBegin <> "" and  oETimeEnd <> "" Then
	    sql = sql & " And oEDate >= '" & oETimeBegin & "' And oEDate <= '" & oETimeEnd & "' "
	End If
	If oETimeBegin <> "" and  oETimeEnd = "" Then
	    sql = sql & " And oEDate = '" & oETimeBegin & "' "
	End If
	else
	If oETimeBegin <> "" and oETimeEnd <> "" Then
	    sql = sql & " And oEDate >= #" & oETimeBegin & "# And oEDate <= #" & oETimeEnd & "# "
	End If
	If oETimeBegin <> "" and  oETimeEnd = "" Then
	    sql = sql & " And oEDate = #" & oETimeBegin & "# "
	End If
	end if
	
	If Session("CRM_level") < 9 Then
        sql = sql & " And oUser In (" & arrUser & ")"
	End If

If oState = "" And oUser = "" And oTimeBegin = "" And oTimeEnd = "" And oETimeBegin = "" And oETimeEnd = "" Then
    If Session("CRM_sql_Export_Order") <> "" Then
        sql = Session("CRM_sql_Export_Order")
	Else
	    If Session("CRM_level") < 9 Then
		    sql = " and oUser In ( " & arrUser & " ) "
		End If
	End If
Else
    Session("CRM_sql_Export_Order") = sql
End If

userfolder = Session("CRM_account") '生成文件夹
filefolder = Server.MapPath("../soft/"&userfolder)
set fso = Server.CreateObject("Scripting.FileSystemObject")
if not fso.FolderExists(filefolder) then '如果文件夹不存在则创建
fso.CreateFolder(filefolder) 
end if

	Exportitem = ""
	for i = 1 to 12
		if Request("Exportitem" & i) = "1" then
			Exportitem = Exportitem & "1"
		else
			Exportitem = Exportitem & "0"
		end if
	next
	
    strLine = strLine & "oId as 编号"
	strLine = strLine & ", cId as 公司名称"
	strLine = strLine & ", oState as 订单状态"
	if mid(Exportitem, 3, 1) = "1" then
	strLine = strLine & ", oCode as 订单编号"
	end if
	if mid(Exportitem, 4, 1) = "1" then
	strLine = strLine & ", oLinkman as 联系人"
	end if
	if mid(Exportitem, 5, 1) = "1" then
	strLine = strLine & ", oSDate as 下单时间"
	end if
	if mid(Exportitem, 6, 1) = "1" then
	strLine = strLine & ", oEDate as 交单时间"
	end if
	if mid(Exportitem, 7, 1) = "1" then
	strLine = strLine & ", oDeposit as 预付款"
	end if
	if mid(Exportitem, 8, 1) = "1" then
	strLine = strLine & ", oMoney as 总金额"
	end if
	if mid(Exportitem, 10, 1) = "1" then
	strLine = strLine & ", oContent as 详情备注"
	end if
	if mid(Exportitem, 11, 1) = "1" then
	strLine = strLine & ", oUser as 业务员"
	end if
	if mid(Exportitem, 12, 1) = "1" then
	strLine = strLine & ", oTime as 录入时间"
	end if

on error resume next'如果有错误继续执行下面的代码 
dim excelfile,tbname
Dim ExcelDriver,DBExcelPath
tbname="Order"
Server.ScriptTimeOut=360000'防止超时
set rs=server.createobject("adodb.recordset")     
sql="select "&strLine&" from ["&tbname&"] where 1 = 1 "&sql&" Order By oId desc "'根据此SQL语句导出至Excel
rs.Open sql,conn,3,3
for Createtablei=0 to rs.Fields.Count-1                 '这里是为了创建sheet1用的
Createtable=Createtable&rs.fields(Createtablei).name&" text ,"
next

Createtablesql="Create table Sheet1("&left(Createtable,len(Createtable)-1)&")"
ExcelFile="../soft/"&userfolder&"/"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls" 
	'同步写入文件柜
	conn.execute "insert into OA_soft(s_class,s_title,s_file,s_user,s_share,s_time) values('"&L_Export_soft&"','"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls','"&ExcelFile&"','"&Session("CRM_name")&"','0','"&now()&"')"
set fso=Server.CreateObject ("Scripting.FileSystemObject")
fpath=Server.MapPath(ExcelFile)  
if fso.FileExists(fpath) then
whichfile=Server.MapPath(ExcelFile)
Set fs = CreateObject("Scripting.FileSystemObject")
Set thisfile = fs.GetFile(whichfile)
thisfile.delete true
end if             
Set connXLS = Server.CreateObject("ADODB.Connection")
ExcelDriver = "Driver={Microsoft Excel Driver (*.xls)};Readonly=0;"
DBExcelPath = "DBQ=" & Server.MapPath(excelfile) 
connXLS.Open ExcelDriver & DBExcelPath
connXLS.Execute(Createtablesql)'在这个conn上执行就得到一个excel
	
for ii=0 to rs.recordcount-1
for i=0 to rs.Fields.Count-1
	'cCompany = EasyCrm.getNewItem("Client","cID","'"&Rs(i)&"'","cCompany")
   Inserttablename=Inserttablename&rs.fields(i).name&","
   if i=1 then
   Inserttable=Inserttable&"'"&EasyCrm.getNewItem("Client","cID",""&Rs(i)&"","cCompany")&"',"
   elseif i=2 then
		if Rs(i)=0 then
			Inserttable=Inserttable&"'未处理',"
		elseif Rs(i)=1 then
			Inserttable=Inserttable&"'处理中',"
		elseif Rs(i)=2 then
			Inserttable=Inserttable&"'已完成',"
		end if
   else
   Inserttable=Inserttable&"'"&Rs(i)&"',"
   end if
Next 
Insertintosql="Insert into Sheet1("&left(Inserttablename,len(Inserttablename)-1)&")values("&left(Inserttable,len(Inserttable)-1)&")"
connXLS.Execute(Insertintosql)
Insertintosql =""
Inserttable=""
Inserttablename=""
rs.MoveNext
Next

Session("CRM_sql_Export_Order") = ""

if ""&YNalert&"" = 1 then
	Response.Write ("<script>alert("""&L_Export_alert&""");</script>")
end if
	Response.Write ("<script>location.href='?action=Order&otype=Order' ;</script>")
end Sub

Sub Hetong()%>
	<tr> 
		<td valign="top" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="6"><B><%=L_Top_Search%></B></td>
							</tr>
						</table>
						<form name="searchForm" action="?Action=HetongtoExcel" method="post">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" /><col width="220" /><col width="100" />
							<tr>
								<td class="td_l_c title" style="border-top:0;"><%=L_Hetong_hType%></td>
								<td class="td_r_l" style="border-top:0;"> <% = EasyCrm.getSelect("SelectData","Select_Hetong","hType","") %></td>
								<td class="td_l_c title" style="border-top:0;"><%=L_Hetong_hSdate%></td>
								<td class="td_r_l" style="border-top:0;"> <input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" /></td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Hetong_hState%></td>
								<td class="td_r_l"> <select name='hState'><option value="">请选择</option><option value="<%=L_Hetong_hState_1%>"><%=L_Hetong_hState_1%></option><option value="<%=L_Hetong_hState_2%>"><%=L_Hetong_hState_2%></option><option value="<%=L_Hetong_hState_3%>"><%=L_Hetong_hState_3%></option></select></td>
								<td class="td_l_c title"><%=L_Hetong_hEdate%></td>
								<td class="td_r_l"> <input name="ETimeBegin" type="text" id="ETimeBegin" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" />&nbsp;~&nbsp;<input name="ETimeEnd" type="text" id="ETimeEnd" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" /></td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Hetong_hUser%></td>
								<td class="td_r_l"> <% if Session("CRM_level") = 9 then %><% = EasyCrm.UserList(2,"User","") %><%else%><% = EasyCrm.UserList(1,"User","") %><%end if%></td>
								<td class="td_l_c title"><%=L_You&L_Wu&L_Hetong_hOwed%></td>
								<td class="td_r_l">
									<input name="hMoney" type="radio" class="noborder" value="" checked><%=L_Export_hOwed_all%>　
									<input name="hMoney" type="radio" class="noborder" value="0"><%=L_Export_hOwed_0%>　
									<input name="hMoney" type="radio" class="noborder" value="1"><%=L_Export_hOwed_1%>　
								</td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Export_content%></td>
								<td class="td_r_l" colspan=3>
									<input type="checkbox" name="Exportitem1" value="1" checked> <%=L_Hetong_hId%>　
									<input type="checkbox" name="Exportitem2" value="1" checked> <%=L_Hetong_cId%>　
									<input type="checkbox" name="Exportitem3" value="1" checked> <%=L_Hetong_oId%>　
									<input type="checkbox" name="Exportitem4" value="1" checked> <%=L_Hetong_hNum%>　
									<input type="checkbox" name="Exportitem5" value="1" checked> <%=L_Hetong_hSdate%>　
									<input type="checkbox" name="Exportitem6" value="1" checked> <%=L_Hetong_hEdate%>　<br>
									<input type="checkbox" name="Exportitem7" value="1" checked> <%=L_Hetong_hType%>　
									<input type="checkbox" name="Exportitem8" value="1" checked> <%=L_Hetong_hRevenue%>　
									<input type="checkbox" name="Exportitem9" value="1" checked> <%=L_Hetong_hOwed%>　
									<input type="checkbox" name="Exportitem10" value="1" checked> <%=L_Hetong_hMoney%>　
									<input type="checkbox" name="Exportitem11" value="1" checked> <%=L_Hetong_hInvoice%>　
									<input type="checkbox" name="Exportitem12" value="1" checked> <%=L_Hetong_hTax%>　
									<input type="checkbox" name="Exportitem13" value="1" checked> <%=L_Hetong_hState%>　<br>
									<input type="checkbox" name="Exportitem14" value="1" checked> <%=L_Hetong_hContent%>　
									<input type="checkbox" name="Exportitem15" value="1" checked> <%=L_Hetong_hAudit%>　
									<input type="checkbox" name="Exportitem16" value="1" checked> <%=L_Hetong_hAuditTime%>　
									<input type="checkbox" name="Exportitem17" value="1" checked> <%=L_Hetong_hAuditReasons%>　
									<input type="checkbox" name="Exportitem18" value="1" checked> <%=L_Hetong_hUser%>　
									<input type="checkbox" name="Exportitem19" value="1" checked> <%=L_Hetong_hTime%>　
								</td>
							</tr>
						</table>  
						<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:10px;">
							<tr>
								<td>
									<input type="submit" name="Submit" class="button45" value=" <%=L_Export%> ">　
									<input type="button" class="button42" value=" 下载文件 " onClick=window.location.href="../Soft/index.asp" />　
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?action=killSession&otype=Records" />　
								</td>
							</tr>
						</table>  
						</form> 
					</td> 
				</tr>
			</table>  
		</td> 
	</tr>
</table>
</body>
</html>
<script src="../data/calendar/WdatePicker.js"></script>
<%
end Sub

sub HetongtoExcel()

	Dim hType,hState,hMoney,hUser,arrUser,hTimeBegin,hTimeEnd,hETimeBegin,hETimeEnd
	hType = Request("hType")
	hState = Request("hState")
	hMoney = Request("hMoney")
	hUser = Request("User")
	hTimeBegin = Request("TimeBegin")
	hTimeEnd = Request("TimeEnd")
	hETimeBegin = Request("ETimeBegin")
	hETimeEnd = Request("ETimeEnd")
	
	arrUser = getUserList(Session("CRM_level"),Session("CRM_group"),Session("CRM_MR"))
	
	Dim sql
    sql = ""
	
    If hType <> "" Then
		sql = sql & " And hType = '" & hType & "'"
	End If
	
    If hState <> "" Then
		sql = sql & " And hState = '" & hState & "'"
	End If
	
    If hMoney <> "" Then
		If hMoney = 0 then 
		sql = sql & " And hOwed = 0 "
		elseif hMoney = 1 then
		sql = sql & " And hOwed > 0 "
		end if
	End If
		
    If hUser <> "" Then
		sql = sql & " and hUser = '" & hUser & "'"
	End If

	if Accsql =1 then
	If hTimeBegin <> "" and  hTimeEnd <> "" Then
	    sql = sql & " And hSdate >= '" & hTimeBegin & "' And hSdate <= '" & hTimeEnd & "' "
	End If
	If hTimeBegin <> "" and  hTimeEnd = "" Then
	    sql = sql & " And hSdate = '" & hTimeBegin & "' "
	End If
	else
	If hTimeBegin <> "" and hTimeEnd <> "" Then
	    sql = sql & " And hSdate >= #" & hTimeBegin & "# And hSdate <= #" & hTimeEnd & "# "
	End If
	If hTimeBegin <> "" and  hTimeEnd = "" Then
	    sql = sql & " And hSdate = #" & hTimeBegin & "# "
	End If
	end if

	if Accsql =1 then
	If hETimeBegin <> "" and  hETimeEnd <> "" Then
	    sql = sql & " And hEdate >= '" & hETimeBegin & "' And hEdate <= '" & hETimeEnd & "' "
	End If
	If hETimeBegin <> "" and  hETimeEnd = "" Then
	    sql = sql & " And hEdate = '" & hETimeBegin & "' "
	End If
	else
	If hETimeBegin <> "" and hETimeEnd <> "" Then
	    sql = sql & " And hEdate >= #" & hETimeBegin & "# And hEdate <= #" & hETimeEnd & "# "
	End If
	If hETimeBegin <> "" and  hETimeEnd = "" Then
	    sql = sql & " And hEdate = #" & hETimeBegin & "# "
	End If
	end if
	
	If Session("CRM_level") < 9 Then
        sql = sql & " And hUser In (" & arrUser & ")"
	End If

If hType = "" And hState = "" And hMoney = "" And hUser = "" And hTimeBegin = "" And hTimeEnd = "" And hETimeBegin = "" And hETimeEnd = "" Then
    If Session("CRM_sql_Export_Hetong") <> "" Then
        sql = Session("CRM_sql_Export_Hetong")
	Else
	    If Session("CRM_level") < 9 Then
		    sql = " and hUser In ( " & arrUser & " ) "
		End If
	End If
Else
    Session("CRM_sql_Export_Hetong") = sql
End If

userfolder = Session("CRM_account") '生成文件夹
filefolder = Server.MapPath("../soft/"&userfolder)
set fso = Server.CreateObject("Scripting.FileSystemObject")
if not fso.FolderExists(filefolder) then '如果文件夹不存在则创建
fso.CreateFolder(filefolder) 
end if

	Exportitem = ""
	for i = 1 to 19
		if Request("Exportitem" & i) = "1" then
			Exportitem = Exportitem & "1"
		else
			Exportitem = Exportitem & "0"
		end if
	next
	
    strLine = strLine & "hId as 编号"
	strLine = strLine & ", cId as 公司名称"
	strLine = strLine & ", oId as 订单编号"
	if mid(Exportitem, 4, 1) = "1" then
	strLine = strLine & ", hNum as 合同编号"
	end if
	if mid(Exportitem, 5, 1) = "1" then
	strLine = strLine & ", hSdate as 起始时间"
	end if
	if mid(Exportitem, 6, 1) = "1" then
	strLine = strLine & ", hEdate as 到期时间"
	end if
	if mid(Exportitem, 7, 1) = "1" then
	strLine = strLine & ", hType as 合同分类"
	end if
	if mid(Exportitem, 8, 1) = "1" then
	strLine = strLine & ", hRevenue as 预付款"
	end if
	if mid(Exportitem, 9, 1) = "1" then
	strLine = strLine & ", hOwed as 欠款"
	end if
	if mid(Exportitem, 10, 1) = "1" then
	strLine = strLine & ", hMoney as 总金额"
	end if
	if mid(Exportitem, 11, 1) = "1" then
	strLine = strLine & ", hInvoice as 发票"
	end if
	if mid(Exportitem, 12, 1) = "1" then
	strLine = strLine & ", hTax as 含税"
	end if
	if mid(Exportitem, 13, 1) = "1" then
	strLine = strLine & ", hState as 合同状态"
	end if
	if mid(Exportitem, 14, 1) = "1" then
	strLine = strLine & ", hContent as 详情备注"
	end if
	if mid(Exportitem, 15, 1) = "1" then
	strLine = strLine & ", hAudit as 审核员"
	end if
	if mid(Exportitem, 16, 1) = "1" then
	strLine = strLine & ", hAuditTime as 审核时间"
	end if
	if mid(Exportitem, 17, 1) = "1" then
	strLine = strLine & ", hAuditReasons as 审核原因"
	end if
	if mid(Exportitem, 18, 1) = "1" then
	strLine = strLine & ", hUser as 业务员"
	end if
	if mid(Exportitem, 19, 1) = "1" then
	strLine = strLine & ", hTime as 录入时间"
	end if

on error resume next'如果有错误继续执行下面的代码 
dim excelfile,tbname
Dim ExcelDriver,DBExcelPath
tbname="Hetong"
Server.ScriptTimeOut=360000'防止超时
set rs=server.createobject("adodb.recordset")     
sql="select "&strLine&" from ["&tbname&"] where 1 = 1 "&sql&" Order By hId desc "'根据此SQL语句导出至Excel
rs.Open sql,conn,3,3
for Createtablei=0 to rs.Fields.Count-1                 '这里是为了创建sheet1用的
Createtable=Createtable&rs.fields(Createtablei).name&" text ,"
next

Createtablesql="Create table Sheet1("&left(Createtable,len(Createtable)-1)&")"
ExcelFile="../soft/"&userfolder&"/"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls" 
	'同步写入文件柜
	conn.execute "insert into OA_soft(s_class,s_title,s_file,s_user,s_share,s_time) values('"&L_Export_soft&"','"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls','"&ExcelFile&"','"&Session("CRM_name")&"','0','"&now()&"')"
set fso=Server.CreateObject ("Scripting.FileSystemObject")
fpath=Server.MapPath(ExcelFile)  
if fso.FileExists(fpath) then
whichfile=Server.MapPath(ExcelFile)
Set fs = CreateObject("Scripting.FileSystemObject")
Set thisfile = fs.GetFile(whichfile)
thisfile.delete true
end if             
Set connXLS = Server.CreateObject("ADODB.Connection")
ExcelDriver = "Driver={Microsoft Excel Driver (*.xls)};Readonly=0;"
DBExcelPath = "DBQ=" & Server.MapPath(excelfile) 
connXLS.Open ExcelDriver & DBExcelPath
connXLS.Execute(Createtablesql)'在这个conn上执行就得到一个excel
	
for ii=0 to rs.recordcount-1
for i=0 to rs.Fields.Count-1
	'cCompany = EasyCrm.getNewItem("Client","cID","'"&Rs(i)&"'","cCompany")
   Inserttablename=Inserttablename&rs.fields(i).name&","
   if i=1 then
   Inserttable=Inserttable&"'"&EasyCrm.getNewItem("Client","cID",""&Rs(i)&"","cCompany")&"',"
   elseif i=2 then
   Inserttable=Inserttable&"'"&EasyCrm.getNewItem("Order","oid",""&Rs(i)&"","oCode")&"',"
   else
   Inserttable=Inserttable&"'"&Rs(i)&"',"
   end if
Next 
Insertintosql="Insert into Sheet1("&left(Inserttablename,len(Inserttablename)-1)&")values("&left(Inserttable,len(Inserttable)-1)&")"
connXLS.Execute(Insertintosql)
Insertintosql =""
Inserttable=""
Inserttablename=""
rs.MoveNext
Next

Session("CRM_sql_Export_Hetong") = ""

if ""&YNalert&"" = 1 then
	Response.Write ("<script>alert("""&L_Export_alert&""");</script>")
end if
	Response.Write ("<script>location.href='?action=Hetong&otype=Hetong' ;</script>")
end Sub

Sub Service()%>
	<tr> 
		<td valign="top" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="6"><B><%=L_Top_Search%></B></td>
							</tr>
						</table>
						<form name="searchForm" action="?Action=ServicetoExcel" method="post">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" /><col width="220" /><col width="100" />
							<tr>
								<td class="td_l_c title" style="border-top:0;"><%=L_Service_sType%></td>
								<td class="td_r_l" style="border-top:0;"> <% = EasyCrm.getSelect("SelectData","Select_Service","sType","") %></td>
								<td class="td_l_c title" style="border-top:0;"><%=L_Service_sSDate%></td>
								<td class="td_r_l" style="border-top:0;"> <input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" /></td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Service_sUser%></td>
								<td class="td_r_l"> <% if Session("CRM_level") = 9 then %><% = EasyCrm.UserList(2,"User","") %><%else%><% = EasyCrm.UserList(1,"User","") %><%end if%></td>
								<td class="td_l_c title"><%=L_Service_sEDate%></td>
								<td class="td_r_l"> <input name="ETimeBegin" type="text" id="ETimeBegin" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" />&nbsp;~&nbsp;<input name="ETimeEnd" type="text" id="ETimeEnd" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" /></td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Export_content%></td>
								<td class="td_r_l" colspan=3>
									<input type="checkbox" name="Exportitem1" value="1" checked> <%=L_Service_sid%>　
									<input type="checkbox" name="Exportitem2" value="1" checked> <%=L_Service_cId%>　
									<input type="checkbox" name="Exportitem3" value="1" checked> <%=L_Service_ProId%>　
									<input type="checkbox" name="Exportitem4" value="1" checked> <%=L_Service_sTitle%>　
									<input type="checkbox" name="Exportitem5" value="1" checked> <%=L_Service_sLinkman%>　
									<input type="checkbox" name="Exportitem6" value="1" checked> <%=L_Service_sType%>　
									<input type="checkbox" name="Exportitem7" value="1" checked> <%=L_Service_sSDate%>　
									<input type="checkbox" name="Exportitem8" value="1" checked> <%=L_Service_sEDate%>　<br>
									<input type="checkbox" name="Exportitem9" value="1" checked> <%=L_Service_sContent%>　
									<input type="checkbox" name="Exportitem10" value="1" checked> <%=L_Service_sSolve%>　
									<input type="checkbox" name="Exportitem11" value="1" checked> <%=L_Service_sInfo%>　
									<input type="checkbox" name="Exportitem12" value="1" checked> <%=L_Service_sUser%>　
									<input type="checkbox" name="Exportitem13" value="1" checked> <%=L_Service_sTime%>　
								</td>
							</tr>
						</table>  
						<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:10px;">
							<tr>
								<td>
									<input type="submit" name="Submit" class="button45" value=" <%=L_Export%> ">　
									<input type="button" class="button42" value=" 下载文件 " onClick=window.location.href="../Soft/index.asp" />　
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?action=killSession&otype=Records" />　
								</td>
							</tr>
						</table>  
						</form> 
					</td> 
				</tr>
			</table>  
		</td> 
	</tr>
</table>
</body>
</html>
<script src="../data/calendar/WdatePicker.js"></script>
<%
end Sub

sub ServicetoExcel()

	Dim sType,sUser,arrUser,sTimeBegin,sTimeEnd,sETimeBegin,sETimeEnd
	sType = Request("sType")
	sUser = Request("User")
	sTimeBegin = Request("TimeBegin")
	sTimeEnd = Request("TimeEnd")
	sETimeBegin = Request("ETimeBegin")
	sETimeEnd = Request("ETimeEnd")
	
	arrUser = getUserList(Session("CRM_level"),Session("CRM_group"),Session("CRM_MR"))
	
	Dim sql
    sql = ""
	
    If sType <> "" Then
		sql = sql & " And sType = '" & sType & "'"
	End If
		
    If sUser <> "" Then
		sql = sql & " and sUser = '" & sUser & "'"
	End If

	if Accsql =1 then
	If hTimeBegin <> "" and  hTimeEnd <> "" Then
	    sql = sql & " And hSdate >= '" & hTimeBegin & "' And hSdate <= '" & hTimeEnd & "' "
	End If
	If hTimeBegin <> "" and  hTimeEnd = "" Then
	    sql = sql & " And hSdate = '" & hTimeBegin & "' "
	End If
	else
	If hTimeBegin <> "" and hTimeEnd <> "" Then
	    sql = sql & " And hSdate >= #" & hTimeBegin & "# And hSdate <= #" & hTimeEnd & "# "
	End If
	If hTimeBegin <> "" and  hTimeEnd = "" Then
	    sql = sql & " And hSdate = #" & hTimeBegin & "# "
	End If
	end if

	if Accsql =1 then
	If hETimeBegin <> "" and  hETimeEnd <> "" Then
	    sql = sql & " And hEdate >= '" & hETimeBegin & "' And hEdate <= '" & hETimeEnd & "' "
	End If
	If hETimeBegin <> "" and  hETimeEnd = "" Then
	    sql = sql & " And hEdate = '" & hETimeBegin & "' "
	End If
	else
	If hETimeBegin <> "" and hETimeEnd <> "" Then
	    sql = sql & " And hEdate >= #" & hETimeBegin & "# And hEdate <= #" & hETimeEnd & "# "
	End If
	If hETimeBegin <> "" and  hETimeEnd = "" Then
	    sql = sql & " And hEdate = #" & hETimeBegin & "# "
	End If
	end if
	
	If Session("CRM_level") < 9 Then
        sql = sql & " And sUser In (" & arrUser & ")"
	End If

If sType = "" And hState = "" And hMoney = "" And sUser = "" And hTimeBegin = "" And hTimeEnd = "" And hETimeBegin = "" And hETimeEnd = "" Then
    If Session("CRM_sql_Export_Service") <> "" Then
        sql = Session("CRM_sql_Export_Service")
	Else
	    If Session("CRM_level") < 9 Then
		    sql = " and sUser In ( " & arrUser & " ) "
		End If
	End If
Else
    Session("CRM_sql_Export_Service") = sql
End If

userfolder = Session("CRM_account") '生成文件夹
filefolder = Server.MapPath("../soft/"&userfolder)
set fso = Server.CreateObject("Scripting.FileSystemObject")
if not fso.FolderExists(filefolder) then '如果文件夹不存在则创建
fso.CreateFolder(filefolder) 
end if

	Exportitem = ""
	for i = 1 to 13
		if Request("Exportitem" & i) = "1" then
			Exportitem = Exportitem & "1"
		else
			Exportitem = Exportitem & "0"
		end if
	next

    strLine = strLine & "sid as 编号"
	strLine = strLine & ", cId as 公司名称"
	strLine = strLine & ", ProId as 相关产品"
	if mid(Exportitem, 4, 1) = "1" then
	strLine = strLine & ", sTitle as 反馈主题"
	end if
	if mid(Exportitem, 5, 1) = "1" then
	strLine = strLine & ", sLinkman as 联系人"
	end if
	if mid(Exportitem, 6, 1) = "1" then
	strLine = strLine & ", sType as 反馈分类"
	end if
	if mid(Exportitem, 7, 1) = "1" then
	strLine = strLine & ", sSDate as 反馈日期"
	end if
	if mid(Exportitem, 8, 1) = "1" then
	strLine = strLine & ", sEDate as 结束日期"
	end if
	if mid(Exportitem, 9, 1) = "1" then
	strLine = strLine & ", sContent as 详情备注"
	end if
	if mid(Exportitem, 10, 1) = "1" then
	strLine = strLine & ", sSolve as 是否解决"
	end if
	if mid(Exportitem, 11, 1) = "1" then
	strLine = strLine & ", sInfo as 处理结果"
	end if
	if mid(Exportitem, 12, 1) = "1" then
	strLine = strLine & ", sUser as 业务员"
	end if
	if mid(Exportitem, 13, 1) = "1" then
	strLine = strLine & ", sTime as 录入时间"
	end if

on error resume next'如果有错误继续执行下面的代码 
dim excelfile,tbname
Dim ExcelDriver,DBExcelPath
tbname="Service"
Server.ScriptTimeOut=360000'防止超时
set rs=server.createobject("adodb.recordset")     
sql="select "&strLine&" from ["&tbname&"] where 1 = 1 "&sql&" Order By sId desc "'根据此SQL语句导出至Excel
rs.Open sql,conn,3,3
for Createtablei=0 to rs.Fields.Count-1                 '这里是为了创建sheet1用的
Createtable=Createtable&rs.fields(Createtablei).name&" text ,"
next

Createtablesql="Create table Sheet1("&left(Createtable,len(Createtable)-1)&")"
ExcelFile="../soft/"&userfolder&"/"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls" 
	'同步写入文件柜
	conn.execute "insert into OA_soft(s_class,s_title,s_file,s_user,s_share,s_time) values('"&L_Export_soft&"','"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls','"&ExcelFile&"','"&Session("CRM_name")&"','0','"&now()&"')"
set fso=Server.CreateObject ("Scripting.FileSystemObject")
fpath=Server.MapPath(ExcelFile)  
if fso.FileExists(fpath) then
whichfile=Server.MapPath(ExcelFile)
Set fs = CreateObject("Scripting.FileSystemObject")
Set thisfile = fs.GetFile(whichfile)
thisfile.delete true
end if             
Set connXLS = Server.CreateObject("ADODB.Connection")
ExcelDriver = "Driver={Microsoft Excel Driver (*.xls)};Readonly=0;"
DBExcelPath = "DBQ=" & Server.MapPath(excelfile) 
connXLS.Open ExcelDriver & DBExcelPath
connXLS.Execute(Createtablesql)'在这个conn上执行就得到一个excel
	
for ii=0 to rs.recordcount-1
for i=0 to rs.Fields.Count-1
	'cCompany = EasyCrm.getNewItem("Client","cID","'"&Rs(i)&"'","cCompany")
   Inserttablename=Inserttablename&rs.fields(i).name&","
   if i=1 then
   Inserttable=Inserttable&"'"&EasyCrm.getNewItem("Client","cID",""&Rs(i)&"","cCompany")&"',"
   elseif i=2 then
   Inserttable=Inserttable&"'"&EasyCrm.getNewItem("Products","ID",""&Rs(i)&"","pTitle")&"',"
   else
   Inserttable=Inserttable&"'"&Rs(i)&"',"
   end if
Next 
Insertintosql="Insert into Sheet1("&left(Inserttablename,len(Inserttablename)-1)&")values("&left(Inserttable,len(Inserttable)-1)&")"
connXLS.Execute(Insertintosql)
Insertintosql =""
Inserttable=""
Inserttablename=""
rs.MoveNext
Next

Session("CRM_sql_Export_Service") = ""

if ""&YNalert&"" = 1 then
	Response.Write ("<script>alert("""&L_Export_alert&""");</script>")
end if
	Response.Write ("<script>location.href='?action=Service&otype=Service' ;</script>")
end Sub

Sub Expense()%>

	<script>
	function Show()
	{
		if (document.getElementById('eOutIn').value=="1") 
		 {
			document.getElementById("eTypeA").style.display = "block";
			document.getElementById("eTypeB").style.display = "none";
		 }
		else if (document.getElementById('eOutIn').value=="0") 
		 {
			document.getElementById("eTypeA").style.display = "none";
			document.getElementById("eTypeB").style.display = "block";
		 }
		 else if (document.getElementById('eOutIn').value=="") 
		 {
			document.getElementById("eTypeA").style.display = "none";
			document.getElementById("eTypeB").style.display = "none";
		 }
	}
	</script>
	<tr> 
		<td valign="top" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="6"><B><%=L_Top_Search%></B></td>
							</tr>
						</table>
						<form name="searchForm" action="?Action=ExpensetoExcel" method="post">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" /><col width="120" /><col width="100" /><col width="250" /><col width="100" /><col width="80" />
							<tr>
								<td class="td_l_c title" style="border-top:0;"><%=L_Expense_eUser%></td>
								<td class="td_r_l" style="border-top:0;"> <% if Session("CRM_level") = 9 then %><% = EasyCrm.UserList(2,"User","") %><%else%><% = EasyCrm.UserList(1,"User","") %><%end if%></td>
								<td class="td_l_c title" style="border-top:0;"><%=L_Expense_eDate%></td>
								<td class="td_r_l" style="border-top:0;"> <input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" style="width:100px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" /></td>
								<td class="td_l_c title" style="border-top:0;"><%=L_Expense_eOutIn%></td>
								<td class="td_r_l" style="border-top:0;border-right:0;"><select name='eOutIn' onchange="Show();"><option value="">请选择</option><option value="0">支出</option><option value="1">收入</option></select> 
								</td>
								<td class="td_r_l" style="border-top:0;">
								<span id=eTypeA STYLE="display:none;" ><% = EasyCrm.getSelect("SelectData","Select_ExpenseIN","eTypeA","") %></span>
								<span id=eTypeB STYLE="display:none;"><% = EasyCrm.getSelect("SelectData","Select_ExpenseOUT","eTypeB","") %></span>
								</td>
							</tr>
							<tr>
								<td class="td_l_c title"><%=L_Export_content%></td>
								<td class="td_r_l" colspan=6>
									<input type="checkbox" name="Exportitem1" value="1" checked> <%=L_Expense_eid%>　
									<input type="checkbox" name="Exportitem2" value="1" checked> <%=L_Expense_cId%>　
									<input type="checkbox" name="Exportitem3" value="1" checked> <%=L_Expense_eDate%>　
									<input type="checkbox" name="Exportitem4" value="1" checked> <%=L_Expense_eOutIn%>　
									<input type="checkbox" name="Exportitem5" value="1" checked> <%=L_Expense_eType%>　
									<input type="checkbox" name="Exportitem6" value="1" checked> <%=L_Expense_eMoney%>　
									<input type="checkbox" name="Exportitem7" value="1" checked> <%=L_Expense_eContent%>　
									<input type="checkbox" name="Exportitem8" value="1" checked> <%=L_Expense_eUser%>　
									<input type="checkbox" name="Exportitem9" value="1" checked> <%=L_Expense_eTime%>　
								</td>
							</tr>
						</table>  
						<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:10px;">
							<tr>
								<td>
									<input type="submit" name="Submit" class="button45" value=" <%=L_Export%> ">　
									<input type="button" class="button42" value=" 下载文件 " onClick=window.location.href="../Soft/index.asp" />　
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?action=killSession&otype=Records" />　
								</td>
							</tr>
						</table>  
						</form> 
					</td> 
				</tr>
			</table>  
		</td> 
	</tr>
</table>
</body>
</html>
<script src="../data/calendar/WdatePicker.js"></script>
<%
end Sub

sub ExpensetoExcel()

	Dim eType,eOutIn,eUser,arrUser,eTimeBegin,eTimeEnd
	eOutIn = Request("eOutIn")
	eTypeA = Request("eTypeA")
	eTypeB = Request("eTypeB")
	IF eTypeA<>"" THEN
	eType = eTypeA
	ELSE
	eType = eTypeB
	END IF
	eUser = Request("User")
	eTimeBegin = Request("TimeBegin")
	eTimeEnd = Request("TimeEnd")
	
	arrUser = getUserList(Session("CRM_level"),Session("CRM_group"),Session("CRM_MR"))
	
	Dim sql
    sql = ""
	
    If eOutIn <> "" Then
		sql = sql & " And eOutIn = '" & eOutIn & "'"
	End If
	
    If eType <> "" Then
		sql = sql & " And eType = '" & eType & "'"
	End If
		
    If eUser <> "" Then
		sql = sql & " and eUser = '" & eUser & "'"
	End If

	if Accsql =1 then
	If eTimeBegin <> "" and  eTimeEnd <> "" Then
	    sql = sql & " And eTime >= '" & eTimeBegin & "' And hSdate <= '" & eTimeEnd & "' "
	End If
	If eTimeBegin <> "" and  eTimeEnd = "" Then
	    sql = sql & " And hSdate = '" & eTimeBegin & "' "
	End If
	else
	If eTimeBegin <> "" and eTimeEnd <> "" Then
	    sql = sql & " And hSdate >= #" & eTimeBegin & "# And hSdate <= #" & eTimeEnd & "# "
	End If
	If eTimeBegin <> "" and  eTimeEnd = "" Then
	    sql = sql & " And hSdate = #" & eTimeBegin & "# "
	End If
	end if
	
	If Session("CRM_level") < 9 Then
        sql = sql & " And eUser In (" & arrUser & ")"
	End If

If eOutIn = "" And eType = "" And eUser = "" And eTimeBegin = "" And eTimeEnd = "" Then
    If Session("CRM_sql_Export_Expense") <> "" Then
        sql = Session("CRM_sql_Export_Expense")
	Else
	    If Session("CRM_level") < 9 Then
		    sql = " and eUser In ( " & arrUser & " ) "
		End If
	End If
Else
    Session("CRM_sql_Export_Expense") = sql
End If

userfolder = Session("CRM_account") '生成文件夹
filefolder = Server.MapPath("../soft/"&userfolder)
set fso = Server.CreateObject("Scripting.FileSystemObject")
if not fso.FolderExists(filefolder) then '如果文件夹不存在则创建
fso.CreateFolder(filefolder) 
end if

	Exportitem = ""
	for i = 1 to 9
		if Request("Exportitem" & i) = "1" then
			Exportitem = Exportitem & "1"
		else
			Exportitem = Exportitem & "0"
		end if
	next

    strLine = strLine & "eid as 编号"
	strLine = strLine & ", cId as 公司名称"
	strLine = strLine & ", eOutIn as 收支类型"
	if mid(Exportitem, 3, 1) = "1" then
	strLine = strLine & ", eDate as 收支日期"
	end if
	if mid(Exportitem, 5, 1) = "1" then
	strLine = strLine & ", eType as 费用类型"
	end if
	if mid(Exportitem, 6, 1) = "1" then
	strLine = strLine & ", eMoney as 金额"
	end if
	if mid(Exportitem, 7, 1) = "1" then
	strLine = strLine & ", eContent as 详情备注"
	end if
	if mid(Exportitem, 8, 1) = "1" then
	strLine = strLine & ", eUser as 业务员"
	end if
	if mid(Exportitem, 9, 1) = "1" then
	strLine = strLine & ", eTime as 录入时间"
	end if

on error resume next'如果有错误继续执行下面的代码 
dim excelfile,tbname
Dim ExcelDriver,DBExcelPath
tbname="Expense"
Server.ScriptTimeOut=360000'防止超时
set rs=server.createobject("adodb.recordset")     
sql="select "&strLine&" from ["&tbname&"] where 1 = 1 "&sql&" Order By eId desc "'根据此SQL语句导出至Excel
rs.Open sql,conn,3,3
for Createtablei=0 to rs.Fields.Count-1                 '这里是为了创建sheet1用的
Createtable=Createtable&rs.fields(Createtablei).name&" text ,"
next

Createtablesql="Create table Sheet1("&left(Createtable,len(Createtable)-1)&")"
ExcelFile="../soft/"&userfolder&"/"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls" 
	'同步写入文件柜
	conn.execute "insert into OA_soft(s_class,s_title,s_file,s_user,s_share,s_time) values('"&L_Export_soft&"','"&tbname&"-"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls','"&ExcelFile&"','"&Session("CRM_name")&"','0','"&now()&"')"
set fso=Server.CreateObject ("Scripting.FileSystemObject")
fpath=Server.MapPath(ExcelFile)  
if fso.FileExists(fpath) then
whichfile=Server.MapPath(ExcelFile)
Set fs = CreateObject("Scripting.FileSystemObject")
Set thisfile = fs.GetFile(whichfile)
thisfile.delete true
end if             
Set connXLS = Server.CreateObject("ADODB.Connection")
ExcelDriver = "Driver={Microsoft Excel Driver (*.xls)};Readonly=0;"
DBExcelPath = "DBQ=" & Server.MapPath(excelfile) 
connXLS.Open ExcelDriver & DBExcelPath
connXLS.Execute(Createtablesql)'在这个conn上执行就得到一个excel
	
for ii=0 to rs.recordcount-1
for i=0 to rs.Fields.Count-1
	'cCompany = EasyCrm.getNewItem("Client","cID","'"&Rs(i)&"'","cCompany")
   Inserttablename=Inserttablename&rs.fields(i).name&","
   if i=1 then
   Inserttable=Inserttable&"'"&EasyCrm.getNewItem("Client","cID",""&Rs(i)&"","cCompany")&"',"
   elseif i=2 then
		if Rs(i)=0 then
			Inserttable=Inserttable&"'支出',"
		elseif Rs(i)=1 then
			Inserttable=Inserttable&"'收入',"
		end if
   else
   Inserttable=Inserttable&"'"&Rs(i)&"',"
   end if
Next 
Insertintosql="Insert into Sheet1("&left(Inserttablename,len(Inserttablename)-1)&")values("&left(Inserttable,len(Inserttable)-1)&")"
connXLS.Execute(Insertintosql)
Insertintosql =""
Inserttable=""
Inserttablename=""
rs.MoveNext
Next

Session("CRM_sql_Export_Expense") = ""

if ""&YNalert&"" = 1 then
	Response.Write ("<script>alert("""&L_Export_alert&""");</script>")
end if
	Response.Write ("<script>location.href='?action=Expense&otype=Expense' ;</script>")
end Sub

sub killSession()
	Session("CRM_sql_Export_Client") = ""
	Session("CRM_sql_Export_Records") = ""
	Session("CRM_sql_Export_Order") = ""
	Session("CRM_sql_Export_Hetong") = ""
	Session("CRM_sql_Export_Service") = ""
	Session("CRM_sql_Export_Expense") = ""
	Response.Write ("<script>location.href='?action="&otype&"&otype="&otype&"' ;</script>")
end Sub
%><%else%>无权限<%end if%><% Set EasyCrm = nothing %>