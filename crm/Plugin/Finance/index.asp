<!--#include file="../../data/conn.asp" --><!--#include file="config.asp" --><!--#include file="../../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
'��ȡ��ǰҳ��
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
otype	=	Request.QueryString("otype")
if otype="" then otype="Main"

subAction = Trim(Request("subAction"))
If subAction = "searchItem" Then
	fUser = EasyCrm.Searchcode(Request("fUser"))
	fSubjects = EasyCrm.Searchcode(Request("fSubjects"))
	fClass = EasyCrm.Searchcode(Request("fClass"))
	TimeBegin = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	TimeEnd = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	fType = EasyCrm.Searchcode(Request.Form("fType"))
	
	Session("Search_Plugin_Finance_fUser") = EasyCrm.Searchcode(Request("fUser"))
	Session("Search_Plugin_Finance_fSubjects") = EasyCrm.Searchcode(Request("fSubjects"))
	Session("Search_Plugin_Finance_fClass") = EasyCrm.Searchcode(Request("fClass"))
	Session("Search_Plugin_Finance_TimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Plugin_Finance_TimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	Session("Search_Plugin_Finance_fType") = EasyCrm.Searchcode(Request.Form("fType"))
		
    If fUser <> "" Then
        sql = sql & " And fUser = '" & fUser & "' "
	End If
    If fSubjects <> "" Then
        sql = sql & " And fSubjects = '" & fSubjects & "' "
	End If
    If fClass <> "" Then
        sql = sql & " And fClass = '" & fClass & "' "
	End If
    If fType = "fDebit" Then
        sql = sql & " And fDebit > 0 "
	elseif fType = "fCredit" then
        sql = sql & " And fCredit > 0 "
	End If
	
	if Accsql =1 then
	If TimeBegin <> "" and  TimeEnd <> "" Then
	    sql = sql & " And fTime >= '" & TimeBegin & "' And fTime <= '" & TimeEnd & "' "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And fTime = '" & TimeBegin & "' "
	End If
	else
	If TimeBegin <> "" and TimeEnd <> "" Then
	    sql = sql & " And fTime >= #" & TimeBegin & "# And fTime <= #" & TimeEnd & "# "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sql = sql & " And fTime = #" & TimeBegin & "# "
	End If
	end if
	
End If

	If fUser = "" And fSubjects = ""  And fClass = "" And TimeBegin = "" And TimeBegin = "" Then
		If Session("Search_Plugin_Finance_Search") <> "" Then
			sql = Session("Search_Plugin_Finance_Search")
		End If
	Else
		Session("Search_Plugin_Finance_Search") = sql
	End If

	If subAction = "killSession" Then
		Session("Search_Plugin_Finance_Search") = ""
		Session("Search_Plugin_Finance_fUser") = ""
		Session("Search_Plugin_Finance_fSubjects") = ""
		Session("Search_Plugin_Finance_fClass") = ""
		Session("Search_Plugin_Finance_TimeBegin") = ""
		Session("Search_Plugin_Finance_TimeEnd") = ""
		Response.Write("<script>location.href='?' ;</script>")
	End If
	
If subAction = "searchBank" Then
	bName = EasyCrm.Searchcode(Request("bName"))
	bClass = EasyCrm.Searchcode(Request("bClass"))
	TimeBegin = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	TimeEnd = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	bType = EasyCrm.Searchcode(Request.Form("bType"))
	
	Session("Search_Plugin_Finance_Bank_Name") = EasyCrm.Searchcode(Request("bName"))
	Session("Search_Plugin_Finance_Bank_Class") = EasyCrm.Searchcode(Request("bClass"))
	Session("Search_Plugin_Finance_Bank_TimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Plugin_Finance_Bank_TimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	Session("Search_Plugin_Finance_Bank_Type") = EasyCrm.Searchcode(Request.Form("bType"))
		
    If bName <> "" Then
        sqlb = sqlb & " And bName = '" & bName & "' "
	End If
    If bClass <> "" Then
        sqlb = sqlb & " And bClass = '" & bClass & "' "
	End If
    If bType = "bDebit" Then
        sqlb = sqlb & " And bDebit > 0 "
	elseif bType = "bCredit" then
        sqlb = sqlb & " And bCredit > 0 "
	End If
	
	if Accsql =1 then
	If TimeBegin <> "" and  TimeEnd <> "" Then
	    sqlb = sqlb & " And bTime >= '" & TimeBegin & "' And bTime <= '" & TimeEnd & "' "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sqlb = sqlb & " And bTime = '" & TimeBegin & "' "
	End If
	else
	If TimeBegin <> "" and TimeEnd <> "" Then
	    sqlb = sqlb & " And bTime >= #" & TimeBegin & "# And bTime <= #" & TimeEnd & "# "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sqlb = sqlb & " And bTime = #" & TimeBegin & "# "
	End If
	end if
	
End If

	If bName = "" And bClass = ""  And bType = "" And TimeBegin = "" And TimeBegin = "" Then
		If Session("Search_Plugin_Finance_Bank_Search") <> "" Then
			sqlb = Session("Search_Plugin_Finance_Bank_Search")
		End If
	Else
		Session("Search_Plugin_Finance_Bank_Search") = sqlb
	End If

	If subAction = "killBankSession" Then
		Session("Search_Plugin_Finance_Bank_Search") = ""
		Session("Search_Plugin_Finance_Bank_Name") = ""
		Session("Search_Plugin_Finance_Bank_Class") = ""
		Session("Search_Plugin_Finance_Bank_TimeBegin") = ""
		Session("Search_Plugin_Finance_Bank_TimeEnd") = ""
		Session("Search_Plugin_Finance_Bank_Type") = ""
		Response.Write("<script>location.href='?otype=Bank' ;</script>")
	End If
	
If subAction = "searchOutin" Then
	oCompany = EasyCrm.Searchcode(Request("oCompany"))
	osType = EasyCrm.Searchcode(Request.Form("osType"))
	TimeBegin = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	TimeEnd = EasyCrm.Searchcode(Request.Form("TimeEnd"))
	
	Session("Search_Plugin_Finance_Outin_oCompany") = EasyCrm.Searchcode(Request("oCompany"))
	Session("Search_Plugin_Finance_Outin_osType") = EasyCrm.Searchcode(Request.Form("osType"))
	Session("Search_Plugin_Finance_Outin_TimeBegin") = EasyCrm.Searchcode(Request.Form("TimeBegin"))
	Session("Search_Plugin_Finance_Outin_TimeEnd") = EasyCrm.Searchcode(Request.Form("TimeEnd"))
		
    If oCompany <> "" Then
        sqlc = sqlc & " And oCompany like '%" & oCompany & "%' "
	End If
    If osType = "oDebit" Then
        sqlc = sqlc & " And oDebit > 0 "
	elseif osType = "oCredit" then
        sqlc = sqlc & " And oCredit > 0 "
	End If
	
	if Accsql =1 then
	If TimeBegin <> "" and  TimeEnd <> "" Then
	    sqlc = sqlc & " And oTime >= '" & TimeBegin & "' And oTime <= '" & TimeEnd & "' "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sqlc = sqlc & " And oTime = '" & TimeBegin & "' "
	End If
	else
	If TimeBegin <> "" and TimeEnd <> "" Then
	    sqlc = sqlc & " And oTime >= #" & TimeBegin & "# And oTime <= #" & TimeEnd & "# "
	End If
	If TimeBegin <> "" and  TimeEnd = "" Then
	    sqlc = sqlc & " And oTime = #" & TimeBegin & "# "
	End If
	end if
	
End If

	If oCompany = "" And osType = "" And TimeBegin = "" And TimeBegin = "" Then
		If Session("Search_Plugin_Finance_Outin_Search") <> "" Then
			sqlc = Session("Search_Plugin_Finance_Outin_Search")
		End If
	Else
		Session("Search_Plugin_Finance_Outin_Search") = sqlc
	End If

	If subAction = "killOutinSession" Then
		Session("Search_Plugin_Finance_Outin_Search") = ""
		Session("Search_Plugin_Finance_Outin_oCompany") = ""
		Session("Search_Plugin_Finance_Outin_osType") = ""
		Session("Search_Plugin_Finance_Outin_TimeBegin") = ""
		Session("Search_Plugin_Finance_Outin_TimeEnd") = ""
		Response.Write("<script>location.href='?otype=Outin' ;</script>")
	End If
%>
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
<style type="text/css">
.textarea {width:100%;height:20px;border:0;overflow:hidden}
img{max-height:500px;_height:expression_r(this.scrollHeight > 500 ? "500px" : "auto");}
</style>
</head>

<body>
<style>body {padding-top:35px;padding-bottom:55px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">��ǰλ�ã����ܲ�� > ������� </td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="ˢ��" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="����" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="ǰ��" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<%if inStr(Plugin_Finance_manage,session("CRM_name"))>0 or Session("CRM_level") = 9 then%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
            <div class="MenuboxS">
              <ul>
                <li <%if otype="Main" then%>class="hover"<%end if%>><span><a href="?action=List&otype=Main">��ˮ��</a></span></li>
                <li <%if otype="Bank" then%>class="hover"<%end if%>><span><a href="?action=List&otype=Bank">���д��</a></span></li>
                <li <%if otype="Outin" then%>class="hover"<%end if%>><span><a href="?action=List&otype=Outin">��֧����</a></span></li>
                <li <%if otype="Manage" then%>class="hover"<%end if%>><span><a href="?action=Manage&otype=Manage">�߼�����</a></span></li>
              </ul>
            </div>
		</td>
	</tr>
</table>
<%
action = Trim(Request("action"))
Select Case action
Case "List"
    Call List()
Case "Manage"
    Call infoManage()
Case "Managesave"
    Call infoManagesave()
Case "delete"
    Call infodelete()
Case "deleteBank"
    Call infodeleteBank()
Case "deleteOutin"
    Call infodeleteOutin()
Case Else
    Call List()
End Select

%>

<%
Sub List()
if otype="Main" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10">
		
			<span  style="float:left;padding:0 10px;height:34px;text-align:left;position:fixed;right:10px;top:45px;color:#000;">
				<input type="button" name="Add" class="button41" value="ɸѡ" onclick="Showhiden(this,'boxSearch',false,'ɸѡ','ɸѡ')" style="cursor:pointer"  />
				<input type="button" name="Add" class="button45" value="����" onclick='Plugin_Finance_InfoAdd()' style="cursor:pointer"  />
				<script>function Plugin_Finance_InfoAdd() {$.dialog.open('GetUpdate.asp?action=Add', {title: '����', width: 700,height: 420, fixed: true}); };</script>
			</span>
			<span  style="float:left;padding:0 10px;text-align:left;position:fixed;right:10px;top:80px;color:#000;background:#666;">
						<form name="searchForm" action="?subAction=searchItem" method="post">
						<table width="250" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" id="boxSearch" style="display:none;margin:10px 0;background-color:#fff;">
							<col width="90" /><col width="160" />
							<tr>
								<td class="td_l_r title">��ʼ����</td>
								<td class="td_r_l"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="20" onFocus="WdatePicker()" value="<%=Session("Search_Plugin_Finance_TimeBegin")%>" /></td>
							</tr>
							<tr>
								<td class="td_l_r title">��������</td>
								<td class="td_r_l"><input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="20" onFocus="WdatePicker()" value="<%=Session("Search_Plugin_Finance_TimeEnd")%>" /> </td>
							</tr>
							<tr>
								<td class="td_l_r title" >����</td>
								<td class="td_r_l"><% = EasyCrm.UserList(2,"fUser",Session("Search_Plugin_Finance_fUser")) %></td>
							</tr>
							<tr>
								<td class="td_l_r title" >�Է���Ŀ</td>
								<td class="td_r_l">
									<select name="fSubjects" class="int" style="width:130px;">
										<option value="">��ѡ��</option>
										<%
										str = split(""&Plugin_Finance_Subjects&"",",")
										for i = 0 to ubound(str)
										if Session("Search_Plugin_Finance_fSubjects") = str(i) then
										response.Write "<option value="&str(i)&" selected>"&str(i)&"</option>"
										else
										response.Write "<option value="&str(i)&">"&str(i)&"</option>"
										end if
										next
										%>
									</select>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title" >����</td>
								<td class="td_r_l">
									<select name="fClass" class="int" style="width:130px;">
										<option value="">��ѡ��</option>
										<%
										str = split(""&Plugin_Finance_Class&"",",")
										for i = 0 to ubound(str)
										if Session("Search_Plugin_Finance_fClass") = str(i) then
										response.Write "<option value="&str(i)&" selected>"&str(i)&"</option>"
										else
										response.Write "<option value="&str(i)&">"&str(i)&"</option>"
										end if
										next
										%>
									</select>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title" >���</td>
								<td class="td_r_l">
									<input name="fType" type="radio" class="noborder" value="fDebit"> ��+��
									<input name="fType" type="radio" class="noborder" value="fCredit"> ��-��
								</td>
							</tr>
							<tr>
								<td class="td_l_c" colspan="2" style="padding:5px 0;">
									<input type="submit" name="Submit" class="button42" value=" <%=L_Search%> ">��
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?SubAction=killSession" /></td>
								</tr>
						</table>   
						</form>
			</span>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td width="80" class="td_l_c">���</td>
					<td width="100" class="td_l_c">����</td>
					<td width="100" class="td_l_c">�Է���Ŀ</td>
					<td width="100" class="td_l_c">����</td>
					<td width="100" class="td_l_c">��Ŀ</td>
					<td width="120" class="td_l_c">ժҪ</td>
					<td width="100" class="td_l_c">��(+)</td>
					<td width="100" class="td_l_c">��(-)</td>
					<td class="td_l_l">��ע</td>
					<td width="80" class="td_l_c">����</td>
					<td width="80" class="td_l_c">���</td>
					<%if Session("CRM_level") = 9 then%>
					<td width="90" class="td_l_c">����</td>
					<%end if%>
				</tr>
	<%
	PN = CLng(ABS(Request("PN")))
	If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
	intPageSize = Plugin_Finance_Page
	pagenum = intPageSize*(PN-1)
	Set rs = Server.CreateObject("ADODB.Recordset")
	IF PN=1 THEN
	rs.Open "Select top "&intPageSize&" * From [Plugin_Finance] where 1=1 "&sql&" Order By Id desc ",conn,1,1 
	ELSE
	rs.Open "Select top "&intPageSize&" * From [Plugin_Finance] where 1=1 "&sql&" and Id < ( SELECT Min(Id) FROM ( SELECT TOP "&pagenum&" Id FROM [Plugin_Finance] where 1=1 "&sql&" ORDER BY Id desc ) AS T ) Order By Id desc ",conn,1,1
	END IF
	SQLstr="Select count(Id) As RecordSum From [Plugin_Finance] where 1=1 "&sql&""
	Set Rsstr=conn.Execute(SQLstr,1,1) 
	TotalRecords=Rsstr("RecordSum") 
	if Int(TotalRecords/intPageSize)=TotalRecords/intPageSize then
	TotalPages=TotalRecords/intPageSize
	else
	TotalPages=Int(TotalRecords/intPageSize)+1
	end if
	Rsstr.Close 
	Set Rsstr=Nothing
	If PN > TotalPages Then PN = TotalPages
	i=0
	Do While Not rs.BOF And Not rs.EOF
	i=i+1
	%>
				<tr class="tr">
					<td class="td_l_c"><%=rs("ID") %></td>
					<td class="td_l_c"><%=EasyCrm.FormatDate(rs("fTime"),2)%></td>
					<td class="td_l_c"><%=rs("fSubjects") %></td>
					<td class="td_l_c"><%=rs("fClass") %></td>
					<td class="td_l_c"><%=rs("fProject") %></td>
					<td class="td_l_c"><%=rs("fDigest") %></td>
					<td class="td_l_c"><%=rs("fDebit") %></td>
					<td class="td_l_c"><%=rs("fCredit") %></td>
					<td class="td_l_l"><input type="button" class="button226" value="Ԥ��" onclick='ContentView<%=rs("Id")%>()' style="cursor:pointer" /></td>
					<td class="td_l_c"><%=rs("fUser") %></td>
					<td class="td_l_c"><% if Session("CRM_level")=9 then%><a onclick='InfoAudit<%=rs("Id")%>()' style="cursor:pointer" ><%else%><a style="cursor:pointer"><%end if%><%if rs("fAudit") <> "" then%><%=rs("fAudit") %><%else%>δ���<%end if%> </a></td>
					<%if Session("CRM_level") = 9 then%>
					<td class="td_l_c">
						<input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Plugin_Finance_InfoEdit<%=rs("Id")%>()' style="cursor:pointer" /> <input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Plugin_Finance_InfoDel<%=rs("Id")%>()' style="cursor:pointer" />
					</td>
					<%end if%>
				</tr>
				<%
				Contents = rs("fRemark")
				Contents = Replace(Contents, "<p>", "")
				Contents = Replace(Contents, "</p>", "")
				Contents = Replace(Contents,  Chr(10), "")
				Contents = Replace(Contents,  Chr(13), "")
				%>
				<script>function ContentView<%=rs("Id")%>() {art.dialog({ title: 'Ԥ��', content: '<%=Contents%>',width:'98%',height:'98%',drag: false,resize: false});};</script>
				<script>function Plugin_Finance_InfoEdit<%=rs("Id")%>() {$.dialog.open('GetUpdate.asp?action=InfoEdit&Id=<%=rs("Id")%>', {title: '�༭', width: 700,height: 420, fixed: true}); };</script>
				<script>function InfoAudit<%=rs("Id")%>() {$.dialog.open('GetUpdate.asp?action=Audit&sType=A&Id=<%=rs("Id")%>', {title: '���', width: 300,height: 100, fixed: true}); };</script>
				<script>function Plugin_Finance_InfoDel<%=rs("Id")%>() {art.dialog({content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {art.dialog.open('?action=delete&Id=<%=rs("Id")%>');art.dialog.close();},cancelVal: '�ر�',cancel: true});};</script>
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
			<% Contrs = conn.execute ("select sum(fDebit) as Debit,sum(fCredit) as Credit from [Plugin_Finance] where 1=1 "&sql&" ") %>
			<span class="r" style="font-size:14px;padding-top:5px;">����ˮ�ʻ��ܡ����� �� <font color=red><%=Contrs("Debit")%></font> RMB������ �� <font color=red><%=Contrs("Credit")%></font> RMB������� �� <font color=red><%=Plugin_Finance_Cash+Contrs("Debit")-Contrs("Credit")%></font> RMB�� </span>
			<%=EasyCrm.pagelist("index.asp", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<%
elseif otype="Bank" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10">
		
			<span  style="float:left;padding:0 10px;height:34px;text-align:left;position:fixed;right:10px;top:45px;color:#000;">
				<input type="button" name="Add" class="button41" value="ɸѡ" onclick="Showhiden(this,'boxBankSearch',false,'ɸѡ','ɸѡ')" style="cursor:pointer"  />
				<input type="button" name="Add" class="button45" value="����" onclick='Plugin_Finance_Bank_InfoAdd()' style="cursor:pointer"  />
				<script>function Plugin_Finance_Bank_InfoAdd() {$.dialog.open('GetUpdate.asp?action=AddBank', {title: '����', width: 700,height: 420, fixed: true}); };</script>
			</span>
			<span  style="float:left;padding:0 10px;text-align:left;position:fixed;right:10px;top:80px;color:#000;background:#666;">
						<form name="searchForm" action="?subAction=searchBank&otype=Bank" method="post">
						<table width="250" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" id="boxBankSearch" style="display:none;margin:10px 0;background-color:#fff;">
							<col width="90" /><col width="160" />
							<tr>
								<td class="td_l_r title">��ʼ����</td>
								<td class="td_r_l"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="20" onFocus="WdatePicker()" value="<%=Session("Search_Plugin_Finance_Bank_TimeBegin")%>" /></td>
							</tr>
							<tr>
								<td class="td_l_r title">��������</td>
								<td class="td_r_l"><input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="20" onFocus="WdatePicker()" value="<%=Session("Search_Plugin_Finance_Bank_TimeEnd")%>" /> </td>
							</tr>
							<tr>
								<td class="td_l_r title" >����</td>
								<td class="td_r_l">
									<select name="bName" class="int" style="width:130px;">
										<option value="">��ѡ��</option>
										<%
										str = split(""&Plugin_Finance_Bank_Name&"",",")
										for i = 0 to ubound(str)
										if Session("Search_Plugin_Finance_Bank_Name") = str(i) then
										response.Write "<option value="&str(i)&" selected>"&str(i)&"</option>"
										else
										response.Write "<option value="&str(i)&">"&str(i)&"</option>"
										end if
										next
										%>
									</select>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title" >����</td>
								<td class="td_r_l">
									<select name="bClass" class="int" style="width:130px;">
										<option value="">��ѡ��</option>
										<%
										str = split(""&Plugin_Finance_Bank_Class&"",",")
										for i = 0 to ubound(str)
										if Session("Search_Plugin_Finance_Bank_Class") = str(i) then
										response.Write "<option value="&str(i)&" selected>"&str(i)&"</option>"
										else
										response.Write "<option value="&str(i)&">"&str(i)&"</option>"
										end if
										next
										%>
									</select>
								</td>
							</tr>
							<tr>
								<td class="td_l_r title" >���</td>
								<td class="td_r_l">
									<input name="bType" type="radio" class="noborder" value="bDebit"> ��+��
									<input name="bType" type="radio" class="noborder" value="bCredit"> ��-��
								</td>
							</tr>
							<tr>
								<td class="td_l_c" colspan="2" style="padding:5px 0;">
									<input type="submit" name="Submit" class="button42" value=" <%=L_Search%> ">��
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?SubAction=killBankSession" /></td>
								</tr>
						</table>   
						</form>
			</span>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td width="80" class="td_l_c">���</td>
					<td width="100" class="td_l_c">����</td>
					<td width="120" class="td_l_c">����</td>
					<td width="100" class="td_l_c">Ʊ��</td>
					<td width="120" class="td_l_c">����</td>
					<td width="150" class="td_l_c">�ʺ�</td>
					<td width="100" class="td_l_c">��(+)</td>
					<td width="100" class="td_l_c">��(-)</td>
					<td class="td_l_l">ժҪ</td>
					<% if Session("CRM_level")=9 then%>
					<td width="80" class="td_l_c">���</td>
					<%end if%>
					<td width="90" class="td_l_c">����</td>
				</tr>
	<%
	PN = CLng(ABS(Request("PN")))
	If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
	intPageSize = Plugin_Finance_Page
	pagenum = intPageSize*(PN-1)
	Set rs = Server.CreateObject("ADODB.Recordset")
	IF PN=1 THEN
	rs.Open "Select top "&intPageSize&" * From [Plugin_Finance_Bank] where 1=1 "&sqlb&" Order By Id desc ",conn,1,1 
	ELSE
	rs.Open "Select top "&intPageSize&" * From [Plugin_Finance_Bank] where 1=1 "&sqlb&" and Id < ( SELECT Min(Id) FROM ( SELECT TOP "&pagenum&" Id FROM [Plugin_Finance_Bank] where 1=1 "&sqlb&" ORDER BY Id desc ) AS T ) Order By Id desc ",conn,1,1
	END IF
	SQLstr="Select count(Id) As RecordSum From [Plugin_Finance_Bank] where 1=1 "&sqlb&""
	Set Rsstr=conn.Execute(SQLstr,1,1) 
	TotalRecords=Rsstr("RecordSum") 
	if Int(TotalRecords/intPageSize)=TotalRecords/intPageSize then
	TotalPages=TotalRecords/intPageSize
	else
	TotalPages=Int(TotalRecords/intPageSize)+1
	end if
	Rsstr.Close 
	Set Rsstr=Nothing
	If PN > TotalPages Then PN = TotalPages
	i=0
	Do While Not rs.BOF And Not rs.EOF
	i=i+1
	%>
				<tr class="tr">
					<td class="td_l_c"><%=rs("ID") %></td>
					<td class="td_l_c"><%=EasyCrm.FormatDate(rs("bTime"),2)%></td>
					<td class="td_l_c"><%=rs("bName") %></td>
					<td class="td_l_c"><%=rs("bInvoice") %></td>
					<td class="td_l_c"><%=rs("bClass") %></td>
					<td class="td_l_c"><%=rs("bCard") %></td>
					<td class="td_l_c"><%=rs("bDebit") %></td>
					<td class="td_l_c"><%=rs("bCredit") %></td>
					<td class="td_l_l"><%=rs("bDigest") %></td>
					<td class="td_l_c"><% if Session("CRM_level")=9 then%><a onclick='InfoAudit<%=rs("Id")%>()' style="cursor:pointer" ><%else%><a style="cursor:pointer"><%end if%><%if rs("bAudit") <> "" then%><%=rs("bAudit") %><%else%>δ���<%end if%> </a></td>
					<%if Session("CRM_level") = 9 then%>
					<td class="td_l_c">
						<input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Plugin_Finance_Bank_InfoEdit<%=rs("Id")%>()' style="cursor:pointer" /> <input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Plugin_Finance_Bank_InfoDel<%=rs("Id")%>()' style="cursor:pointer" />
					</td>
					<%end if%>
				</tr>
				<script>function Plugin_Finance_Bank_InfoEdit<%=rs("Id")%>() {$.dialog.open('GetUpdate.asp?action=InfoEditBank&Id=<%=rs("Id")%>', {title: '�༭', width: 700,height: 420, fixed: true}); };</script>
				<script>function InfoAudit<%=rs("Id")%>() {$.dialog.open('GetUpdate.asp?action=Audit&sType=B&Id=<%=rs("Id")%>', {title: '���', width: 300,height: 100, fixed: true}); };</script>
				<script>function Plugin_Finance_Bank_InfoDel<%=rs("Id")%>() {art.dialog({content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {art.dialog.open('?action=deleteBank&Id=<%=rs("Id")%>');art.dialog.close();},cancelVal: '�ر�',cancel: true});};</script>
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
			<% Contrs = conn.execute ("select sum(bDebit) as Debit,sum(bCredit) as Credit from [Plugin_Finance_Bank] where 1=1 "&sqlb&" ") %>
			<span class="r" style="font-size:14px;padding-top:5px;">�����д����ܡ����� �� <font color=red><%=Contrs("Debit")%></font> RMB������ �� <font color=red><%=Contrs("Credit")%></font> RMB������� �� <font color=red><%=Plugin_Finance_Card+Contrs("Debit")-Contrs("Credit")%></font> RMB�� </span>
			<%=EasyCrm.pagelist("index.asp?oTyp=Bank", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<%

elseif otype="Outin" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10">
		
			<span  style="float:left;padding:0 10px;height:34px;text-align:left;position:fixed;right:10px;top:45px;color:#000;">
				<input type="button" name="Add" class="button41" value="ɸѡ" onclick="Showhiden(this,'boxOutinSearch',false,'ɸѡ','ɸѡ')" style="cursor:pointer"  />
				<input type="button" name="Add" class="button45" value="����" onclick='Plugin_Finance_Outin_InfoAdd()' style="cursor:pointer"  />
				<script>function Plugin_Finance_Outin_InfoAdd() {$.dialog.open('GetUpdate.asp?action=AddOutin', {title: '����', width: 700,height: 420, fixed: true}); };</script>
			</span>
			<span  style="float:left;padding:0 10px;text-align:left;position:fixed;right:10px;top:80px;color:#000;background:#666;">
						<form name="searchForm" action="?subAction=searchOutin&otype=Outin" method="post">
						<table width="250" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" id="boxOutinSearch" style="display:none;margin: 10px 0;background:#fff;">
							<col width="90" /><col width="160" />
							<tr>
								<td class="td_l_r title">��ʼ����</td>
								<td class="td_r_l"><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate" size="20" onFocus="WdatePicker()" value="<%=Session("Search_Plugin_Finance_Outin_TimeBegin")%>" /></td>
							</tr>
							<tr>
								<td class="td_l_r title">��������</td>
								<td class="td_r_l"><input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate" size="20" onFocus="WdatePicker()" value="<%=Session("Search_Plugin_Finance_Outin_TimeEnd")%>" /> </td>
							</tr>
							<tr>
								<td class="td_l_r title" >��˾����</td>
								<td class="td_r_l">
									<input name="oCompany" type="text" id="oCompany" class="int" size="20" value="<%=Session("Search_Plugin_Finance_Outin_oCompany")%>" />
								</td>
							</tr>
							<tr>
								<td class="td_l_r title" >����</td>
								<td class="td_r_l">
									<input name="osType" type="radio" class="noborder" value="oDebit" <%if Session("Search_Plugin_Finance_Outin_osType") = "oDebit" then%>checked<%end if%>> ���롡
									<input name="osType" type="radio" class="noborder" value="oCredit" <%if Session("Search_Plugin_Finance_Outin_osType") = "oCredit" then%>checked<%end if%>> ֧����
								</td>
							</tr>
							<tr>
								<td class="td_l_c" colspan="2" style="padding:5px 0;">
									<input type="submit" name="Submit" class="button42" value=" <%=L_Search%> ">��
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?SubAction=killOutinSession" /></td>
								</tr>
						</table>   
						</form>
			</span>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td width="80" class="td_l_c">���</td>
					<td width="100" class="td_l_c">����</td>
					<td width="150" class="td_l_c">Ʊ��</td>
					<td width="200" class="td_l_l">��˾����</td>
					<td width="100" class="td_l_c">����(+)</td>
					<td width="100" class="td_l_c">֧��(-)</td>
					<td class="td_l_l">ժҪ</td>
					<td width="80" class="td_l_c">״̬</td>
					<% if Session("CRM_level")=9 then%>
					<td width="80" class="td_l_c">���</td>
					<%end if%>
					<td width="90" class="td_l_c">����</td>
				</tr>
	<%
	PN = CLng(ABS(Request("PN")))
	If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
	intPageSize = Plugin_Finance_Page
	pagenum = intPageSize*(PN-1)
	Set rs = Server.CreateObject("ADODB.Recordset")
	IF PN=1 THEN
	rs.Open "Select top "&intPageSize&" * From [Plugin_Finance_Outin] where 1=1 "&sqlc&" Order By Id desc ",conn,1,1 
	ELSE
	rs.Open "Select top "&intPageSize&" * From [Plugin_Finance_Outin] where 1=1 "&sqlc&" and Id < ( SELECT Min(Id) FROM ( SELECT TOP "&pagenum&" Id FROM [Plugin_Finance_Outin] where 1=1 "&sqlc&" ORDER BY Id desc ) AS T ) Order By Id desc ",conn,1,1
	END IF
	SQLstr="Select count(Id) As RecordSum From [Plugin_Finance_Outin] where 1=1 "&sqlc&""
	Set Rsstr=conn.Execute(SQLstr,1,1) 
	TotalRecords=Rsstr("RecordSum") 
	if Int(TotalRecords/intPageSize)=TotalRecords/intPageSize then
	TotalPages=TotalRecords/intPageSize
	else
	TotalPages=Int(TotalRecords/intPageSize)+1
	end if
	Rsstr.Close 
	Set Rsstr=Nothing
	If PN > TotalPages Then PN = TotalPages
	i=0
	Do While Not rs.BOF And Not rs.EOF
	i=i+1
	%>
				<tr class="tr">
					<td class="td_l_c"><%=rs("ID") %></td>
					<td class="td_l_c"><%=EasyCrm.FormatDate(rs("oTime"),2)%></td>
					<td class="td_l_c"><%=rs("oInvoice") %></td>
					<td class="td_l_l"><%=rs("oCompany") %></td>
					<td class="td_l_c"><%=rs("oDebit") %></td>
					<td class="td_l_c"><%=rs("oCredit") %></td>
					<td class="td_l_l"><%=rs("oDigest") %></td>
					<td class="td_l_c"><%=rs("oState") %></td>
					<td class="td_l_c"><% if Session("CRM_level")=9 then%><a onclick='InfoAudit<%=rs("Id")%>()' style="cursor:pointer" ><%else%><a style="cursor:pointer"><%end if%><%if rs("oAudit") <> "" then%><%=rs("oAudit") %><%else%>δ���<%end if%> </a></td>
					<%if Session("CRM_level") = 9 then%>
					<td class="td_l_c">
						<input type="button" class="button_info_edit" value=" " title="<%=L_Edit%>"  onclick='Plugin_Finance_Outin_InfoEdit<%=rs("Id")%>()' style="cursor:pointer" /> 
						<input type="button" class="button_info_del" value=" " title="<%=L_Del%>" onclick='Plugin_Finance_Outin_InfoDel<%=rs("Id")%>()' style="cursor:pointer" />
					</td>
					<%end if%>
				</tr>
				<script>function Plugin_Finance_Outin_InfoEdit<%=rs("Id")%>() {$.dialog.open('GetUpdate.asp?action=InfoEditOutin&Id=<%=rs("Id")%>', {title: '�༭', width: 700,height: 420, fixed: true}); };</script>
				<script>function InfoAudit<%=rs("Id")%>() {$.dialog.open('GetUpdate.asp?action=Audit&sType=C&Id=<%=rs("Id")%>', {title: '���', width: 300,height: 100, fixed: true}); };</script>
				<script>function Plugin_Finance_Outin_InfoDel<%=rs("Id")%>() {art.dialog({content: '<%=Alert_del_YN%>',icon: 'error',ok: function () {art.dialog.open('?action=deleteOutin&Id=<%=rs("Id")%>');art.dialog.close();},cancelVal: '�ر�',cancel: true});};</script>
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
			<% Contrs = conn.execute ("select sum(oDebit) as Debit,sum(oCredit) as Credit from [Plugin_Finance_Outin] where 1=1 "&sqlb&" ") %>
			<span class="r" style="font-size:14px;padding-top:5px;">����֧���ܡ������� �� <font color=red><%=Contrs("Debit")%></font> RMB����֧�� �� <font color=red><%=Contrs("Credit")%></font> RMB����ӯ�� �� <font color=red><%=Contrs("Debit")-Contrs("Credit")%></font> RMB�� </span>
			<%=EasyCrm.pagelist("index.asp?oTyp=Outin", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<%

end if
end sub

Sub infoManage()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr> 
		<td valign="top" class="td_n">
		
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" colspan=2 style="padding:10px 10px 0;" class="td_n">
						
						<form name="Managesave" action="?action=Managesave" method="post" onSubmit="return CheckInput();">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1" style="border-bottom:0px;">
							<tr class="tr_t"style="border-bottom:0px;"> 
								<td class="td_l_l">
								<B>�߼�����</B>
								</td>
							</tr>
						</table> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr">
								<td class="td_l_c title">��ʼ��ˮ���</td>
								<td class="td_l_l">
									�� <input name="Plugin_Finance_Cash" type="text" class="int" id="Plugin_Finance_Cash" size="10" value="<%=Plugin_Finance_Cash%>" <%if Plugin_Finance_Cash_lock="1" then%>readonly<%end if%>> RMB�� <input name="Plugin_Finance_Cash_lock" type="checkbox" style="vertical-align:middle; " id="Plugin_Finance_Cash_lock" size="40" value="1" <%if Plugin_Finance_Cash_lock="1" then%>checked<%end if%>> ����
								</td>
							</tr>
							<tr class="tr">
								<td class="td_l_c title">��ʼ���д��</td>
								<td class="td_l_l">
									�� <input name="Plugin_Finance_Card" type="text" class="int" id="Plugin_Finance_Card" size="10" value="<%=Plugin_Finance_Card%>" <%if Plugin_Finance_Card_lock="1" then%>readonly<%end if%>> RMB�� <input name="Plugin_Finance_Card_lock" type="checkbox" style="vertical-align:middle; " id="Plugin_Finance_Card_lock" size="40" value="1" <%if Plugin_Finance_Card_lock="1" then%>checked<%end if%>> ����
								</td>
							</tr>
							<tr class="tr">
								<td class="td_l_c title">��ҳ����</td>
								<td class="td_r_l">
									<input name="Plugin_Finance_Page" type="text" class="int" id="Plugin_Finance_Page" size="40" value="<%=Plugin_Finance_Page%>"> <span class="info_help help01">ÿҳ��ʾXX��</span>
								</td>
							</tr>
							<tr >
								<td class="td_l_c title">��ˮ��</td>
								<td class="td_r_l" style="padding:10px;">
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
									<col width="100" />
										<tr class="tr">
											<td class="td_l_c title">�Է���Ŀ</td>
											<td class="td_r_l">
												<input name="Plugin_Finance_Subjects" type="text" class="int" id="Plugin_Finance_Subjects" size="40" value="<%=Plugin_Finance_Subjects%>"> <span class="info_help help01">�������֮���ð�Ƕ��ŷָ�</span>
											</td>
										</tr>
										<tr class="tr">
											<td class="td_l_c title">����</td>
											<td class="td_r_l">
												<input name="Plugin_Finance_Class" type="text" class="int" id="Plugin_Finance_Class" size="40" value="<%=Plugin_Finance_Class%>"> <span class="info_help help01">�������֮���ð�Ƕ��ŷָ�</span>
											</td>
										</tr>
										<tr class="tr">
											<td class="td_l_c title">��Ӧ��Ŀ</td>
											<td class="td_r_l">
												<input name="Plugin_Finance_Project" type="text" class="int" id="Plugin_Finance_Project" size="40" value="<%=Plugin_Finance_Project%>"> <span class="info_help help01">�������֮���ð�Ƕ��ŷָ�</span>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr >
								<td class="td_l_c title">���д��</td>
								<td class="td_r_l" style="padding:10px;">
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
									<col width="100" />
										<tr class="tr">
											<td class="td_l_c title">����</td>
											<td class="td_r_l">
												<input name="Plugin_Finance_Bank_Name" type="text" class="int" id="Plugin_Finance_Bank_Name" size="40" value="<%=Plugin_Finance_Bank_Name%>"> <span class="info_help help01">�������֮���ð�Ƕ��ŷָ�</span>
											</td>
										</tr>
										<tr class="tr">
											<td class="td_l_c title">����</td>
											<td class="td_r_l">
												<input name="Plugin_Finance_Bank_Class" type="text" class="int" id="Plugin_Finance_Bank_Class" size="40" value="<%=Plugin_Finance_Bank_Class%>"> <span class="info_help help01">�������֮���ð�Ƕ��ŷָ�</span>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr >
								<td class="td_l_c title">���Ȩ��</td>
								<td class="td_r_l" style="padding:10px;">
									<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
									<col width="100" />
									<%
										Set rsg = Server.CreateObject("ADODB.Recordset")
										rsg.Open "Select * From [system_group]",conn,1,1
										Do While Not rsg.BOF And Not rsg.EOF
									%>
										<tr> 
											<td class="td_l_c title"><%=rsg("gName")%></td>
											<td  class="td_l_l">
											<%
												Set rsm = Server.CreateObject("ADODB.Recordset")
												rsm.Open "Select * From [user] where uGroup="&rsg("gId")&" ",conn,1,1
												Do While Not rsm.BOF And Not rsm.EOF
											%>
											<input type="checkbox" name="Plugin_Finance_manage" value="<%=rsm("uName")%>" <%if inStr(Plugin_Finance_manage,rsm("uName"))>0 then%>checked<%end if%>> <%=rsm("uName")%>��
											<%
												rsm.MoveNext
												Loop
												rsm.Close
												Set rsm = Nothing
											%>
											</td>
										</tr> 
									<%
										rsg.MoveNext
										Loop
										rsg.Close
										Set rsg = Nothing
									%>
									</table>
								</td>
							</tr>
							<tr>
								<td class="td_r_l" colspan="4">
								<input type="submit" name="Submit" class="button45" value=" <%=L_Edit%> ">��
								<input name="Back" type="button" id="Back" class="button43" value=" <%=L_Back%> " onClick="history.back();">
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
<%
End Sub

Sub infoManagesave()
	Plugin_Finance_Cash = replace(Trim(Request.Form("Plugin_Finance_Cash")),CHR(34),"'")
	if Plugin_Finance_Cash = "" then Plugin_Finance_Cash=0
	Plugin_Finance_Cash_lock = replace(Trim(Request.Form("Plugin_Finance_Cash_lock")),CHR(34),"'")
	Plugin_Finance_Card = replace(Trim(Request.Form("Plugin_Finance_Card")),CHR(34),"'")
	if Plugin_Finance_Card = "" then Plugin_Finance_Card=0
	Plugin_Finance_Card_lock = replace(Trim(Request.Form("Plugin_Finance_Card_lock")),CHR(34),"'")
	Plugin_Finance_Subjects = replace(Trim(Request.Form("Plugin_Finance_Subjects")),CHR(34),"'")
	Plugin_Finance_Class = replace(Trim(Request.Form("Plugin_Finance_Class")),CHR(34),"'")
	Plugin_Finance_Project = replace(Trim(Request.Form("Plugin_Finance_Project")),CHR(34),"'")
	Plugin_Finance_Bank_Name = replace(Trim(Request.Form("Plugin_Finance_Bank_Name")),CHR(34),"'")
	Plugin_Finance_Bank_Class = replace(Trim(Request.Form("Plugin_Finance_Bank_Class")),CHR(34),"'")
	Plugin_Finance_Page = replace(Trim(Request.Form("Plugin_Finance_Page")),CHR(34),"'")
	if Plugin_Finance_Page = "" then Plugin_Finance_Page=10
	Plugin_Finance_manage = replace(Trim(Request.Form("Plugin_Finance_manage")),CHR(34),"'")
	Dim TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "Dim Plugin_Finance_Cash,Plugin_Finance_Cash_lock,Plugin_Finance_Card,Plugin_Finance_Card_lock,Plugin_Finance_Subjects,Plugin_Finance_Class,Plugin_Finance_Project,Plugin_Finance_Bank_Name,Plugin_Finance_Bank_Class,Plugin_Finance_Page,Plugin_Finance_manage" & VbCrLf
	
	TempStr = TempStr & "'��������" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_Cash="& Chr(34) & Plugin_Finance_Cash & Chr(34) &" '" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_Cash_lock="& Chr(34) & Plugin_Finance_Cash_lock & Chr(34) &" '" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_Card="& Chr(34) & Plugin_Finance_Card & Chr(34) &" '" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_Card_lock="& Chr(34) & Plugin_Finance_Card_lock & Chr(34) &" '" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_Subjects="& Chr(34) & Plugin_Finance_Subjects & Chr(34) &" '" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_Class="& Chr(34) & Plugin_Finance_Class & Chr(34) &" '" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_Project="& Chr(34) & Plugin_Finance_Project & Chr(34) &" '" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_Bank_Name="& Chr(34) & Plugin_Finance_Bank_Name & Chr(34) &" '" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_Bank_Class="& Chr(34) & Plugin_Finance_Bank_Class & Chr(34) &" '" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_Page="& Chr(34) & Plugin_Finance_Page & Chr(34) &" '" & VbCrLf
	TempStr = TempStr & "Plugin_Finance_manage="& Chr(34) & Plugin_Finance_manage & Chr(34) &" 'Ȩ��" & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf
	ADODB_SaveToFile TempStr,"Config.asp"
	Response.Write("<script>alert(""�޸ĳɹ���"");</script>")
	Response.Write "<script>location.href='?action=List&otype=Main';</script>"
End Sub

Sub infodelete()
    Dim Id
	Id = CLng(ABS(Request("Id")))
	If Not IsNumeric(Id) Or Id <= 0 Then Response.Write "<script>alert(""������"");</script>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Plugin_Finance] Where Id = " & Id,conn,3,2
	If rs.RecordCount <> 1 Then Response.Write "<script>alert(""������"");</script>"
	Id = rs("Id")
	rs.Delete
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

Sub infodeleteBank()
    Dim Id
	Id = CLng(ABS(Request("Id")))
	If Not IsNumeric(Id) Or Id <= 0 Then Response.Write "<script>alert(""������"");</script>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Plugin_Finance_Bank] Where Id = " & Id,conn,3,2
	If rs.RecordCount <> 1 Then Response.Write "<script>alert(""������"");</script>"
	Id = rs("Id")
	rs.Delete
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

Sub infodeleteOutin()
    Dim Id
	Id = CLng(ABS(Request("Id")))
	If Not IsNumeric(Id) Or Id <= 0 Then Response.Write "<script>alert(""������"");</script>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Plugin_Finance_Outin] Where Id = " & Id,conn,3,2
	If rs.RecordCount <> 1 Then Response.Write "<script>alert(""������"");</script>"
	Id = rs("Id")
	rs.Delete
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
End Sub

Sub ADODB_SaveToFile(ByVal strBody,ByVal File)
	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If
	
	If FSFlag = 1 Then
		Set WriteFile = fs.CreateTextFile(Server.MapPath(File),True)
		WriteFile.Write strBody
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "����������֧��ADODB.Stream���޷���ɲ�������ʹ��FTP�ȹ��ܣ���<font color=Red >data/config.asp</font>�ļ������滻�ɿ�������"
			Err.Clear
			Set objStream = Noting
			Exit Sub
		End If
		With objStream
			.Type = 2
			.Open
			.Charset = "GB2312"
			.Position = objStream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	End If
End Sub

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
		</td>
	</tr>
</table>
<script src="../../data/calendar/WdatePicker.js"></script>
<%else%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10">   
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>������ʾ</B></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr>
					<td class="td_r_l" style="border-top:0;">
						����Ȩʹ�øò����
					</td>
				</tr>
			</table>   
		</td>
	</tr>
</table>

<%end if%>
</body>
</html><% Set EasyCrm = nothing %>
