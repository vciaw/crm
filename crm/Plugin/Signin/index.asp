<!--#include file="../../data/conn.asp" --><!--#include file="config.asp" --><!--#include file="../../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
'��ȡ��ǰҳ��
PNN = Trim(Request.QueryString("PN"))
if PNN="" then PNN=1 
otype	=	Request.QueryString("otype")
if otype="" then otype="Main"

if Trim(Request("SubAction")) = "Search" then
	Dim sStime,sEtime,sUser,sState,sql
	
	sUser = EasyCrm.Searchcode(Request("User"))
	sState = EasyCrm.Searchcode(Request("State"))
	sStime = EasyCrm.Searchcode(Request("sStime"))
	sEtime = EasyCrm.Searchcode(Request("sEtime"))
	Session("CRM_Plugin_Signin_sUser") = EasyCrm.Searchcode(Request("User"))
	Session("CRM_Plugin_Signin_sState") = EasyCrm.Searchcode(Request("State"))
	Session("CRM_Plugin_Signin_sStime") = EasyCrm.Searchcode(Request("sStime"))
	Session("CRM_Plugin_Signin_sEtime") = EasyCrm.Searchcode(Request("sEtime"))
	sql = ""
	
	If sUser <> "" Then
	    sql = sql & " And sUser = '" & sUser & "' "
	End If
	
	If sState <> "" Then
	    sql = sql & " And ( sSstate = '" & sState & "' or sEstate = '" & sState & "' ) "
	End If
	
	if Accsql=1 then
	If sStime <> "" Then
	    sql = sql & " And sDate >= '" & sStime & "' "
	End If
			
	If sEtime <> "" Then
	    sql = sql & " And sDate <= '" & sEtime & "' "
	End If
	else
	If sStime <> "" Then
	    sql = sql & " And sDate >= #" & sStime & "# "
	End If
			
	If sEtime <> "" Then
	    sql = sql & " And sDate <= #" & sEtime & "# "
	End If
	End If
	
end if

If sUser = "" And sStime = "" And sEtime = "" Then
    If Session("CRM_Plugin_Signin_Search") <> "" Then
        sql = Session("CRM_Plugin_Signin_Search")
	End If
Else
    Session("CRM_Plugin_Signin_Search") = sql
End If

if Trim(Request("SubAction")) = "killSession" then
	Session("CRM_Plugin_Signin_Search") = ""
	Session("CRM_Plugin_Signin_sUser") = ""
	Session("CRM_Plugin_Signin_sState") = ""
	Session("CRM_Plugin_Signin_sStime") = ""
	Session("CRM_Plugin_Signin_sEtime") = ""
	Response.Write "<script>location.href='?action=Infolist&otype=Infolist' ;</script>"
end if

	Dim intTotalRecords,intTotalPages,PN,intPageSize'��¼��������ҳ������ǰҳ����ҳ����
	PN = CLng(ABS(Request("PN")))

    If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
    intPageSize = DataPageSize
	pageNums = intPageSize*(PN-1)

		Set rs = Server.CreateObject("ADODB.Recordset")
		IF PN=1 THEN
	    rs.Open "Select top "&intPageSize&" * From [Plugin_Signin] where 1=1 "&sql&" Order By sId desc ",conn,1,1 
		ELSE
	    rs.Open "Select top "&intPageSize&" * From [Plugin_Signin] where 1=1 "&sql&" and sId < ( SELECT Min(sId) FROM ( SELECT TOP "&pageNums&" sId FROM [Plugin_Signin] where  1=1 "&sql&" ORDER BY sId desc ) AS T ) Order By sId desc ",conn,1,1
		END IF
		SQLstr="Select count(sId) As RecordSum From [Plugin_Signin] where 1=1 "&sql&" " 'ͳ��ҳ��

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

Dim i
i = 0
Do While Not rs.BOF And Not rs.EOF
    i = i + 1
	strToPrint = strToPrint & "			<tr class=""tr"">" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("sId") & "</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & EasyCrm.FormatDate(rs("sDate"),2) & "</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("sUser") & "</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("sSstate") & "</td>" & VBCrlf
	if rs("sStart") = 1 then
	strToPrint = strToPrint & "				<td class=""td_l_c"">��</td>" & VBCrlf
	else
	strToPrint = strToPrint & "				<td class=""td_l_c""></td>" & VBCrlf
	end if
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("sStime") & "</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("sEstate") & "</td>" & VBCrlf
	if rs("sEnd") = 1 then
	strToPrint = strToPrint & "				<td class=""td_l_c"">��</td>" & VBCrlf
	else
	strToPrint = strToPrint & "				<td class=""td_l_c""></td>" & VBCrlf
	end if
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("sEtime") & "</td>" & VBCrlf
    strToPrint = strToPrint & "        		<td class=""td_l_c"">" & VBCrlf
    strToPrint = strToPrint & "        			<input type=""button"" class=""button_info_del"" value='��' title="""&L_Del&""" onClick=""window.location.href='?action=delete&sid=" & rs("sId") & "&PN="&PN&"'"" onClick=""return confirm('"&Alert_del_YN&"');"" />" & VBCrlf
    strToPrint = strToPrint & "        		</td>" & VBCrlf
	strToPrint = strToPrint & "			</tr>" & VBCrlf
    If i >= intPageSize Then Exit Do
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Function FormatStr(String)
  String = Replace(String, CHR(13), "")
  String = Replace(String, CHR(10) & CHR(10), "</P><P>")
  String = Replace(String, CHR(10), "<BR>")
  FormatStr = String
End Function

Const intCharToShow = 20
Const bolEditable   = true

Dim dtToday,dtCurrentDate,aCalendarDays(42),iFirstDayOfMonth,iDaysInMonth,iColumns, iRows	, iDay, iWeek
Dim counter,strNextMonth, strPrevMonth,dailyMsg,dailyuser,dtOnDay,strPage

dtToday = Date()
iFirstDayOfMonth = DatePart("w", DateSerial(Year(dtToday), Month(dtToday), 1))
iDaysInMonth = DatePart("d", DateSerial(Year(dtToday), Month(dtToday)+1, 1-1))

For counter = 1 to iDaysInMonth
  aCalendarDays(counter + iFirstDayOfMonth - 1) = counter
Next

iColumns = 7
iRows= 6 - Int((42 - (iFirstDayOfMonth + iDaysInMonth - 1 )) / 7)
strPrevMonth = Server.URLEncode(DateAdd("m", -1, dtToday))
strNextMonth = Server.URLEncode(DateAdd("m",  1, dtToday))
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<style>
.button_ico_yes {border:none; height:23px; cursor:pointer; *padding:2px 0 0 0; background:url(<%=SiteUrl&skinurl%>images/ico.gif) no-repeat;font-weight:normal;font-size:12px;padding-left:18px; width:80px; color:#666; background-position:left -69px;}
.button_ico_no {border:none; height:23px; cursor:pointer; *padding:2px 0 0 0; background:url(<%=SiteUrl&skinurl%>images/ico.gif) no-repeat;font-weight:normal;font-size:12px;padding-left:18px; width:80px; color:#666; background-position:-80px -69px;}
</style>
</head>

<body style="padding-top:35px;">

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">��ǰλ�ã����ܲ�� > Ա��ǩ��</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="ˢ��" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="����" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="ǰ��" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
            <div class="MenuboxS">
              <ul>
                <li <%if otype="Main" then%>class="hover"<%end if%>><span><a href="?action=Main&otype=Main">ǩ����</a></span></li>
				<%if inStr(Plugin_Signin_manage,session("CRM_name"))>0 then%>
                <li <%if otype="Infolist" then%>class="hover"<%end if%>><span><a href="?action=Infolist&otype=Infolist">�����б�</a></span></li>
				<%end if%><%if session("CRM_level") = 9 then%>
                <li <%if otype="Manage" then%>class="hover"<%end if%>><span><a href="?action=Manage&otype=Manage">�߼�����</a></span></li>
				<%end if%>
              </ul>
            </div>
		</td>
	</tr>
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10">
<%
action = Trim(Request("action"))
Select Case action
Case "add"
    Call infoadd()
Case "Infolist"
    Call Infolist()
Case "Export"
    Call Export()
Case "Manage"
    Call infoManage()
Case "Managesave"
    Call infoManagesave()
Case "delete"
    Call infodelete()
Case "install"
    Call install()
Case Else
    Call Main()
End Select
%>

<%
Sub Main()
%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="7"><B>��ǰ�·ݣ�<%=EasyCrm.FormatDate(now(),8)%> - <script language="javascript" src="time.js"></script></B></TD>
				</tr>
				<tr class="tr_f">
					<% For iDay = vbSunday To vbSaturday %>
					<TD class="td_l_c" width="14%"><%= WeekDayName(iDay, True) %></TD>
					<% Next %>
				</tr>
<%
    For iWeek = 1 To iRows
	Response.Write "<TR>"
	For iDay = 1 To iColumns
If aCalendarDays((iWeek-1)*7 + iDay) > 0 then
	dtOnDay = DateSerial(Year(dtToday), Month(dtToday), aCalendarDays((iWeek-1)*7 + iDay))
	dtOnMonth = Month(now())
	'��ѯ���յ�ǰԱ�����޼�¼
if accsql=1 then
	Set Rsstr=conn.Execute("Select count(sid) As snum From Plugin_Signin where sUser ='"&Session("CRM_name")&"' and sDate = '"&EasyCrm.FormatDate(now(),2)&"' " ,1,1)
	If dtOnDay = date() Then
	strSQL = "SELECT * FROM Plugin_Signin WHERE sUser='" & Session("CRM_name") & "' and sDate = '"&EasyCrm.FormatDate(now(),2)&"' "
	else
	strSQL = "SELECT * FROM Plugin_Signin WHERE sUser='" & Session("CRM_name") & "' and sDate = '"&EasyCrm.FormatDate(dtOnDay,2)&"' "
	end if
else
	Set Rsstr=conn.Execute("Select count(sid) As snum From Plugin_Signin where sUser ='"&Session("CRM_name")&"' and sDate = #"&EasyCrm.FormatDate(now(),2)&"# " ,1,1)
	If dtOnDay = date() Then
	strSQL = "SELECT * FROM Plugin_Signin WHERE sUser='" & Session("CRM_name") & "' and sDate = #"&EasyCrm.FormatDate(now(),2)&"# "
	else
	strSQL = "SELECT * FROM Plugin_Signin WHERE sUser='" & Session("CRM_name") & "' and sDate = #"&EasyCrm.FormatDate(dtOnDay,2)&"# "
	end if
end if
	Set RS = Conn.Execute(strSQL)
	If NOT RS.EOF Then 
		sId		= RS("sId")'���
		sUser	= RS("sUser")'ǩ����
		sSstate	= RS("sSstate")'���磺�������ٵ�������
		sStart	= RS("sStart")'�Ƿ�ǩ����1�� 0 ��
		sStime	= RS("sStime")'ǩ��ʱ��
		sEstate	= RS("sEstate")'���磺���������ˡ�����
		sEnd	= RS("sEnd")'�Ƿ�ǩ�ˣ�1�� 0 ��
		sEtime	= RS("sEtime")'ǩ��ʱ��
	else
		sStart	= 0
		sEnd	= 0
	End If
	Set RS = Nothing

    Response.Write "<TD valign=top CLASS='td_l_c' style='height:60px;text-align:left;vertical-align:top;color:#d00;padding:5px'>"
	
	If dtOnDay = dtToday Then
		Response.Write ("<p style='text-align:center;font-size:16px;font-weight:bold;'>" & aCalendarDays((iWeek-1)*7 + iDay) & ""&L_Hao&"</p><p  style=""text-align:center"">")
		if Hour(now()) < 12 then
		if sStart = 1 then
		Response.Write ("<input type=""button"" class=""button_ico_yes input_no"" value="" ��ǩ�� "" >")
		else
		Response.Write ("<input type=""button"" class=""button242"" value="" ǩ�� "" onClick=""window.location.href='?action=add&sClass=ǩ��'"">")
		end if 
		else
		if sEnd = 1 then
		Response.Write (" <input type=""button"" class=""button_ico_yes input_no"" value="" ��ǩ "">")
		else
		Response.Write (" <input type=""button"" class=""button242"" value="" ǩ�� "" onClick=""window.location.href='?action=add&sClass=ǩ��'"">")
		end if
		end if
		Response.Write ("</p>")
	else
		if aCalendarDays((iWeek-1)*7 + iDay) < day(now())then
		
		if sStart = 1 or sEnd = 1 then
		Response.Write ("<p style='text-align:center;font-size:16px;color:#000;'>" & aCalendarDays((iWeek-1)*7 + iDay) & ""&L_Hao&"</p><p style=""text-align:center;""><input type=""button"" class=""button_ico_yes input_no"" value="" ��ǩ ""></p>")
		else
		Response.Write ("<p style='text-align:center;font-size:16px;color:#000;'>" & aCalendarDays((iWeek-1)*7 + iDay) & ""&L_Hao&"</p><p style=""text-align:center;""><input type=""button"" class=""button_ico_no input_no"" value="" δǩ ""></p>")
		end if
		
		else
		Response.Write ("<p style=""text-align:center;font-size:16px;color:#000;"">" & aCalendarDays((iWeek-1)*7 + iDay) & ""&L_Hao&"</p>")
		end if
	end if
Else 
		Response.Write ("<TD CLASS='td_l_c' style='height:70px;'>")
End IF

		Response.Write "</TD>"
	Next
		Response.Write "</TR>"
    Next
    Conn.Close
    set Conn = Nothing
%>
			</table>


<%
end Sub

Sub infolist()
%>
				<form name="searchComForm" method="post" action="?SubAction=Search&action=Infolist&otype=Infolist">
				<img src="<%=SiteUrl&skinurl%>images/ico/search.png" id="ico">
					ʱ�䣺<input name="sStime" type="text" maxlength="10" id="sStime" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("CRM_Plugin_Signin_sStime")%>" /> ~ <input name="sEtime" type="text" maxlength="10" id="sEtime" class="Wdate" size="15" onFocus="WdatePicker()" value="<%=Session("CRM_Plugin_Signin_sEtime")%>" />��
					Ա����<% = EasyCrm.UserList(2,"User","") %>��
					״̬��<select name="State" class="int"><option value="">��ѡ��</option><option value="�ٵ�">�ٵ�</option><option value="����">����</option><option value="����">����</option></select>��
					<input name="Search" type="submit" id="Search" class="button42" value="<%=L_Search%>">
					<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?SubAction=killSession" />
				</form>
<script language="JavaScript">
<!--
for(var i=0;i<document.all.User.options.length;i++){
    if(document.all.User.options[i].value == "<% = Session("CRM_Plugin_Signin_sUser") %>"){
    document.all.User.options[i].selected = true;}}
	
for(var i=0;i<document.all.State.options.length;i++){
    if(document.all.State.options[i].value == "<% = Session("CRM_Plugin_Signin_sState") %>"){
    document.all.State.options[i].selected = true;}}
-->
</script>
		</td>
	</tr>
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td class="td_l_l" colspan=3>Ա��������ϸ��¼</td>
					<td class="td_l_c" colspan=3>����</td>
					<td class="td_l_c" colspan=3>����</td>
					<td width="60" class="td_l_c">����</td>
				
				</tr>
				<tr class="tr_f">
					<td width="70" class="td_l_c">���</td>
					<td width="100" class="td_l_c">����</td>
					<td class="td_l_c">����</td>
					<td width="70" class="td_l_c">״̬</td>
					<td width="70" class="td_l_c">����</td>
					<td width="130" class="td_l_c">ǩ��ʱ��</td>
					<td width="70" class="td_l_c">״̬</td>
					<td width="70" class="td_l_c">����</td>
					<td width="130" class="td_l_c">ǩ��ʱ��</td>
					<td class="td_l_c"></td>
				</tr>
				<% = strToPrint %>
				<tr class="tr_b"> 
					<td class="td_l_l" colspan=10><%if sql<>"" then%><input type="submit" name="Submit" class="button245" value="����Excel" onClick="window.location.href='?action=Export'"><%end if%></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<%=EasyCrm.pagelist("index.asp?action=Infolist&otype=Infolist", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<%
end sub
%>
<%
Sub infoadd()
	dim sClass,sState,sUser,sTime,sContent
	sClass = Request.QueryString("sClass")
	'��ѯ���յ�ǰԱ�����޼�¼
	if accsql=1 then
	Set Rsstr=conn.Execute("Select count(sid) As snum From Plugin_Signin where sUser ='"&Session("CRM_name")&"' and sDate = '"&EasyCrm.FormatDate(now(),2)&"' " ,1,1)
	else
	Set Rsstr=conn.Execute("Select count(sid) As snum From Plugin_Signin where sUser ='"&Session("CRM_name")&"' and sDate = #"&EasyCrm.FormatDate(now(),2)&"# " ,1,1)
	end if
	if sClass="ǩ��" then 'δǩ��������£������ڽ���ǩ����¼��ֱ������
		if Rsstr("snum") = 0 then
		if datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_stime&"") >= 0 then
			conn.execute ("insert into Plugin_Signin(sUser,sSstate,sStart,sStime,sDate) values('"&Session("CRM_name")&"','����','1','"&now()&"','"&date()&"')")
		elseif datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_stime&"") < 0 and datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_stime&"") > -60  then
			conn.execute ("insert into Plugin_Signin(sUser,sSstate,sStart,sStime,sDate) values('"&Session("CRM_name")&"','�ٵ�','1','"&now()&"','"&date()&"')")
		else
			conn.execute ("insert into Plugin_Signin(sUser,sSstate,sStart,sStime,sDate) values('"&Session("CRM_name")&"','����','1','"&now()&"','"&date()&"')")
		end if
		else
		Response.Write("<script>alert('�����ظ�ǩ��');</script>")
		end if
	else 'ǩ����Ҫ�ж��Ƿ��м�¼�����޼�¼��������¼�����ڼ�¼�����¼�¼
		if Rsstr("snum") = 0 then '�޼�¼��������¼
			if datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_etime&"") <= 0 then
				conn.execute ("insert into Plugin_Signin(sUser,sEstate,sEnd,sEtime,sDate) values('"&Session("CRM_name")&"','����','1','"&now()&"','"&date()&"')")
			elseif datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_etime&"") > 0 and datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_etime&"") < 60  then
				conn.execute ("insert into Plugin_Signin(sUser,sEstate,sEnd,sEtime,sDate) values('"&Session("CRM_name")&"','����','1','"&now()&"','"&date()&"')")
			else
				conn.execute ("insert into Plugin_Signin(sUser,sEstate,sEnd,sEtime,sDate) values('"&Session("CRM_name")&"','����','1','"&now()&"','"&date()&"')")
			end if
		else '���ڼ�¼�����¼�¼
		if accsql=1 then
			if datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_etime&"") <= 0 then
				conn.execute ("UPDATE Plugin_Signin SET sEstate='����',sEnd='1',sEtime='"&now()&"' Where sUser ='"&Session("CRM_name")&"' and sDate='"&date()&"' ")
			elseif datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_etime&"") > 0 and datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_etime&"") < 60  then
				conn.execute ("UPDATE Plugin_Signin SET sEstate='����',sEnd='1',sEtime='"&now()&"' Where sUser ='"&Session("CRM_name")&"' and sDate='"&date()&"' ")
			else
				conn.execute ("UPDATE Plugin_Signin SET sEstate='����',sEnd='1',sEtime='"&now()&"' Where sUser ='"&Session("CRM_name")&"' and sDate='"&date()&"' ")
			end if
		else
			if datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_etime&"") <= 0 then
				conn.execute ("UPDATE Plugin_Signin SET sEstate='����',sEnd='1',sEtime='"&now()&"' Where sUser ='"&Session("CRM_name")&"' and sDate=#"&date()&"# ")
			elseif datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_etime&"") > 0 and datediff("n",""&Hour(now())&":"&Minute(now())&"",""&Plugin_Signin_etime&"") < 60  then
				conn.execute ("UPDATE Plugin_Signin SET sEstate='����',sEnd='1',sEtime='"&now()&"' Where sUser ='"&Session("CRM_name")&"' and sDate=#"&date()&"# ")
			else
				conn.execute ("UPDATE Plugin_Signin SET sEstate='����',sEnd='1',sEtime='"&now()&"' Where sUser ='"&Session("CRM_name")&"' and sDate=#"&date()&"# ")
			end if
		end if
		end if
	end if
	Response.Write "<script>location.href='index.asp';</script>"
End Sub
%>
<%
Sub Export()
	userfolder = Session("CRM_account") '�����ļ���
	filefolder = Server.MapPath("../../soft/"&userfolder)
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(filefolder) then '����ļ��в������򴴽�
	fso.CreateFolder(filefolder) 
	end if
	dim rs,filename,fs,myfile,x,sql
	Set fs = server.CreateObject("scripting.filesystemobject") 
	filename = Server.MapPath("../../soft/"&userfolder&"/"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls") 
	if fs.FileExists(filename) then
	   fs.DeleteFile(filename) 
	end if
	set myfile = fs.CreateTextFile(filename,true) 
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "select * from Plugin_Signin where 1 = 1 "&Session("CRM_Plugin_Signin_Search")&" Order By sId desc"
	rs.Open sql,conn,1,1
	if rs.EOF and rs.BOF then
	else
	dim strLine,strLiner,responsestr 
	myfile.writeline "����"&chr(9)&"����"&chr(9)&"����"&chr(9)&"����"&chr(9)&"ʱ��"&chr(9)&"����"&chr(9)&"����"&chr(9)&"ʱ��"&chr(9)&""
	Do while Not rs.EOF 
    strLine=""
    for each x in rs.Fields 
    next
	item01 = ""&rs("sDate")&""&chr(9)&""
	item02 = ""&rs("sUser")&""&chr(9)&""
	if rs("sSstate") <> "" then
	item03 = ""&rs("sSstate")&""&chr(9)&""
	else
	item03 = "δǩ"&chr(9)&""
	end if
	if rs("sStart") = 1 then
	item04 = "��"&chr(9)&""
	else
	item04 = "δ֪"&chr(9)&""
	end if
	item05 = ""&rs("sStime")&""&chr(9)&""
	item06 = ""&rs("sEstate")&""&chr(9)&""
	if rs("sEnd") = 1 then
	item07 = "��"&chr(9)&""
	else
	item07 = "δ֪"&chr(9)&""
	end if
	item08 = ""&rs("sEtime")&""&chr(9)&""
    myfile.writeline item01&item02&item03&item04&item05&item06&item07&item08
    rs.MoveNext
	loop
	end if
	
	rs.Close 
	set rs = nothing
	
	conn.execute ("insert into OA_soft(s_class,s_title,s_file,s_user,s_share,s_time) values('"&L_Export_soft&"','"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls','../soft/"&Session("CRM_account")&"/"&year(now())&""&month(now())&""&day(now())&""&hour(now())&""&minute(now())&".xls','"&Session("CRM_name")&"','0','"&now()&"')")
	
	Response.Write("<script>alert('�����ɹ�');</script>")
	Response.Write "<script>location.href='index.asp?action=Infolist&otype=Infolist';</script>"

End Sub

Sub infoManage()
%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>�߼����� <font color="#color:#CC0000">(*)</font></B></td>
				</tr>
			</table>
			<form name="Managesave" action="?action=Managesave" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" />
				<tr>
					<td class="td_l_c title" style="border-top:0;">�ϰ�ʱ��</td>
					<td class="td_r_l" style="border-top:0;">
						<input name="Plugin_Signin_stime" type="text" class="int" id="Plugin_Signin_stime" size="10" value="<%=Plugin_Signin_stime%>" onfocus="WdatePicker({dateFmt:'H:mm'})" class="Wdate"> �ٵ����ϰ��1Сʱ��ǩ��������������1Сʱ
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">�°�ʱ��</td>
					<td class="td_r_l">
						<input name="Plugin_Signin_etime" type="text" class="int" id="Plugin_Signin_etime" size="10" value="<%=Plugin_Signin_etime%>" onfocus="WdatePicker({dateFmt:'H:mm'})" class="Wdate"> ���ˣ��°�ǰ1Сʱ��ǩ��������������1Сʱ
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">����Ȩ��</td>
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
								<input type="checkbox" name="Plugin_Signin_manage" value="<%=rsm("uName")%>" <%if inStr(Plugin_Signin_manage,rsm("uName"))>0 then%>checked<%end if%>> <%=rsm("uName")%>��
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

<%
End Sub

Sub infoManagesave()
	Plugin_Signin_stime = replace(Trim(Request.Form("Plugin_Signin_stime")),CHR(34),"'")
	Plugin_Signin_etime = replace(Trim(Request.Form("Plugin_Signin_etime")),CHR(34),"'")
	Plugin_Signin_manage = replace(Trim(Request.Form("Plugin_Signin_manage")),CHR(34),"'")
	Dim TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "Dim Plugin_Signin_stime,Plugin_Signin_etime,Plugin_Signin_manage" & VbCrLf
	
	TempStr = TempStr & "'��ϸ����" & VbCrLf
	TempStr = TempStr & "Plugin_Signin_stime="& Chr(34) & Plugin_Signin_stime & Chr(34) &" '�ϰ�ʱ��" & VbCrLf
	TempStr = TempStr & "Plugin_Signin_etime="& Chr(34) & Plugin_Signin_etime & Chr(34) &" '�°�ʱ��" & VbCrLf
	TempStr = TempStr & "Plugin_Signin_manage="& Chr(34) & Plugin_Signin_manage & Chr(34) &" 'Ȩ��" & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf
	ADODB_SaveToFile TempStr,"Config.asp"
	Response.Write("<script>alert(""�޸ĳɹ���"");</script>")
	Response.Write "<script>location.href='?action=List&otype=Main';</script>"
End Sub

Sub infodelete()
    Dim sId
	sId = CLng(ABS(Request("sId")))
	If Not IsNumeric(sId) Or sId <= 0 Then Response.Write "<script>alert(""������"");</script>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Plugin_Signin] Where sId = " & sId,conn,3,2
	If rs.RecordCount <> 1 Then Response.Write "<script>alert(""������"");</script>"
	sId = rs("sId")
	rs.Delete
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Write("<script>alert('�����ɹ�');</script>")
	Response.Write "<script>location.href='index.asp?action=Infolist&otype=Infolist';</script>"
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
</body>
</html><% Set EasyCrm = nothing %>
