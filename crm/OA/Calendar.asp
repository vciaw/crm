<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 73, 1) = 1 Then %>
<% 
Session("CRM_thispage") = "Calendar.asp"
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
If Request("currentDate") <> "" Then
  dtCurrentDate = Request("currentDate")
Else
  dtCurrentDate = dtToday
End If
iFirstDayOfMonth = DatePart("w", DateSerial(Year(dtCurrentDate), Month(dtCurrentDate), 1))
iDaysInMonth = DatePart("d", DateSerial(Year(dtCurrentDate), Month(dtCurrentDate)+1, 1-1))

For counter = 1 to iDaysInMonth
  aCalendarDays(counter + iFirstDayOfMonth - 1) = counter
Next

iColumns = 7
iRows= 6 - Int((42 - (iFirstDayOfMonth + iDaysInMonth - 1 )) / 7)
strPrevMonth = Server.URLEncode(DateAdd("m", -1, dtCurrentDate))
strNextMonth = Server.URLEncode(DateAdd("m",  1, dtCurrentDate))
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>
</head>

<body style="padding-top:35px;"> 
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>��<%=L_Page_OA%> > <%=L_Page_Calendar%></td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="ˢ��" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="����" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="ǰ��" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<FORM NAME="pageForm" ACTION="calendar.asp" METHOD="GET">

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
			<div class="MenuboxS">
				 <ul>
					<li class="hover"><span><a href="?">������ͼ</a></span></li>
					<li class=""><span><a href="#" onclick="location.href='Calendar_List.asp';" style="cursor:pointer">�б���ͼ</a></span></li>
				</ul>
			</div>
		</td>
		<td class="td_l_r pdr10" COLSPAN="6" style="border-right:0;padding-top:5px;">
			<span class="tips01" style="float:left;padding:0 10px;height:34px;text-align:left;position:fixed;right:10px;top:43px;color:#000;">
			<input type="button" class="button221" value="����" onClick=window.location.href="calendar.asp?currentDate=<%= strPrevMonth %>" />��
			<B>
			<SELECT NAME="currentDate" CLASS="int" onChange="pageForm.submit()">
			<% For counter = 1 to 12 %>
				<OPTION VALUE="<%= DateSerial(Year(dtCurrentDate), counter, 1) %>" <% If (DatePart("m", dtCurrentDate) = counter) Then Response.Write "SELECTED"%>><%= MonthName(counter) & " " & Year(dtCurrentDate) %></OPTION>
			<% Next %>
			</SELECT>��
			</B>
			<input type="button" class="button221" value="����" onClick=window.location.href="calendar.asp?currentDate=<%= strNextMonth %>" />��
			<input type="button" class="button247" value="�ص�����" onClick=window.location.href="calendar.asp" />
			</span>
		</td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<% For iDay = vbSunday To vbSaturday %>
					<TD class="td_l_c" width="14%"><%= WeekDayName(iDay, True) %></TD>
					<% Next %>
				</tr>
<%
    For iWeek = 1 To iRows
	Response.Write "<TR>"
	For iDay = 1 To iColumns
If aCalendarDays((iWeek-1)*7 + iDay) > 0 then
	dtOnDay = DateSerial(Year(dtCurrentDate), Month(dtCurrentDate), aCalendarDays((iWeek-1)*7 + iDay))
	
	'if Accsql= 1 then
	'strSQL = "SELECT * FROM [calendar] WHERE calendaruser='" & Session("CRM_name") & "' and calendarDate = '" & dtOnDay & "'"
	'else
	'strSQL = "SELECT * FROM [calendar] WHERE calendaruser='" & Session("CRM_name") & "' and calendarDate = #" & dtOnDay & "#"
	'end if
	'Set RS = Conn.Execute(strSQL)
	'If NOT RS.EOF Then 
	'    id = RS("id")
	'    dailyMsg = RS("calendarText")
	'    dailyuser = RS("calendaruser")
	'Else 
	'    dailyMsg = ""
	'End If
	'Set RS = Nothing
	if Accsql = 1 then 
	dailyMsg = EasyCrm.getNewItem("calendar","calendaruser","'"&Session("CRM_name")&"' and DATEDIFF(d,calendarDate,'"&dtOnDay&"')=0 ","calendarText")
	id = EasyCrm.getNewItem("calendar","calendaruser","'"&Session("CRM_name")&"' and DATEDIFF(d,calendarDate,'"&dtOnDay&"')=0 ","id")
	else
	dailyMsg = EasyCrm.getNewItem("calendar","calendaruser","'"&Session("CRM_name")&"' and DATEDIFF('d',calendarDate,'"&dtOnDay&"')=0 ","calendarText")
	id = EasyCrm.getNewItem("calendar","calendaruser","'"&Session("CRM_name")&"' and DATEDIFF('d',calendarDate,'"&dtOnDay&"')=0 ","id")
	end if
	if dailyMsg = "0" then dailyMsg=""

    Response.Write "<TD valign=top CLASS='td_l_c' style='height:65px;text-align:left;vertical-align:top;color:#d00;padding:0px'>"

	IF dailyMsg<>"" THEN
		strPage = "GetUpdate.asp?action=Calendar&sType=Edit&id="&id
	ELSE
		strPage = "GetUpdate.asp?action=Calendar&sType=Add&currentDate=" & dtOnDay
	END IF

    If (Trim(dailyMsg) = Trim(Left(dailyMsg, intCharToShow))) Then
	Else 
	dailyMsg = Trim(Left(dailyMsg, intCharToShow-4)) & " "
    End If
	
	if len(dailyMsg)>7 then
	dailyMsg = EasyCrm.ReName(left(dailyMsg,7))&"..."
	end if
	
    if FormatStr(dailyMsg)<>"" then
		If dtOnDay = dtToday Then
			Response.Write ("<p style='background:#FFF;text-align:center;font-size:16px;font-weight:bold;cursor:pointer' onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='�༭'><span style='float:right;font-size:12px;color:#999;padding-right:5px;'>"&nl(dtOnDay)&"</span><span style='float:left;padding-left:10px;'>" & aCalendarDays((iWeek-1)*7 + iDay) & "</span>"&L_Calendar_today&"</p><p class=""Calendarinfo"" onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' ><span style='float:right;font-size:12px;padding-right:5px;cursor:pointer' onClick='InfoDel"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='ɾ��'>��</span>"&dailyMsg&"</p>")
		else
			Response.Write ("<p style='background:#FFF;text-align:left;font-size:16px;font-weight:bold;padding-left:10px;cursor:pointer' onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='�༭'><span style='float:right;font-size:12px;color:#999;padding-right:5px;'>"&nl(dtOnDay)&"</span>" & aCalendarDays((iWeek-1)*7 + iDay) & "</p><p class=""Calendarinfo""><span style='float:right;font-size:12px;padding-right:5px;cursor:pointer' onClick='InfoDel"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='ɾ��'>��</span><span  onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()'>"&dailyMsg&"</span></p>")
		end if
 	else
		If dtOnDay = dtToday Then
			Response.Write ("<p style='background:#FFF;text-align:center;font-size:16px;font-weight:bold;cursor:pointer' onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='���'><span style='float:right;font-size:12px;color:#999;padding-right:5px;'>"&nl(dtOnDay)&"</span><span style='float:left;padding-left:10px;color:#333;'>" & aCalendarDays((iWeek-1)*7 + iDay) & "</span>"&L_Calendar_today&"</p><p style=""text-align:center;""><input type='button' class='button242' value='"&L_Calendar_add&"' onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' /></p>")
		else
		Response.Write ("<p style=""font-size:16px;font-weight:bold;padding-left:10px;color:#333;cursor:pointer"" onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='���' ><span style='float:right;font-size:12px;color:#999;padding-right:5px;'>"&nl(dtOnDay)&"</span>" & aCalendarDays((iWeek-1)*7 + iDay) & "</p>")
		end if
	end if
	Response.Write ("<script>function InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"() {$.dialog.open('" & strPage & "', {title: '�༭', width: 400, height: 280,fixed: true}); };</script>")
	Response.Write ("<script>function InfoDel"&EasyCrm.FormatDate(dtOnDay,14)&"() {$.dialog.open('GetUpdate.asp?action=Calendar&sType=Del&id="&id&"');art.dialog.close(); };</script>")
Else 
		Response.Write ("<TD CLASS='td_l_c' style='height:65px;'>")
End IF

		Response.Write "</TD>"
	Next
		Response.Write "</TR>"
    Next
    Conn.Close
    set Conn = Nothing
%>
			</table>
		</td> 
	</tr>
</TABLE>
</FORM>
</body>
</html>

<%
Function nl(str)
'��ȡ��ǰϵͳʱ��
if str="" then
curTime = Now()
else
curTime = str
end if
Dim WeekName(7), MonthAdd(11), NongliData(99), TianGan(9), DiZhi(11), ShuXiang(11), DayName(30), MonName(12)
'������
WeekName(0) = " * "
WeekName(1) = "������"
WeekName(2) = "����һ"
WeekName(3) = "���ڶ�"
WeekName(4) = "������"
WeekName(5) = "������"
WeekName(6) = "������"
WeekName(7) = "������"
'ũ��������
DayName(0) = "*"
DayName(1) = "��һ"
DayName(2) = "����"
DayName(3) = "����"
DayName(4) = "����"
DayName(5) = "����"
DayName(6) = "����"
DayName(7) = "����"
DayName(8) = "����"
DayName(9) = "����"
DayName(10) = "��ʮ"
DayName(11) = "ʮһ"
DayName(12) = "ʮ��"
DayName(13) = "ʮ��"
DayName(14) = "ʮ��"
DayName(15) = "ʮ��"
DayName(16) = "ʮ��"
DayName(17) = "ʮ��"
DayName(18) = "ʮ��"
DayName(19) = "ʮ��"
DayName(20) = "��ʮ"
DayName(21) = "إһ"
DayName(22) = "إ��"
DayName(23) = "إ��"
DayName(24) = "إ��"
DayName(25) = "إ��"
DayName(26) = "إ��"
DayName(27) = "إ��"
DayName(28) = "إ��"
DayName(29) = "إ��"
DayName(30) = "��ʮ"
'ũ���·���
MonName(0) = "*"
MonName(1) = "��"
MonName(2) = "��"
MonName(3) = "��"
MonName(4) = "��"
MonName(5) = "��"
MonName(6) = "��"
MonName(7) = "��"
MonName(8) = "��"
MonName(9) = "��"
MonName(10) = "ʮ"
MonName(11) = "ʮһ"
MonName(12) = "��"
'����ÿ��ǰ�������
MonthAdd(0) = 0
MonthAdd(1) = 31
MonthAdd(2) = 59
MonthAdd(3) = 90
MonthAdd(4) = 120
MonthAdd(5) = 151
MonthAdd(6) = 181
MonthAdd(7) = 212
MonthAdd(8) = 243
MonthAdd(9) = 273
MonthAdd(10) = 304
MonthAdd(11) = 334
'ũ������
NongliData(0) = 2635
NongliData(1) = 333387
NongliData(2) = 1701
NongliData(3) = 1748
NongliData(4) = 267701
NongliData(5) = 694
NongliData(6) = 2391
NongliData(7) = 133423
NongliData(8) = 1175
NongliData(9) = 396438
NongliData(10) = 3402
NongliData(11) = 3749
NongliData(12) = 331177
NongliData(13) = 1453
NongliData(14) = 694
NongliData(15) = 201326
NongliData(16) = 2350
NongliData(17) = 465197
NongliData(18) = 3221
NongliData(19) = 3402
NongliData(20) = 400202
NongliData(21) = 2901
NongliData(22) = 1386
NongliData(23) = 267611
NongliData(24) = 605
NongliData(25) = 2349
NongliData(26) = 137515
NongliData(27) = 2709
NongliData(28) = 464533
NongliData(29) = 1738
NongliData(30) = 2901
NongliData(31) = 330421
NongliData(32) = 1242
NongliData(33) = 2651
NongliData(34) = 199255
NongliData(35) = 1323
NongliData(36) = 529706
NongliData(37) = 3733
NongliData(38) = 1706
NongliData(39) = 398762
NongliData(40) = 2741
NongliData(41) = 1206
NongliData(42) = 267438
NongliData(43) = 2647
NongliData(44) = 1318
NongliData(45) = 204070
NongliData(46) = 3477
NongliData(47) = 461653
NongliData(48) = 1386
NongliData(49) = 2413
NongliData(50) = 330077
NongliData(51) = 1197
NongliData(52) = 2637
NongliData(53) = 268877
NongliData(54) = 3365
NongliData(55) = 531109
NongliData(56) = 2900
NongliData(57) = 2922
NongliData(58) = 398042
NongliData(59) = 2395
NongliData(60) = 1179
NongliData(61) = 267415
NongliData(62) = 2635
NongliData(63) = 661067
NongliData(64) = 1701
NongliData(65) = 1748
NongliData(66) = 398772
NongliData(67) = 2742
NongliData(68) = 2391
NongliData(69) = 330031
NongliData(70) = 1175
NongliData(71) = 1611
NongliData(72) = 200010
NongliData(73) = 3749
NongliData(74) = 527717
NongliData(75) = 1452
NongliData(76) = 2742
NongliData(77) = 332397
NongliData(78) = 2350
NongliData(79) = 3222
NongliData(80) = 268949
NongliData(81) = 3402
NongliData(82) = 3493
NongliData(83) = 133973
NongliData(84) = 1386
NongliData(85) = 464219
NongliData(86) = 605
NongliData(87) = 2349
NongliData(88) = 334123
NongliData(89) = 2709
NongliData(90) = 2890
NongliData(91) = 267946
NongliData(92) = 2773
NongliData(93) = 592565
NongliData(94) = 1210
NongliData(95) = 2651
NongliData(96) = 395863
NongliData(97) = 1323
NongliData(98) = 2707
NongliData(99) = 265877
'���ɵ�ǰ�����ꡢ�¡��� ==> GongliStr
curYear = Year(curTime)
curMonth = Month(curTime)
curDay = Day(curTime)
GongliStr = curYear & "��"
If (curMonth < 10) Then
GongliStr = GongliStr & "0" & curMonth & "��"
Else
GongliStr = GongliStr & curMonth & "��"
End If
If (curDay < 10) Then
GongliStr = GongliStr & "0" & curDay & "��"
Else
GongliStr = GongliStr & curDay & "��"
End If
'���ɵ�ǰ�������� ==> WeekdayStr
curWeekday = Weekday(curTime)
WeekdayStr = WeekName(curWeekday)
'���㵽��ʼʱ��1921��2��8�յ�������1921-2-8(���³�һ)
TheDate = (curYear - 1921) * 365 + Int((curYear - 1921) / 4) + curDay + MonthAdd(curMonth - 1) - 38
If ((curYear Mod 4) = 0 And curMonth > 2) Then
TheDate = TheDate + 1
End If
'����ũ����ɡ���֧���¡���
isEnd = 0
m = 0
Do
If (NongliData(m) < 4095) Then
k = 11
Else
k = 12
End If
n = k
Do
If (n < 0) Then
Exit Do
End If
'��ȡNongliData(m)�ĵ�n��������λ��ֵ
bit = NongliData(m)
For i = 1 To n Step 1
bit = Int(bit / 2)
Next
bit = bit Mod 2
If (TheDate <= 29 + bit) Then
isEnd = 1
Exit Do
End If
TheDate = TheDate - 29 - bit
n = n - 1
Loop
If (isEnd = 1) Then
Exit Do
End If
m = m + 1
Loop
curYear = 1921 + m
curMonth = k - n + 1
curDay = TheDate
If (k = 12) Then
If (curMonth = (Int(NongliData(m) / 65536) + 1)) Then
curMonth = 1 - curMonth
ElseIf (curMonth > (Int(NongliData(m) / 65536) + 1)) Then
curMonth = curMonth - 1
End If
End If
'����ũ����ɡ���֧������ ==> NongliStr
NongliStr = "ũ��" & TianGan(((curYear - 4) Mod 60) Mod 10) & DiZhi(((curYear - 4) Mod 60) Mod 12) & "��"
NongliStr = NongliStr & "(" & ShuXiang(((curYear - 4) Mod 60) Mod 12) & ")"
'����ũ���¡��� ==> NongliDayStr
if EasyCrm.FormatDate(str,10) = "01/01" then
NongliDayStr = NongliDayStr & "<font color=green>Ԫ��</font>"
elseif EasyCrm.FormatDate(str,10) = "02/14" then
NongliDayStr = NongliDayStr & "<font color=green>���˽�</font>"
elseif EasyCrm.FormatDate(str,10) = "03/08" then
NongliDayStr = NongliDayStr & "<font color=green>��Ů��</font>"
elseif EasyCrm.FormatDate(str,10) = "04/01" then
NongliDayStr = NongliDayStr & "<font color=green>���˽�</font>"
elseif EasyCrm.FormatDate(str,10) = "05/01" then
NongliDayStr = NongliDayStr & "<font color=green>�Ͷ���</font>"
elseif EasyCrm.FormatDate(str,10) = "05/04" then
NongliDayStr = NongliDayStr & "<font color=green>�����</font>"
elseif EasyCrm.FormatDate(str,10) = "06/01" then
NongliDayStr = NongliDayStr & "<font color=green>��ͯ��</font>"
elseif EasyCrm.FormatDate(str,10) = "07/01" then
NongliDayStr = NongliDayStr & "<font color=green>���� �۹�</font>"
elseif EasyCrm.FormatDate(str,10) = "08/01" then
NongliDayStr = NongliDayStr & "<font color=green>������</font>"
elseif EasyCrm.FormatDate(str,10) = "09/10" then
NongliDayStr = NongliDayStr & "<font color=green>��ʦ��</font>"
elseif EasyCrm.FormatDate(str,10) = "10/01" then
NongliDayStr = NongliDayStr & "<font color=green>�����</font>"
elseif EasyCrm.FormatDate(str,10) = "12/25" then
NongliDayStr = NongliDayStr & "<font color=green>ʥ����</font>"
elseif curMonth = "1" and curDay = "1" then
NongliDayStr = NongliDayStr & "<font color=red>����</font>"
elseif curMonth = "1" and curDay = "15" then
NongliDayStr = NongliDayStr & "<font color=red>Ԫ����</font>"
elseif curMonth = "5" and curDay = "5" then
NongliDayStr = NongliDayStr & "<font color=red>�����</font>"
elseif curMonth = "7" and curDay = "7" then
NongliDayStr = NongliDayStr & "<font color=red>��Ϧ</font>"
elseif curMonth = "8" and curDay = "15" then
NongliDayStr = NongliDayStr & "<font color=red>�����</font>"
elseif curMonth = "9" and curDay = "9" then
NongliDayStr = NongliDayStr & "<font color=red>������</font>"
elseif curMonth = "12" and curDay = "8" then
NongliDayStr = NongliDayStr & "<font color=red>���˽�</font>"
else
If (curMonth < 1) Then
NongliDayStr = "��" & MonName(-1 * curMonth)
Else
NongliDayStr = MonName(curMonth)
End If
NongliDayStr = NongliDayStr & "."
NongliDayStr = NongliDayStr & DayName(curDay)
end if
nl = NongliDayStr
End Function
'response.write nl(now()) '������
%><%else%>��Ȩ��<%end if%><% Set EasyCrm = nothing %>