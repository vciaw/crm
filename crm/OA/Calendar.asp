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
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Page_OA%> > <%=L_Page_Calendar%></td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<FORM NAME="pageForm" ACTION="calendar.asp" METHOD="GET">

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
			<div class="MenuboxS">
				 <ul>
					<li class="hover"><span><a href="?">月历视图</a></span></li>
					<li class=""><span><a href="#" onclick="location.href='Calendar_List.asp';" style="cursor:pointer">列表视图</a></span></li>
				</ul>
			</div>
		</td>
		<td class="td_l_r pdr10" COLSPAN="6" style="border-right:0;padding-top:5px;">
			<span class="tips01" style="float:left;padding:0 10px;height:34px;text-align:left;position:fixed;right:10px;top:43px;color:#000;">
			<input type="button" class="button221" value="上月" onClick=window.location.href="calendar.asp?currentDate=<%= strPrevMonth %>" />　
			<B>
			<SELECT NAME="currentDate" CLASS="int" onChange="pageForm.submit()">
			<% For counter = 1 to 12 %>
				<OPTION VALUE="<%= DateSerial(Year(dtCurrentDate), counter, 1) %>" <% If (DatePart("m", dtCurrentDate) = counter) Then Response.Write "SELECTED"%>><%= MonthName(counter) & " " & Year(dtCurrentDate) %></OPTION>
			<% Next %>
			</SELECT>　
			</B>
			<input type="button" class="button221" value="下月" onClick=window.location.href="calendar.asp?currentDate=<%= strNextMonth %>" />　
			<input type="button" class="button247" value="回到当月" onClick=window.location.href="calendar.asp" />
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
			Response.Write ("<p style='background:#FFF;text-align:center;font-size:16px;font-weight:bold;cursor:pointer' onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='编辑'><span style='float:right;font-size:12px;color:#999;padding-right:5px;'>"&nl(dtOnDay)&"</span><span style='float:left;padding-left:10px;'>" & aCalendarDays((iWeek-1)*7 + iDay) & "</span>"&L_Calendar_today&"</p><p class=""Calendarinfo"" onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' ><span style='float:right;font-size:12px;padding-right:5px;cursor:pointer' onClick='InfoDel"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='删除'>×</span>"&dailyMsg&"</p>")
		else
			Response.Write ("<p style='background:#FFF;text-align:left;font-size:16px;font-weight:bold;padding-left:10px;cursor:pointer' onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='编辑'><span style='float:right;font-size:12px;color:#999;padding-right:5px;'>"&nl(dtOnDay)&"</span>" & aCalendarDays((iWeek-1)*7 + iDay) & "</p><p class=""Calendarinfo""><span style='float:right;font-size:12px;padding-right:5px;cursor:pointer' onClick='InfoDel"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='删除'>×</span><span  onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()'>"&dailyMsg&"</span></p>")
		end if
 	else
		If dtOnDay = dtToday Then
			Response.Write ("<p style='background:#FFF;text-align:center;font-size:16px;font-weight:bold;cursor:pointer' onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='添加'><span style='float:right;font-size:12px;color:#999;padding-right:5px;'>"&nl(dtOnDay)&"</span><span style='float:left;padding-left:10px;color:#333;'>" & aCalendarDays((iWeek-1)*7 + iDay) & "</span>"&L_Calendar_today&"</p><p style=""text-align:center;""><input type='button' class='button242' value='"&L_Calendar_add&"' onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' /></p>")
		else
		Response.Write ("<p style=""font-size:16px;font-weight:bold;padding-left:10px;color:#333;cursor:pointer"" onClick='InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"()' title='添加' ><span style='float:right;font-size:12px;color:#999;padding-right:5px;'>"&nl(dtOnDay)&"</span>" & aCalendarDays((iWeek-1)*7 + iDay) & "</p>")
		end if
	end if
	Response.Write ("<script>function InfoAdd"&EasyCrm.FormatDate(dtOnDay,14)&"() {$.dialog.open('" & strPage & "', {title: '编辑', width: 400, height: 280,fixed: true}); };</script>")
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
'获取当前系统时间
if str="" then
curTime = Now()
else
curTime = str
end if
Dim WeekName(7), MonthAdd(11), NongliData(99), TianGan(9), DiZhi(11), ShuXiang(11), DayName(30), MonName(12)
'星期名
WeekName(0) = " * "
WeekName(1) = "星期日"
WeekName(2) = "星期一"
WeekName(3) = "星期二"
WeekName(4) = "星期三"
WeekName(5) = "星期四"
WeekName(6) = "星期五"
WeekName(7) = "星期六"
'农历日期名
DayName(0) = "*"
DayName(1) = "初一"
DayName(2) = "初二"
DayName(3) = "初三"
DayName(4) = "初四"
DayName(5) = "初五"
DayName(6) = "初六"
DayName(7) = "初七"
DayName(8) = "初八"
DayName(9) = "初九"
DayName(10) = "初十"
DayName(11) = "十一"
DayName(12) = "十二"
DayName(13) = "十三"
DayName(14) = "十四"
DayName(15) = "十五"
DayName(16) = "十六"
DayName(17) = "十七"
DayName(18) = "十八"
DayName(19) = "十九"
DayName(20) = "二十"
DayName(21) = "廿一"
DayName(22) = "廿二"
DayName(23) = "廿三"
DayName(24) = "廿四"
DayName(25) = "廿五"
DayName(26) = "廿六"
DayName(27) = "廿七"
DayName(28) = "廿八"
DayName(29) = "廿九"
DayName(30) = "三十"
'农历月份名
MonName(0) = "*"
MonName(1) = "正"
MonName(2) = "二"
MonName(3) = "三"
MonName(4) = "四"
MonName(5) = "五"
MonName(6) = "六"
MonName(7) = "七"
MonName(8) = "八"
MonName(9) = "九"
MonName(10) = "十"
MonName(11) = "十一"
MonName(12) = "腊"
'公历每月前面的天数
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
'农历数据
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
'生成当前公历年、月、日 ==> GongliStr
curYear = Year(curTime)
curMonth = Month(curTime)
curDay = Day(curTime)
GongliStr = curYear & "年"
If (curMonth < 10) Then
GongliStr = GongliStr & "0" & curMonth & "月"
Else
GongliStr = GongliStr & curMonth & "月"
End If
If (curDay < 10) Then
GongliStr = GongliStr & "0" & curDay & "日"
Else
GongliStr = GongliStr & curDay & "日"
End If
'生成当前公历星期 ==> WeekdayStr
curWeekday = Weekday(curTime)
WeekdayStr = WeekName(curWeekday)
'计算到初始时间1921年2月8日的天数：1921-2-8(正月初一)
TheDate = (curYear - 1921) * 365 + Int((curYear - 1921) / 4) + curDay + MonthAdd(curMonth - 1) - 38
If ((curYear Mod 4) = 0 And curMonth > 2) Then
TheDate = TheDate + 1
End If
'计算农历天干、地支、月、日
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
'获取NongliData(m)的第n个二进制位的值
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
'生成农历天干、地支、属相 ==> NongliStr
NongliStr = "农历" & TianGan(((curYear - 4) Mod 60) Mod 10) & DiZhi(((curYear - 4) Mod 60) Mod 12) & "年"
NongliStr = NongliStr & "(" & ShuXiang(((curYear - 4) Mod 60) Mod 12) & ")"
'生成农历月、日 ==> NongliDayStr
if EasyCrm.FormatDate(str,10) = "01/01" then
NongliDayStr = NongliDayStr & "<font color=green>元旦</font>"
elseif EasyCrm.FormatDate(str,10) = "02/14" then
NongliDayStr = NongliDayStr & "<font color=green>情人节</font>"
elseif EasyCrm.FormatDate(str,10) = "03/08" then
NongliDayStr = NongliDayStr & "<font color=green>妇女节</font>"
elseif EasyCrm.FormatDate(str,10) = "04/01" then
NongliDayStr = NongliDayStr & "<font color=green>愚人节</font>"
elseif EasyCrm.FormatDate(str,10) = "05/01" then
NongliDayStr = NongliDayStr & "<font color=green>劳动节</font>"
elseif EasyCrm.FormatDate(str,10) = "05/04" then
NongliDayStr = NongliDayStr & "<font color=green>青年节</font>"
elseif EasyCrm.FormatDate(str,10) = "06/01" then
NongliDayStr = NongliDayStr & "<font color=green>儿童节</font>"
elseif EasyCrm.FormatDate(str,10) = "07/01" then
NongliDayStr = NongliDayStr & "<font color=green>建党 港归</font>"
elseif EasyCrm.FormatDate(str,10) = "08/01" then
NongliDayStr = NongliDayStr & "<font color=green>建军节</font>"
elseif EasyCrm.FormatDate(str,10) = "09/10" then
NongliDayStr = NongliDayStr & "<font color=green>教师节</font>"
elseif EasyCrm.FormatDate(str,10) = "10/01" then
NongliDayStr = NongliDayStr & "<font color=green>国庆节</font>"
elseif EasyCrm.FormatDate(str,10) = "12/25" then
NongliDayStr = NongliDayStr & "<font color=green>圣诞节</font>"
elseif curMonth = "1" and curDay = "1" then
NongliDayStr = NongliDayStr & "<font color=red>春节</font>"
elseif curMonth = "1" and curDay = "15" then
NongliDayStr = NongliDayStr & "<font color=red>元宵节</font>"
elseif curMonth = "5" and curDay = "5" then
NongliDayStr = NongliDayStr & "<font color=red>端午节</font>"
elseif curMonth = "7" and curDay = "7" then
NongliDayStr = NongliDayStr & "<font color=red>七夕</font>"
elseif curMonth = "8" and curDay = "15" then
NongliDayStr = NongliDayStr & "<font color=red>中秋节</font>"
elseif curMonth = "9" and curDay = "9" then
NongliDayStr = NongliDayStr & "<font color=red>重阳节</font>"
elseif curMonth = "12" and curDay = "8" then
NongliDayStr = NongliDayStr & "<font color=red>腊八节</font>"
else
If (curMonth < 1) Then
NongliDayStr = "闰" & MonName(-1 * curMonth)
Else
NongliDayStr = MonName(curMonth)
End If
NongliDayStr = NongliDayStr & "."
NongliDayStr = NongliDayStr & DayName(curDay)
end if
nl = NongliDayStr
End Function
'response.write nl(now()) '输出结果
%><%else%>无权限<%end if%><% Set EasyCrm = nothing %>