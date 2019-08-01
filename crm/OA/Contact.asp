<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 72, 1) = 1 Then %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script src="<%=SiteUrl&skinurl%>Js/Common.js" type="text/javascript"></script>
</head>
<body> 
<style>body{padding:35px 0 48px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n"><%=L_Here%>：<%=L_Page_OA%> > <%=L_Page_Contact%></td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">

	<tr> 
		<td valign="top" class="td_n pd10">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td valign="top" colspan=2 class="td_n">
						<table width="100%" border="0" cellpadding="0" cellspacing="0" class="table_1">
							<tr class="tr_t">
								<td width="60" class="td_l_c"><%=L_User_uId%></td>
								<td width="100" class="td_l_c"><%=L_User_uName%></td>
								<td width="120" class="td_l_c"><%=L_User_uMobile%></td>
								<td class="td_l_l"><%=L_User_uEmail%></td>
								<td width="100" class="td_l_c"><%=L_User_uBirthday%></td>
								<td width="100" class="td_l_c"><%=L_User_uAddtime%></td>
							</tr>
						<%
						Dim rs
						
						Dim intTotalRecords,intTotalPages,PN,intPageSize
						PN = CLng(ABS(Request("PN")))

						If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
						intPageSize = DataPageSize
						pagenum = intPageSize*(PN-1)
						
						Set rs = Server.CreateObject("ADODB.Recordset")
							IF PN=1 THEN
							rs.Open "Select top "&intPageSize&" * From [user] Order By uId asc ",conn,1,1 
							ELSE
							rs.Open "Select top "&intPageSize&" * From [user] where uId > ( SELECT Max(uId) FROM ( SELECT TOP "&pagenum&" uId FROM [user] ORDER BY uId asc ) AS T ) Order By uId asc ",conn,1,1
							END IF
							SQLstr="Select count(uId) As RecordSum From [user] "
							
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
								<TD class="td_l_c"><%=rs("uid")%></TD>
								<TD class="td_l_c"><%=rs("uName")%></a></TD>
								<TD class="td_l_c"><%=rs("uMobile")%></TD>
								<TD class="td_l_l"><%=rs("uEmail")%></TD>
								<TD class="td_l_c"><%=EasyCrm.FormatDate(rs("uBirthday"),2)%></TD>
								<TD class="td_l_c"><%=EasyCrm.FormatDate(rs("uaddtime"),2)%></TD>
							</TR>
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
        </td>
	</tr>
</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd ">
			<%=EasyCrm.pagelist("Contact.asp", PN,TotalPages,TotalService)%>
		</td> 
	</tr>
</table>
</div>
</body>
</html><%else%>无权限<%end if%><% Set EasyCrm = nothing %>
