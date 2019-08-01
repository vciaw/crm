<!--#include file="../../data/conn.asp" --><!--#include file="config.asp" --><!--#include file="../../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
'获取当前页码
PNN = Trim(Request.QueryString("PN"))
subAction = Trim(Request("subAction"))
if PNN="" then PNN=1 
otype	=	Request.QueryString("otype")
if otype="" then otype="Main"

ecarr = split(""&Plugin_Workschedule_State&"",",")
if otype="" or otype="Main" then
	sql = sql & " and wClass <> '草稿' "
elseif otype="Item1" then
	if Accsql=1 then
	sql = sql & " And ( wState <> '" & ecarr(0) & "' or wState is Null ) And wSH = 0 And wCompletiontime > '" & EasyCrm.FormatDate(now(),1) & "' And wClass <> '草稿'"
	else
	sql = sql & " And ( wState <> '" & ecarr(0) & "' or wState is Null ) And wSH = 0 And wCompletiontime > #" & EasyCrm.FormatDate(now(),1) & "# And wClass <> '草稿' "
	end if
elseif otype="Item2" then
	sql = sql & " And wState = '" & ecarr(0) & "' And wSH = 1 and wClass <> '草稿' " '待审核，不限制完成期限
elseif otype="Item3" then
	sql = sql & " And wState = '" & ecarr(0) & "' And wSH = 2 and wClass <> '草稿' " '已完成，不限制完成期限
elseif otype="Item4" then
	if Accsql = 1 then
	sql = sql & " And wSH = 0 And wCompletiontime < '" & EasyCrm.FormatDate(now(),1) & "' and wClass <> '草稿' " '未提交完成并且超过期限
	else
	sql = sql & " And wSH = 0 And wCompletiontime < #" & EasyCrm.FormatDate(now(),1) & "# and wClass <> '草稿' " '未提交完成并且超过期限
	end if
elseif otype="Item5" then
	sql = sql & " And wClass = '草稿' " '草稿
end if
if Session("CRM_level")< 9 then
if inStr(Plugin_Workschedule_manage,session("CRM_name"))=0 then
	sql = sql & " And ( wUserb = '" & Session("CRM_name") & "' or wUsers like '" & Session("CRM_name") & "' )"
end if
end if

	If subAction = "searchItem" Then
	Dim wTitle,wUserb,TimeBegin,TimeEnd
	wTitle = EasyCrm.Searchcode(Request("wTitle"))
	wUserb = EasyCrm.Searchcode(Request("User"))
	TimeBegin = EasyCrm.Searchcode(Request("TimeBegin"))
	TimeEnd = EasyCrm.Searchcode(Request("TimeEnd"))
	Session("Search_Plugin_Workschedule_wTitle") = EasyCrm.Searchcode(Request("wTitle"))
	Session("Search_Plugin_Workschedule_wUserb") = EasyCrm.Searchcode(Request("User"))
	Session("Search_Plugin_Workschedule_TimeBegin") = EasyCrm.Searchcode(Request("TimeBegin"))
	Session("Search_Plugin_Workschedule_TimeEnd") = EasyCrm.Searchcode(Request("TimeEnd"))
	
	If wTitle <> "" Then
	    sql = sql & " And wTitle like '%" & wTitle & "%' "
	End If
	
	If wUserb <> "" Then
	    sql = sql & " And wUserb = '" & wUserb & "' "
	End If
	
	if Accsql =1 then
	If TimeBegin <> "" Then
	    sql = sql & " And wCompletiontime >= '" & TimeBegin & "' "
	End If
			
	If TimeEnd <> "" Then
	    sql = sql & " And wCompletiontime <= '" & TimeEnd & "' "
	End If
	else
	If TimeBegin <> "" Then
	    sql = sql & " And wCompletiontime >= #" & TimeBegin & "# "
	End If
			
	If TimeEnd <> "" Then
	    sql = sql & " And wCompletiontime <= #" & TimeEnd & "# "
	End If
	end if
	
	end if
	
	If wTitle = "" And wUserb = "" And TimeBegin="" And TimeEnd="" Then
		If Session("Search_Plugin_Workschedule_Search") <> "" Then
			sql = Session("Search_Plugin_Workschedule_Search")
		End If
	Else
		Session("Search_Plugin_Workschedule_Search") = sql
	End If

	If subAction = "killSession" Then
		Session("Search_Plugin_Workschedule_Search") = ""
		Session("Search_Plugin_Workschedule_wTitle") = ""
		Session("Search_Plugin_Workschedule_wUserb") = ""
		Session("Search_Plugin_Workschedule_TimeBegin") = ""
		Session("Search_Plugin_Workschedule_TimeEnd") = ""
		Response.Write "<script>location.href='?PN="&PNN&"';</script>"
	End If
	
	Dim intTotalRecords,intTotalPages,PN,intPageSize'记录总数，总页数，当前页，分页数量
	PN = CLng(ABS(Request("PN")))

    If Not IsNumeric(PN) Or PN <= 0 Then PN = 1
    intPageSize = DataPageSize
	pageNums = intPageSize*(PN-1)

		Set rs = Server.CreateObject("ADODB.Recordset")
		IF PN=1 THEN
	    rs.Open "Select top "&intPageSize&" * From [Plugin_Workschedule] where 1=1 "&sql&" Order By wID desc ",conn,1,1 
		ELSE
	    rs.Open "Select top "&intPageSize&" * From [Plugin_Workschedule] where 1=1 "&sql&" and wID < ( SELECT Min(wID) FROM ( SELECT TOP "&pageNums&" wID FROM [Plugin_Workschedule] where  1=1 "&sql&" ORDER BY wID desc ) AS T ) Order By wID desc ",conn,1,1
		END IF
		SQLstr="Select count(wID) As RecordSum From [Plugin_Workschedule] where 1=1 "&sql&" " '统计页码

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

'翻页代码开始
	
	 strCounter = strCounter & " "&EasyCrm.pagelist("index.asp", PN,TotalPages,TotalRecords)&""
	
'翻页代码结束

Dim i
i = 0
Do While Not rs.BOF And Not rs.EOF
    i = i + 1
	strToPrint = strToPrint & "			<tr class=""tr"">" & VBCrlf
	if rs("wYd") = 1 then
	strToPrint = strToPrint & "				<td class=""td_l_c""><img src="""&SiteUrl&skinurl&"images/ico/message_old.png""></td>" & VBCrlf
	else
	strToPrint = strToPrint & "				<td class=""td_l_c""><img src="""&SiteUrl&skinurl&"images/ico/message_new.png""></td>" & VBCrlf
	end if
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("wID") & "</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("wCompletiontime") & "</td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("wClass") & "</td>" & VBCrlf
	if rs("wStar")="" or isnull(rs("wStar")) then
	strToPrint = strToPrint & "				<td class=""td_l_c"">无</td>" & VBCrlf
	elseif  rs("wStar") =1 then
	strToPrint = strToPrint & "				<td class=""td_l_c""><img src=""ico/star.png""></td>" & VBCrlf
	elseif  rs("wStar") =2 then
	strToPrint = strToPrint & "				<td class=""td_l_c""><img src=""ico/star.png""><img src=""ico/star.png""></td>" & VBCrlf
	elseif  rs("wStar") =3 then
	strToPrint = strToPrint & "				<td class=""td_l_c""><img src=""ico/star.png""><img src=""ico/star.png""><img src=""ico/star.png""></td>" & VBCrlf
	elseif  rs("wStar") =4 then
	strToPrint = strToPrint & "				<td class=""td_l_c""><img src=""ico/star.png""><img src=""ico/star.png""><img src=""ico/star.png""><img src=""ico/star.png""></td>" & VBCrlf
	elseif  rs("wStar") =5 then
	strToPrint = strToPrint & "				<td class=""td_l_c""><img src=""ico/star.png""><img src=""ico/star.png""><img src=""ico/star.png""><img src=""ico/star.png""><img src=""ico/star.png""></td>" & VBCrlf
	end if
	strToPrint = strToPrint & "				<td class=""td_l_l""><a href='?action=view&wID=" & rs("wID") & "&PN="&PN&"'>" & rs("wTitle") & "</a></td>" & VBCrlf
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("wUserb") & "</td>" & VBCrlf
	if rs("wSH") = 1 then
	strToPrint = strToPrint & "				<td class=""td_l_c"">待审核</td>" & VBCrlf
	else
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("wState") & "</td>" & VBCrlf
	end if
	if rs("wMsg") = 1 then
	strToPrint = strToPrint & "				<td class=""td_l_c""><img src=""ico/new.gif""></td>" & VBCrlf
	else
	strToPrint = strToPrint & "				<td class=""td_l_c"">暂无</td>" & VBCrlf
	end if
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & rs("wUser") & "</td>" & VBCrlf
	if inStr(Plugin_Workschedule_manage,session("CRM_name"))>0 or Session("CRM_level")=9 then
	strToPrint = strToPrint & "				<td class=""td_l_c"">" & VBCrlf	
	strToPrint = strToPrint & "				<input type=""button"" class=""button_info_edit"" value='　' title="""&L_Edit&""" onClick=""window.location.href='?action=edit&wID=" & rs("wID") & "&PN="&PN&"'"" />" & VBCrlf
	strToPrint = strToPrint & "				<input type=""button"" class=""button_info_del"" value='　' title="""&L_Del&""" onClick="" if(confirm('"&Alert_del_YN&"'))window.location.href='?action=delete&wID=" & rs("wID") & "&PN="&PN&"';else return false;"" />" & VBCrlf
	strToPrint = strToPrint & "				</td>" & VBCrlf
	end if 
	strToPrint = strToPrint & "			</tr>" & VBCrlf
    If i >= intPageSize Then Exit Do
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js" type="text/javascript"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Float.js"></script>
<script language="JavaScript">
<!--
function CheckInput()
{
		if(document.all.wCompletiontime.value == ""){
			alert("任务要求完成时间不能为空！");
			document.all.wCompletiontime.focus();
			return false;
		}
		if(document.all.wTitle.value == ""){
			alert("任务标题不能为空！");
			document.all.wTitle.focus();
			return false;
		}
		if(document.all.wUserb.value == ""){
			alert("负责人不能为空！");
			document.all.wUserb.focus();
			return false;
		}
}
-->
</script>
</head>

<body style="padding-top:35px;">

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：功能插件 > 任务发布</td>
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
                <li <%if otype="Main" then%>class="hover"<%end if%>><span><a href="?action=List&otype=Main">全部事项</a></span></li>
				<li class="" id="CheckA"><span><a href="javascript:void(0)" style="cursor:pointer">高级搜索</a></span></li>
                <li <%if otype="Item1" then%>class="hover"<%end if%>><span><a href="?action=Item1&otype=Item1">待完成</a></span></li>
                <li <%if otype="Item2" then%>class="hover"<%end if%>><span><a href="?action=Item2&otype=Item2">待审核</a></span></li>
                <li <%if otype="Item3" then%>class="hover"<%end if%>><span><a href="?action=Item3&otype=Item3">已完成</a></span></li>
                <li <%if otype="Item4" then%>class="hover"<%end if%>><span><a href="?action=Item4&otype=Item4">未完成</a></span></li>
                <li <%if otype="Item5" then%>class="hover"<%end if%>><span><a href="?action=Item5&otype=Item5">草稿箱</a></span></li>
				<% if inStr(Plugin_Workschedule_manage,session("CRM_name"))>0 or Session("CRM_level") = 9 then %>
                <li <%if otype="Add" then%>class="hover"<%end if%>><span><a href="?action=Add&otype=Add">新增事项</a></span></li>
                <li <%if otype="Manage" then%>class="hover"<%end if%>><span><a href="?action=Manage&otype=Manage">高级管理</a></span></li>
				<%end if%>
              </ul>
            </div>
		</td>
	</tr>
</table>
		<div id="SearchBox" style="position: absolute; width:100%; height:450px; background:#ffffff; display:none; z-index:10;">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10 pdb10">
						<form name="searchForm" action="?subAction=searchItem&otype=<%=otype%>" method="post">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="6"><B><%=L_Top_Search%> <font color="#FFFFFF">(*)</font></B></td>
							</tr>
						</table>
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin-bottom:10px;">
							<col width="100" /><col width="160" /><col width="100" /><col width="100" /><col width="100" />
							<tr>
								<td class="td_l_c title" style="border-top:0;">任务关键词</td>
								<td class="td_r_l" style="border-top:0;"><input name="wTitle" type="text" id="wTitle" class="int" size="20" value="<%=Session("Search_Plugin_Workschedule_wTitle")%>" ></td>
								<td class="td_l_c title" style="border-top:0;">负责人</td>
								<td class="td_r_l" style="border-top:0;"><% If Session("CRM_level") = 9 Then %><% = EasyCrm.UserList(2,"User",""&Session("Search_Plugin_Workschedule_wUserb")&"") %><%else%><% = EasyCrm.UserList(1,"User",""&Session("Search_Plugin_Workschedule_wUserb")&"") %><%end if%></td>
								<td class="td_l_c title" style="border-top:0;">到期时间</td>
								<td class="td_r_l" style="border-top:0;"><input name="TimeBegin" type="text" id="TimeBegin" class="Wdate" value="<%=Session("Search_Plugin_Workschedule_TimeBegin")%>" style="width:130px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" />&nbsp;~&nbsp;<input name="TimeEnd" type="text" id="TimeEnd" class="Wdate" value="<%=Session("Search_Plugin_Workschedule_TimeEnd")%>" style="width:130px;" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" /></td>
							</tr>
							<tr>
								<td class="td_r_l" colspan="6" style="padding:5px 10px;">
									<input type="submit" name="Submit" class="button45" value=" <%=L_Search%> ">　
									<input type="button" name="button" class="button43" value=" <%=L_Clear%> " onClick=window.location.href="?SubAction=killSession" /></td>
								</tr>
						</table>   
						</form>
					</td>
				</tr>
			</table>
		</div>
<%
action = Trim(Request("action"))
Select Case action
Case "Install"
    Call Install()
Case "Add"
    Call infoadd()
Case "save"
    Call infosave()
Case "view"
    Call infoview()
Case "infore"
    Call infore()
Case "inforepf"
    Call inforepf()
Case "Audit"
    Call Audit()
Case "edit"
    Call infoedit()
Case "saveEdit"
    Call infosaveEdit()
Case "Manage"
    Call infoManage()
Case "Managesave"
    Call infoManagesave()
Case "delete"
    Call infodelete()
Case Else
    Call infolist()
End Select
%>

<%
Sub infolist()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10 pdb10">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t">
					<td width="50" class="td_l_c">阅读</td>
					<td width="80" class="td_l_c">编号</td>
					<td width="130" class="td_l_c">期限</td>
					<td width="80" class="td_l_c">分类</td>
					<td width="100" class="td_l_c">星标</td>
					<td class="td_l_c">任务标题</td>
					<td width="80" class="td_l_c">负责人</td>
					<td width="60" class="td_l_c">进度</td>
					<td width="60" class="td_l_c">反馈</td>
					<td width="60" class="td_l_c">发布人</td>
					<%if inStr(Plugin_Workschedule_manage,session("CRM_name"))>0 or Session("CRM_level")=9 then%>
					<td width="100" class="td_l_c">管理</td>
					<%end if%>
				</tr>
				<% = strToPrint %>
			</table>
		</td>
	</tr>
</table>
<div class="fixed_bg">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd ">
			<%if sql<>"" then%><span class="r"><input name="Back" type="button" id="Back" class="button227" value="清空" onClick=window.location.href="?SubAction=killSession"></span><%end if%>
			<%=EasyCrm.pagelist("index.asp", PN,TotalPages,TotalRecords)%>
		</td>
	</tr>
</table>
</div>
<%
end sub
%>
<%
Sub infoadd()
%><style>body{padding-bottom:55px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10 pdb10">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>新增 <font color="#FFF">(*)</font></B></td>
				</tr>
			</table>
			<form name="infoadd" id="infoadd" action="?action=save" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" /><col width="300" /><col width="100" />
				<tr>
					<td class="td_l_c title" style="border-top:0;">分类</td>
					<td class="td_r_l" style="border-top:0;" colspan=3>
							<%
							str = split(""&Plugin_Workschedule_class&"",",")
							for i = 0 to ubound(str)
							response.Write "<input name=""wClass"" type=""radio"" class=""noborder"" value="""&str(i)&"""> "&str(i)&"　"
							next
							response.Write "<input name=""wClass"" type=""radio"" class=""noborder"" value=""草稿""> 草稿　"
							%>
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">任务标题</td>
					<td class="td_r_l"><input name="wTitle" type="text" class="int" id="wTitle" size="40"> <font color="#ff0000">*</font></td>
					<td class="td_l_c title" >要求完成时间</td>
					<td class="td_r_l" ><input name="wCompletiontime" type="text" maxlength="10" id="wCompletiontime" class="Wdate" size="25" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:00'})" value="" /> <font color="#ff0000">*</font></td>
				</tr>
				<tr>
					<td class="td_l_c title">任务星标</td>
					<td class="td_r_l" colspan=3>
						<input name="wStar" type="radio" class="noborder" value="" checked> 无　
						<input name="wStar" type="radio" class="noborder" value="1"> <img src="ico/star.png">　
						<input name="wStar" type="radio" class="noborder" value="2"> <img src="ico/star.png"><img src="ico/star.png">　
						<input name="wStar" type="radio" class="noborder" value="3"> <img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png">　
						<input name="wStar" type="radio" class="noborder" value="4"> <img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png">　
						<input name="wStar" type="radio" class="noborder" value="5"> <img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png">
					
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">负责人</td>
					<td class="td_r_l" colspan=3> <font color="#ff0000">主要：</font><% = EasyCrm.UserList(2,"wUserb","") %> <font color="#ff0000">*</font>　<font color="#ff0000">协助：</font>
					<%
						Set rsm = Server.CreateObject("ADODB.Recordset")
						rsm.Open "Select * From [user] ",conn,1,1
						Do While Not rsm.BOF And Not rsm.EOF
					%>
					<input type="checkbox" name="wUsers" value="<%=rsm("uName")%>"> <%=rsm("uName")%>　
					<%
						rsm.MoveNext
						Loop
						rsm.Close
						Set rsm = Nothing
					%></td>
				</tr>
				<tr>
					<td class="td_l_c title">内容</td>
					<td class="td_r_l" colspan=3 style="padding:10px;"><textarea name="wContent" id="wContent" style="width:80%;height:150px;"></textarea>
					</td>
				</tr>
			</table>   
		</td>
	</tr>
</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="wUser" type="hidden" id="wUser" value="<%=Session("CRM_name")%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="<%=L_Back%>" onClick="history.back();">
		</td>
	</tr>
</table>
</div>
			</form>
<script type="text/javascript"> 
 new tqEditor('wContent',{toolbar: 'crm',
imageUploadUrl: '<%=skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
</script>
<%
End Sub

Sub infosave()
    Dim wClass,wStar,wTitle,wContent,wUserb,wUsers,wUser,wCompletiontime
	wClass = Trim(Request("wClass"))
	wStar = Trim(Request("wStar"))
	wTitle = Trim(Request("wTitle"))
	wContent = Trim(Request("wContent"))
	wUserb = Trim(Request("wUserb"))
	wUsers = Trim(Request("wUsers"))
	wUser = Trim(Request("wUser"))
	wCompletiontime = Trim(Request("wCompletiontime"))
	if wClass="草稿" then
	wUserb = ""
	wCompletiontime = "2099-12-31 23:59:59"
	else
	if wUserb = "" or wCompletiontime = "" then
	Response.Write("<script>alert(""完成时间或主要负责人不能为空!"");history.back(1);</script>")
	end if
	end if
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select Top 1 * From [Plugin_Workschedule]",conn,3,2
	rs.AddNew
	rs("wClass") = wClass
	rs("wStar") = wStar
	rs("wTitle") = wTitle
	rs("wContent") = wContent
	rs("wState") = ecarr(1)
	rs("wUserb") = wUserb
	rs("wUsers") = wUsers
	rs("wUser") = wUser
	rs("wCompletiontime") = wCompletiontime
	rs("wTime") = now()
	rs("wMsg") = 0
	rs("wSH") = 0
	rs("wYd") = 0
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("index.asp")
End Sub

Sub infoview()
    Dim wId
	wId = CLng(ABS(Request("wId")))
	Dim wClass,wTitle,wContent,wUserb,wUsers,wState,wSH,wCompletiontime
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Plugin_Workschedule] Where wId = " & wId,conn,1,1
	If rs.RecordCount <> 1 Then Response.Write "<script>alert(""不存在"");</script>"
	wClass = rs("wClass")
	wTitle = rs("wTitle")
	wContent = rs("wContent")
	wUserb = rs("wUserb")
	wUsers = rs("wUsers")
	wState = rs("wState")
	wUser = rs("wUser")
	wSH = rs("wSH")
	wCompletiontime = rs("wCompletiontime")
	rs.Close
	Set rs = Nothing
	
	if Session("CRM_name") = ""&wUserb&"" then
	conn.execute ("UPDATE Plugin_Workschedule SET wYd=1 Where wID ="&wID&" ")
	end if
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10 ">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>查看 <font color="#FFF">(*)</font></B></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" /><col width="350" /><col width="100" />
				<tr>
					<td class="td_l_c title" style="border-top:0;">分类</td>
					<td class="td_r_l" style="border-top:0;"><%=wClass%></td>
					<td class="td_l_c title" style="border-top:0;">要求完成时间</td>
					<td class="td_r_l" style="border-top:0;"><%=wCompletiontime%></td>
				</tr>
				<tr>
					<td class="td_l_c title">任务标题</td>
					<td class="td_r_l"><%=wTitle%></td>
					<td class="td_l_c title">当前进度</td>
					<td class="td_r_l"><%=wState%></td>
				</tr>
				<tr>
					<td class="td_l_c title">负责人</td>
					<td class="td_r_l" colspan=3> <font color="#ff0000">主要：</font><%=wUserb%>　<font color="#ff0000">协助：</font><%=wUsers%></td>
				</tr>
				<tr>
					<td class="td_l_c title">内容</td>
					<td class="td_r_l" colspan="3"><%=wContent%>
					</td>
				</tr>
				<% if wSH = "1" and wState = ""&ecarr(0)&"" then%>
				<tr>
					<td class="td_l_c title">审核</td>
					<td class="td_r_l" colspan="3"> <input type="button" class="button46" value="通过" onClick="window.location.href='?action=Audit&wID=<%=wID%>&SHtype=通过'" /> <input type="button" class="button47" value="拒绝" onClick="window.location.href='?action=Audit&wID=<%=wID%>&SHtype=拒绝'" /></td>
				</tr>
				<%end if%>
			</table> 
			<%
			Dim rsre
			Set rsre = Server.CreateObject("ADODB.Recordset")
			rsre.Open "Select * From Plugin_Workschedule_re Where wId = " & wId&" Order By rId Desc",conn,1,1
			Do While Not rsre.BOF And Not rsre.EOF
			%>
			<style>.titletips { background-color:#ebfffc;border:1px solid #00c8ab;}
			</style>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin-top:10px;">
				<col width="100" />
				<tr>
					<td class="td_l_c titletips">进度</td>
					<td class="td_r_l titletips"><%=rsre("rState")%> &nbsp;&nbsp; ( Time : <%=rsre("rTime")%> )</td>
				</tr>
				<tr>
					<td class="td_l_c title">内容</td>
					<td class="td_r_l"><%=rsre("rContent")%></td>
				</tr>
				<%if rsre("rRE")<>"" then %>
				<tr>
					<td class="td_l_c title">领导批复</td>
					<td class="td_r_l"><%=rsre("rRE")%></td>
				</tr>
				<%elseif Session("CRM_level") = 9 and rRE="" then %>
				<form name="inforepf" id="inforepf" action="?action=inforepf" method="post">
				<tr>
					<td class="td_l_c title">领导批复</td>
					<td class="td_r_l" style="padding:10px;"><textarea name="rRE" id="rRE" style="width:400px;height:100px;"></textarea></td>
				</tr>
				<tr>
					<td class="td_l_l " colspan=2 style="padding:5px 10px;">
						<input name="rId" type="hidden" id="rId" value="<%=rsre("rID")%>">
						<input name="wID" type="hidden" id="wID" value="<%=wID%>">
						<input type="submit" name="Submit" class="button45" value="保存">　
						
						<input name="Back" type="button" id="Back" class="button43" value="<%=L_Back%>" onClick="history.back();">

					</td>
				</tr>
				</form>
				<%end if%>
			</table> 
			<%
				rsre.MoveNext
			Loop
			rsre.Close
			Set rsre = Nothing
			%>
			<% if wSH <> "2" And Session("CRM_name") = ""&wUserb&""  then%>
<script language="JavaScript">
function CheckInput(f){for(i=0;i<f.rState.length;i++)if(f.rState[i].checked)return true;alert('进度不能为空！')return false}
</script>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin-top:10px;">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>负责人反馈 <font color="#color:#CC0000">(*)</font></B></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin-bottom:10px;">
			<form name="infore" id="infore" action="?action=infore" method="post" onSubmit="return CheckInput(this);">
				<col width="100" />
				<tr>
					<td class="td_l_c title" style="border-top:0;">进度</td>
					<td class="td_r_l" style="border-top:0;">
						<%
							str = split(""&Plugin_Workschedule_State&"",",")
							for i = 0 to ubound(str)
							response.Write "<input name=""rState"" type=""radio"" class=""noborder"" value="""&str(i)&"""> "&str(i)&"　"
							next
						%>
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">内容</td>
					<td class="td_r_l" style="padding:10px;"><textarea name="rContent" id="rContent" style="width:80%;height:150px;"><%=rContent%></textarea>
					</td>
				</tr>
				<tr>
					<td class="td_l_l " colspan=2 style="padding:5px 10px;">
						<% if wSH <> "2" And Session("CRM_name") = ""&wUserb&""  then%>
						<input name="wId" type="hidden" id="wId" value="<%=wId%>">
						<input name="rUser" type="hidden" id="rUser" value="<%=Session("CRM_name")%>">
						<input type="submit" name="Submit" class="button45" value="保存">　
						<%end if%>
						<input name="Back" type="button" id="Back" class="button43" value="<%=L_Back%>" onClick="history.back();">

					</td>
				</tr>
			</table> 
			<%end if%><BR>
		</td>
	</tr>
</table>
			</form> 
<script type="text/javascript"> 
 new tqEditor('rContent',{toolbar: 'crm',
imageUploadUrl: '<%=skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
</script>
<%
End Sub

Sub infore()
	wssarr = split(""&Plugin_Workschedule_State&"",",")
    Dim wID,rState,rContent,rUser
	wID = Trim(Request("wID"))
	rState = Trim(Request("rState"))
	rContent = Trim(Request("rContent"))
	rUser = Trim(Request("rUser"))
	if wssarr(0)=rState then
	conn.execute ("UPDATE Plugin_Workschedule SET wState='"&rState&"',wMsg=1,wSH=1 Where wID ="&wID&" ")
	else
	conn.execute ("UPDATE Plugin_Workschedule SET wState='"&rState&"',wMsg=1,wSH=0 Where wID ="&wID&" ")
	end if
	conn.execute ("insert into Plugin_Workschedule_re(wID,rState,rContent,rUser,rTime) values('"&wID&"','"&rState&"','"&rContent&"','"&Session("CRM_name")&"','"&now()&"')")
	Response.Redirect("index.asp")
End Sub

Sub inforepf()
    Dim rID,wID,rRe
	rID = Trim(Request("rID"))
	wID = Trim(Request("wID"))
	rRe = Trim(Request("rRe"))
	if rRe<>"" then
	conn.execute ("UPDATE Plugin_Workschedule_re SET rRe='"&rRe&"' Where rID ="&rID&" ")
	conn.execute ("UPDATE Plugin_Workschedule SET wMsg=0 Where wID ="&wID&" ")
	end if
	Response.Redirect("index.asp?action=view&wID="&wID&"")
End Sub

Sub Audit()
    Dim wID
	wID = Trim(Request("wID"))
	SHtype = Trim(Request("SHtype"))
	if SHtype="通过" then
	conn.execute ("UPDATE Plugin_Workschedule SET wSH = '2' Where wID ="&wID&" ")
	else
	conn.execute ("UPDATE Plugin_Workschedule SET wSH = '0', wYd = '2' , wState = '" & ecarr(1) & "' Where wID ="&wID&" ")
	end if
	Response.Redirect("?action=Item3&otype=Item3")
End Sub

Sub infoedit()
    Dim wId
	wId = CLng(ABS(Request("wId")))
	Dim wClass,wStar,wTitle,wContent,wUserb,wUsers,wState,wSH,wCompletiontime
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Plugin_Workschedule] Where wId = " & wId,conn,1,1
	If rs.RecordCount <> 1 Then Response.Write "<script>alert(""不存在"");</script>"
	wClass = rs("wClass")
	wStar = rs("wStar")
	if wStar="" then wStar=0
	wTitle = rs("wTitle")
	wContent = rs("wContent")
	wUserb = rs("wUserb")
	wUsers = rs("wUsers")
	wState = rs("wState")
	wUser = rs("wUser")
	wSH = rs("wSH")
	wCompletiontime = rs("wCompletiontime")
	rs.Close
	Set rs = Nothing
%><style>body{padding-bottom:55px;}</style>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10 pdb10">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>编辑 <font color="#FFF">(*)</font></B></td>
				</tr>
			</table>
			<form name="infoEdit" id="infoEdit" action="?action=saveEdit&PN=<%=PNN%>" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" /><col width="350" /><col width="100" />
				<tr>
					<td class="td_l_c title" style="border-top:0;">分类</td>
					<td class="td_r_l" style="border-top:0;">
							<%
							str = split(""&Plugin_Workschedule_class&"",",")
							for i = 0 to ubound(str)
							if wClass = str(i) then
							response.Write "<input name=""wClass"" type=""radio"" class=""noborder"" value="""&str(i)&""" checked> "&str(i)&"　"
							else
							response.Write "<input name=""wClass"" type=""radio"" class=""noborder"" value="""&str(i)&""" > "&str(i)&"　"
							end if
							next
							if wClass="草稿" then
							response.Write "<input name=""wClass"" type=""radio"" class=""noborder"" value=""草稿"" checked> 草稿"
							else
							response.Write "<input name=""wClass"" type=""radio"" class=""noborder"" value=""草稿"" > 草稿"
							end if
							%>
					</td>
					<td class="td_l_c title" style="border-top:0;">要求完成时间</td>
					<td class="td_r_l" style="border-top:0;"><input name="wCompletiontime" type="text" maxlength="10" id="wCompletiontime" class="Wdate" size="20" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:00'})" value="<%=wCompletiontime%>" /> <font color="#ff0000">*</font></td>
				</tr>
				<tr>
					<td class="td_l_c title">任务标题</td>
					<td class="td_r_l"><input name="wTitle" type="text" class="int" id="wTitle" size="40" value="<%=wTitle%>" > <font color="#ff0000">*</font></td>
					<td class="td_l_c title">任务星标</td>
					<td class="td_r_l">
						<input name="wStar" type="radio" class="noborder" value="" <%if wStar=0 then%>checked<%end if%>> 无　
						<input name="wStar" type="radio" class="noborder" value="1" <%if wStar=1 then%>checked<%end if%>> <img src="ico/star.png">　
						<input name="wStar" type="radio" class="noborder" value="2" <%if wStar=2 then%>checked<%end if%>> <img src="ico/star.png"><img src="ico/star.png">　
						<input name="wStar" type="radio" class="noborder" value="3" <%if wStar=3 then%>checked<%end if%>> <img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png">　
						<input name="wStar" type="radio" class="noborder" value="4" <%if wStar=4 then%>checked<%end if%>> <img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png">　
						<input name="wStar" type="radio" class="noborder" value="5" <%if wStar=5 then%>checked<%end if%>> <img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png"><img src="ico/star.png">
					
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">负责人</td>
					<td class="td_r_l" colspan=3> <font color="#ff0000">主要：</font><% = EasyCrm.UserList(2,"wUserb",""&wUserb&"") %> <font color="#ff0000">*</font>　<font color="#ff0000">协助：</font>					
					<%
						Set rsm = Server.CreateObject("ADODB.Recordset")
						rsm.Open "Select * From [user] ",conn,1,1
						Do While Not rsm.BOF And Not rsm.EOF
					%>
					<input type="checkbox" name="wUsers" value="<%=rsm("uName")%>" <%if inStr(wUsers,rsm("uName"))>0 then%>checked<%end if%>> <%=rsm("uName")%>　
					<%
						rsm.MoveNext
						Loop
						rsm.Close
						Set rsm = Nothing
					%></td>
				</tr>
				<tr>
					<td class="td_l_c title">内容</td>
					<td class="td_r_l" colspan="3" style="padding:10px;"><textarea name="wContent" id="wContent" style="width:100%;height:200px;"><%=wContent%></textarea>
					</td>
				</tr>
			</table> 
		</td>
	</tr>
</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input name="wId" type="hidden" id="wId" value="<%=wId%>">
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="<%=L_Back%>" onClick="history.back();">
		</td>
	</tr>
</table>
</div>
			</form>
<script type="text/javascript"> 
 new tqEditor('wContent',{toolbar: 'crm',
imageUploadUrl: '<%=skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
</script>
<%
End Sub

Sub infosaveEdit()   
    Dim wId
	wId = CLng(ABS(Request("wId")))
	PN = Request("PN")
	If Not IsNumeric(wId) Or wId <= 0 Then Response.Write "<script>alert(""不存在"");</script>"
	Dim wClass,wStar,wTitle,wContent,wUserb,wUsers,wUser,wCompletiontime
	wClass = Trim(Request("wClass"))
	wStar = Trim(Request("wStar"))
	wTitle = Trim(Request("wTitle"))
	wContent = Trim(Request("wContent"))
	wUserb = Trim(Request("wUserb"))
	wUsers = Trim(Request("wUsers"))
	wCompletiontime = Trim(Request("wCompletiontime"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select Top 1 * From [Plugin_Workschedule] Where wId = " & wId,conn,3,2
	rs("wClass") = wClass
	rs("wStar") = wStar
	rs("wTitle") = wTitle
	rs("wContent") = wContent
	rs("wUserb") = wUserb
	rs("wUsers") = wUsers
	rs("wCompletiontime") = wCompletiontime
	rs("wTime") = now()
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("index.asp?PN="&PN&"")
End Sub

Sub infoManage()
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdl10 pdr10 pdt10 pdb10">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>高级配置 <font color="#color:#CC0000">(*)</font></B></td>
				</tr>
			</table>
			<form name="Managesave" action="?action=Managesave" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" />
				<tr>
					<td class="td_l_c title" style="border-top:0;">工作分类</td>
					<td class="td_r_l" style="border-top:0;">
						<input name="Plugin_Workschedule_class" type="text" class="int" id="Plugin_Workschedule_class" size="40" value="<%=Plugin_Workschedule_class%>"> <span class="info_help help01">不同分类之间用半角逗号分割，结尾不含逗号，下同。</span>
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">当前进度</td>
					<td class="td_r_l">
						<input name="Plugin_Workschedule_State" type="text" class="int" id="Plugin_Workschedule_State" size="60" value="<%=Plugin_Workschedule_State%>"> <span class="info_help help01">第一节点为进度终止，作为判断是否需要审核的条件。</span>
					</td>
				</tr>
				<tr>
					<td class="td_l_c title">发布权限</td>
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
								<input type="checkbox" name="Plugin_Workschedule_manage" value="<%=rsm("uName")%>" <%if inStr(Plugin_Workschedule_manage,rsm("uName"))>0 then%>checked<%end if%>> <%=rsm("uName")%>　
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
					<input type="submit" name="Submit" class="button45" value=" <%=L_Edit%> ">　
					<input name="Back" type="button" id="Back" class="button43" value=" <%=L_Back%> " onClick="history.back();">
					</td>
				</tr>
			</table>   
			</form>
		</td>
	</tr>
</table>

<%
End Sub

Sub infoManagesave()
	Plugin_Workschedule_class = replace(Trim(Request.Form("Plugin_Workschedule_class")),CHR(34),"'")
	Plugin_Workschedule_State = replace(Trim(Request.Form("Plugin_Workschedule_State")),CHR(34),"'")
	Plugin_Workschedule_manage = replace(Trim(Request.Form("Plugin_Workschedule_manage")),CHR(34),"'")
	Dim TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "Dim Plugin_Workschedule_class,Plugin_Workschedule_company,Plugin_Workschedule_manage" & VbCrLf
	
	TempStr = TempStr & "'详细配置" & VbCrLf
	TempStr = TempStr & "Plugin_Workschedule_class="& Chr(34) & Plugin_Workschedule_class & Chr(34) &" '工作分类" & VbCrLf
	TempStr = TempStr & "Plugin_Workschedule_State="& Chr(34) & Plugin_Workschedule_State & Chr(34) &" '当前进度" & VbCrLf
	TempStr = TempStr & "Plugin_Workschedule_manage="& Chr(34) & Plugin_Workschedule_manage & Chr(34) &" '权限" & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf
	ADODB_SaveToFile TempStr,"Config.asp"
	Response.Write("<script>alert(""修改成功！"");</script>")
	Response.Write "<script>location.href='?action=List&otype=Main';</script>"
End Sub

Sub infodelete()
    Dim wId
	wId = CLng(ABS(Request("wId")))
	PN = Request("PN")
	If Not IsNumeric(wId) Or wId <= 0 Then Response.Write "<script>alert(""不存在"");</script>"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [Plugin_Workschedule] Where wId = " & wId,conn,3,2
	If rs.RecordCount <> 1 Then Response.Write "<script>alert(""不存在"");</script>"
	wId = rs("wId")
	rs.Delete
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("index.asp?PN="&PN&"")
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
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请使用FTP等功能，将<font color=Red >data/config.asp</font>文件内容替换成框中内容"
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
<script src="../../data/calendar/WdatePicker.js"></script>
</body>
</html><% Set EasyCrm = nothing %>
