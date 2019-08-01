<!--#include file="../../Data/Conn.asp"--><!--#include file="../../UpLoad/UpLoad.asp"--><!--#include file="config.asp"--><!--#include file="../../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>
<body>
<%
action = Trim(Request("action"))
sType = Trim(Request("sType"))
tipinfo = Trim(Request("tipinfo"))
Id = Trim(Request("Id"))

Select Case action
Case "Add"
    Call Add()
Case "SaveAdd"
    Call SaveAdd()
Case "ReasonView"
    Call ReasonView()
Case "ContentView"
    Call ContentView()
Case "InfoEdit"
    Call InfoEdit()
Case "SaveEdit"
    Call SaveEdit()
Case "Audit"
    Call Audit()
Case "SaveAudit"
    Call SaveAudit()
Case "ViewInfo"
    Call ViewInfo()
End Select

Sub Add()
%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.getElementById('TimeBegin').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '外出时间<%=alert04%>'});document.getElementById('TimeBegin').focus();return false;}
			if(document.getElementById('TimeEnd').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '返回时间<%=alert04%>'});document.getElementById('TimeEnd').focus();return false;}
		}
		-->
		</script>
	<form name="Save" action="?action=SaveAdd" method="post" onSubmit="return CheckInput();">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<col width="100"><col width="200"><col width="100">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="4"><B>员工外出申请单</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">外出原因</td>
						<td class="td_r_l" colspan=3> 
						<%
							ReasonStr = split(""&Plugin_GoOut_Reason&"",",")
							for i = 0 to ubound(ReasonStr)
							response.Write "<input name=""gReason"" type=""radio"" class=""noborder"" value="""&ReasonStr(i)&"""> "&ReasonStr(i)&"　"
							next
						%>
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" >外出时间</td>
						<td class="td_r_l" colspan=3 ><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate int" size="15" onFocus="WdatePicker()"  /> ～ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate int" size="15" onFocus="WdatePicker()" /> </td>
					</tr>
					<tr>
						<td class="td_l_r title" >拜访人</td>
						<td class="td_r_l" ><input name="gLinkman" type="text" id="gLinkman" class="int" size="15" /></td>
						<td class="td_l_r title" >电话</td>
						<td class="td_r_l" ><input name="gTel" type="text" id="gTel" class="int" size="25" /></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" style="padding:10px;" colspan=3> <textarea name="gContent" id="gContent" style="width:99%;height:140px;"></textarea></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
					<input name="gUser" type="hidden" value="<%=Session("CRM_name")%>">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('gContent',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
<%
End Sub

Sub SaveAdd()
	gSTime = Request.Form("TimeBegin")
	gETime = Request.Form("TimeEnd")
	gReason = Request.Form("gReason")
	gLinkman = Request.Form("gLinkman")
	gTel = Request.Form("gTel")
	gContent = Request.Form("gContent")
	gUser = Request.Form("gUser")
	gTime = Request.Form("gTime")
	conn.execute("insert into [Plugin_GoOut] (gSTime,gETime,gReason,gLinkman,gTel,gContent,gUser,gTime,gState) values('"&gSTime&"','"&gETime&"','"&gReason&"','"&gLinkman&"','"&gTel&"','"&gContent&"','"&gUser&"','"&Now()&"',0)")
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub

Sub InfoEdit()
%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.getElementById('TimeBegin').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '外出时间<%=alert04%>'});document.getElementById('TimeBegin').focus();return false;}
			if(document.getElementById('TimeEnd').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '返回时间<%=alert04%>'});document.getElementById('TimeEnd').focus();return false;}
		}
		-->
		</script>
	<form name="Save" action="?action=SaveEdit" method="post" onSubmit="return CheckInput();">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<col width="100"><col width="200"><col width="100">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="4"><B>员工外出申请单</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">外出原因</td>
						<td class="td_r_l" colspan=3> 
						<%
							ReasonStr = split(""&Plugin_GoOut_Reason&"",",")
							for i = 0 to ubound(ReasonStr)
							if EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gReason") = ReasonStr(i) then
							response.Write "<input name=""gReason"" type=""radio"" class=""noborder"" value="""&ReasonStr(i)&""" checked> "&ReasonStr(i)&"　"
							else
							response.Write "<input name=""gReason"" type=""radio"" class=""noborder"" value="""&ReasonStr(i)&""" > "&ReasonStr(i)&"　"
							end if
							next
						%>
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" >外出时间</td>
						<td class="td_r_l" colspan=3 ><input name="TimeBegin" type="text" maxlength="10" id="TimeBegin" class="Wdate int" size="15" value="<%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gSTime")%>" onFocus="WdatePicker()"  /> ～ <input name="TimeEnd" type="text" maxlength="10" id="TimeEnd" class="Wdate int" size="15" value="<%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gETime")%>" onFocus="WdatePicker()" /> </td>
					</tr>
					<tr>
						<td class="td_l_r title" >拜访人</td>
						<td class="td_r_l" ><input name="gLinkman" type="text" id="gLinkman" class="int" size="15" value="<%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gLinkman")%>" /></td>
						<td class="td_l_r title" >电话</td>
						<td class="td_r_l" ><input name="gTel" type="text" id="gTel" class="int" size="25" value="<%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gTel")%>" /></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" style="padding:10px;" colspan=3> <textarea name="gContent" id="gContent" style="width:99%;height:140px;"><%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gContent")%></textarea></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
				<span class="Bottom_pd r fontnobold">注：已审批的信息不可修改</span>
					<input name="id" type="hidden" value="<%=id%>">
					<%if EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gState")=1 then%>
					<input type="button" class="button45" value="保存">　
					<%else%>
					<input type="submit" name="Submit" class="button45" value="保存">　
					<%end if%>
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('gContent',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
<%
End Sub

Sub SaveEdit()
	id = Request("id")
	gSTime = Request.Form("TimeBegin")
	gETime = Request.Form("TimeEnd")
	gReason = Request.Form("gReason")
	gLinkman = Request.Form("gLinkman")
	gTel = Request.Form("gTel")
	gContent = Request.Form("gContent")
	conn.execute "UPDATE Plugin_GoOut SET gSTime='"&gSTime&"',gETime='"&gETime&"',gReason='"&gReason&"',gLinkman='"&gLinkman&"',gTel='"&gTel&"',gContent='"&gContent&"' Where id="&id
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub

Sub Audit()
%>
	<style>body{padding:0 0 55px 0;}</style>
	<form name="Save" action="?action=SaveAudit" method="post" onSubmit="return CheckInput();">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<col width="100"><col width="200"><col width="100">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="4"><B>员工外出申请单</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">外出原因</td>
						<td class="td_r_l" colspan=3> <%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gReason")%> </td>
					</tr>
					<tr>
						<td class="td_l_r title" >外出时间</td>
						<td class="td_r_l" colspan=3 ><%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gSTime")%> ～ <%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gETime")%> </td>
					</tr>
					<tr>
						<td class="td_l_r title" >拜访人</td>
						<td class="td_r_l" ><%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gLinkman")%></td>
						<td class="td_l_r title" >电话</td>
						<td class="td_r_l" ><%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gTel")%></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" colspan=3><%=EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gContent")%></td>
					</tr>
				</table>
			</td> 
		</tr>
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<col width="100"><col width="200"><col width="100">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="4"><B>领导审核</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">审核状态</td>
						<td class="td_r_l" colspan=3> 
							<input name="gState" type="radio" class="noborder" value="1" <%if EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gState")=1 then%>checked<%end if%>> 通过 　
							<input name="gState" type="radio" class="noborder" value="2" <%if EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gState")=2 then%>checked<%end if%> > 不通过 　
							<input name="gState" type="radio" class="noborder" value="0" <%if EasyCrm.getNewItem("Plugin_GoOut","Id",""&id&"","gState")=0 then%>checked<%end if%>> 暂不处理
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" >审核人</td>
						<td class="td_r_l" colspan=3 ><% = Session("CRM_name") %></td>
					</tr>
					<tr>
						<td class="td_l_r title" >审核时间</td>
						<td class="td_r_l" colspan=3 ><%=Now()%></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
					<input name="id" type="hidden" value="<%=id%>">
					<input name="gAudit" type="hidden" value="<%=Session("CRM_name")%>">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
<%
End Sub

Sub SaveAudit()
	id = Request("id")
	gState = Request.Form("gState")
	gAudit = Request.Form("gAudit")
	conn.execute "UPDATE [Plugin_GoOut] SET gState='"&gState&"',gAudit='"&gAudit&"',gAuditTime='"&now()&"' Where id="&id
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub

%>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>