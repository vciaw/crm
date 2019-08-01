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
Case "Audit"
    Call Audit()
Case "SaveAdd"
    Call SaveAdd()
Case "InfoEdit"
    Call InfoEdit()
Case "SaveEdit"
    Call SaveEdit()
Case "AddBank"
    Call AddBank()
Case "SaveAddBank"
    Call SaveAddBank()
Case "InfoEditBank"
    Call InfoEditBank()
Case "SaveEditBank"
    Call SaveEditBank()
Case "AddOutin"
    Call AddOutin()
Case "SaveAddOutin"
    Call SaveAddOutin()
Case "InfoEditOutin"
    Call InfoEditOutin()
Case "SaveEditOutin"
    Call SaveEditOutin()
End Select

Sub Add()
%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.all.fUser.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '姓名<%=alert04%>'});document.all.fUser.focus();return false;}
			if(document.all.fMoney.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '金额<%=alert04%>'});document.all.fMoney.focus();return false;}
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
						<td class="td_l_l" COLSPAN="4"><B>新增账目</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" >日期</td>
						<td class="td_r_l"><input name="fTime" type="text" maxlength="10" id="fTime" class="Wdate int" size="15" onFocus="WdatePicker()" value="<%=EasyCrm.FormatDate(Now(),2)%>" /> </td>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 姓名</td>
						<td class="td_r_l"><% = EasyCrm.UserList(2,"fUser","") %></td>
					</tr>
					<tr>
						<td class="td_l_r title" >对方科目</td>
						<td class="td_r_l">
							<select name="fSubjects" class="int" style="width:150px;">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_Finance_Subjects&"",",")
								for i = 0 to ubound(str)
								response.Write "<option value="&str(i)&">"&str(i)&"</option>"
								next
								%>
							</select>
						</td>
						<td class="td_l_r title" >类型</td>
						<td class="td_r_l">
							<select name="fClass" class="int" style="width:150px;">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_Finance_Class&"",",")
								for i = 0 to ubound(str)
								response.Write "<option value="&str(i)&">"&str(i)&"</option>"
								next
								%>
							</select>
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" >票号</td>
						<td class="td_r_l" ><input name="fInvoice" type="text" id="fInvoice" class="int" size="20" /></td>
						<td class="td_l_r title" >对应项目</td>
						<td class="td_r_l">
							<select name="fProject" class="int" style="width:150px;">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_Finance_Project&"",",")
								for i = 0 to ubound(str)
								response.Write "<option value="&str(i)&">"&str(i)&"</option>"
								next
								%>
							</select>
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" >摘要</td>
						<td class="td_r_l" ><input name="fDigest" type="text" id="fDigest" class="int" size="30" /></td>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 金额</td>
						<td class="td_r_l" >
							<input name="fType" type="radio" class="noborder" value="fDebit" > 借+　
							<input name="fType" type="radio" class="noborder" value="fCredit" checked> 贷-　
							<b style="color:red;">￥</b><input name="fMoney" type="text" id="fMoney" class="int" size="6" onkeyup='this.value=this.value.replace(/\D/gi,"")' /> RMB</td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" style="padding:10px;" colspan=3> <textarea name="fRemark" id="fRemark" style="width:99%;height:100px;"></textarea></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd ">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('fRemark',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
<%
End Sub

Sub SaveAdd()

	fTime = Request.Form("fTime")
	fUser = Request.Form("fUser")
	fGroup = EasyCrm.getNewItem("user","uName","'"&Request.Form("fUser")&"'","uGroup")
	fSubjects = Request.Form("fSubjects")
	fClass = Request.Form("fClass")
	fInvoice = Request.Form("fInvoice")
	fProject = Request.Form("fProject")
	fDigest = Request.Form("fDigest")
	fAudit = Request.Form("fAudit")
	if fAudit ="" then fAudit = "未审核"
	
	if Request.Form("fType") = "fDebit" then
	fDebit = Request.Form("fMoney")
	fCredit = 0
	elseif Request.Form("fType") = "fCredit" then
	fDebit = 0
	fCredit = Request.Form("fMoney")
	end if
	
	fRemark = Request.Form("fRemark")
	
	conn.execute("insert into [Plugin_Finance] (fTime,fUser,fGroup,fSubjects,fClass,fInvoice,fProject,fDigest,fDebit,fCredit,fRemark,fAudit) values('"&fTime&"','"&fUser&"','"&fGroup&"','"&fSubjects&"','"&fClass&"','"&fInvoice&"','"&fProject&"','"&fDigest&"','"&fDebit&"','"&fCredit&"','"&fRemark&"','"&fAudit&"')")
	
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
			if(document.all.fUser.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '姓名<%=alert04%>'});document.all.fUser.focus();return false;}
			if(document.all.fMoney.value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '金额<%=alert04%>'});document.all.fMoney.focus();return false;}
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
						<td class="td_l_l" COLSPAN="4"><B>编辑账目</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" >日期</td>
						<td class="td_r_l"><input name="fTime" type="text" maxlength="10" id="fTime" class="Wdate int" size="15" onFocus="WdatePicker()" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Plugin_Finance","Id",""&id&"","fTime"),2)%>" /> </td>
						<td class="td_l_r title" >姓名</td>
						<td class="td_r_l"><% = EasyCrm.UserList(2,"fUser",""&EasyCrm.getNewItem("Plugin_Finance","Id",""&id&"","fUser")&"") %></td>
					</tr>
					<tr>
						<td class="td_l_r title" >对方科目</td>
						<td class="td_r_l">
							<select name="fSubjects" class="int" style="width:150px;">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_Finance_Subjects&"",",")
								for i = 0 to ubound(str)
								if EasyCrm.getNewItem("Plugin_Finance","Id",""&id&"","fSubjects") = str(i) then
								response.Write "<option value="&str(i)&" selected>"&str(i)&"</option>"
								else
								response.Write "<option value="&str(i)&">"&str(i)&"</option>"
								end if
								next
								%>
							</select>
						</td>
						<td class="td_l_r title" >类型</td>
						<td class="td_r_l">
							<select name="fClass" class="int" style="width:150px;">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_Finance_Class&"",",")
								for i = 0 to ubound(str)
								if EasyCrm.getNewItem("Plugin_Finance","Id",""&id&"","fClass") = str(i) then
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
						<td class="td_l_r title" >票号</td>
						<td class="td_r_l" ><input name="fInvoice" type="text" id="fInvoice" class="int" size="20" value="<%=EasyCrm.getNewItem("Plugin_Finance","Id",""&id&"","fInvoice")%>" /></td>
						<td class="td_l_r title" >对应项目</td>
						<td class="td_r_l">
							<select name="fProject" class="int" style="width:150px;">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_Finance_Project&"",",")
								for i = 0 to ubound(str)
								if EasyCrm.getNewItem("Plugin_Finance","Id",""&id&"","fProject") = str(i) then
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
						<td class="td_l_r title" >摘要</td>
						<td class="td_r_l" ><input name="fDigest" type="text" id="fDigest" class="int" size="30" value="<%=EasyCrm.getNewItem("Plugin_Finance","Id",""&id&"","fDigest")%>" /></td>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 金额</td>
						<td class="td_r_l" >
							<%
							fDebit = EasyCrm.getNewItem("Plugin_Finance","Id",""&id&"","fDebit")
							fCredit = EasyCrm.getNewItem("Plugin_Finance","Id",""&id&"","fCredit")
							fMoney = fDebit + fCredit
							%>
							<input name="fType" type="radio" class="noborder" value="fDebit" <%if fDebit>0 then%>checked<%end if%>> 借+　
							<input name="fType" type="radio" class="noborder" value="fCredit" <%if fCredit>0 then%>checked<%end if%>> 贷-　
							<b style="color:red;">￥</b><input name="fMoney" type="text" id="fMoney" class="int" size="6" onkeyup='this.value=this.value.replace(/\D/gi,"")' value="<%=fMoney%>" /> RMB</td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" style="padding:10px;" colspan=3> <textarea name="fRemark" id="fRemark" style="width:99%;height:100px;"><%=EasyCrm.getNewItem("Plugin_Finance","Id",""&id&"","fRemark")%></textarea></td>
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
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('fRemark',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
	
<%
End Sub

Sub SaveEdit()
	id = Request("id")
	
	fTime = Request.Form("fTime")
	fUser = Request.Form("fUser")
	fGroup = EasyCrm.getNewItem("user","uName","'"&Request.Form("fUser")&"'","uGroup")
	fSubjects = Request.Form("fSubjects")
	fClass = Request.Form("fClass")
	fInvoice = Request.Form("fInvoice")
	fProject = Request.Form("fProject")
	fDigest = Request.Form("fDigest")
	
	if Request.Form("fType") = "fDebit" then
	fDebit = Request.Form("fMoney")
	fCredit = 0
	elseif Request.Form("fType") = "fCredit" then
	fDebit = 0
	fCredit = Request.Form("fMoney")
	end if
	
	fRemark = Request.Form("fRemark")
	
	conn.execute "UPDATE Plugin_Finance SET fTime='"&fTime&"',fUser='"&fUser&"',fGroup='"&fGroup&"',fSubjects='"&fSubjects&"',fClass='"&fClass&"',fInvoice='"&fInvoice&"',fProject='"&fProject&"',fDigest='"&fDigest&"',fDebit='"&fDebit&"',fCredit='"&fCredit&"',fRemark='"&fRemark&"' Where id="&id
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub

Sub Audit()
	
	if sType="A" then '填写审核原因
%>	
			<table width="100%" border="0" cellpadding="0" cellspacing="0" >
				<form name="Save" action="?action=Audit&sType=SaveA&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						状态：
						<input name="fAudit" type="radio" value="已审核" <%if EasyCrm.getNewItem("Plugin_Finance","ID",""&ID&"","fAudit") = "已审核" then%> checked<%end if%>> 已审核　 
						<input name="fAudit" type="radio" value="未审核" <%if EasyCrm.getNewItem("Plugin_Finance","ID",""&ID&"","fAudit") = "未审核" then%> checked<%end if%>> 未审核 　
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveA" then '保存
		If Id = "" Then Exit Sub
		fAudit = Request.Form("fAudit")
		conn.execute("update [Plugin_Finance] set fAudit = '"&fAudit&"' where Id = "&ID&" ")
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
	elseif sType="B" then '填写审核原因
%>	
			<table width="100%" border="0" cellpadding="0" cellspacing="0" >
				<form name="Save" action="?action=Audit&sType=SaveB&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						状态：
						<input name="bAudit" type="radio" value="已审核" <%if EasyCrm.getNewItem("Plugin_Finance_Bank","ID",""&ID&"","bAudit") = "已审核" then%> checked<%end if%>> 已审核　 
						<input name="bAudit" type="radio" value="未审核" <%if EasyCrm.getNewItem("Plugin_Finance_Bank","ID",""&ID&"","bAudit") = "未审核" then%> checked<%end if%>> 未审核 　
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	
	elseif sType="SaveB" then '保存
		If Id = "" Then Exit Sub
		bAudit = Request.Form("bAudit")
		conn.execute("update [Plugin_Finance_Bank] set bAudit = '"&bAudit&"' where Id = "&ID&" ")
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
	elseif sType="C" then '填写审核原因
%>	
			<table width="100%" border="0" cellpadding="0" cellspacing="0" >
				<form name="Save" action="?action=Audit&sType=SaveC&Id=<%=ID%>" method="post" onSubmit="return CheckInput();">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						状态：
						<input name="oAudit" type="radio" value="已审核" <%if EasyCrm.getNewItem("Plugin_Finance_Outin","ID",""&ID&"","oAudit") = "已审核" then%> checked<%end if%>> 已审核　 
						<input name="oAudit" type="radio" value="未审核" <%if EasyCrm.getNewItem("Plugin_Finance_Outin","ID",""&ID&"","oAudit") = "未审核" then%> checked<%end if%>> 未审核 　
					</td>
				</tr>
			</table>
<div class="fixed_bg_B">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n Bottom_pd "> 
			<input type="submit" name="Submit" class="button45" value="保存">　
			<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
		</td>
	</tr>
</table>
</div>
				</form>
	<%
	elseif sType="SaveC" then '保存
		If Id = "" Then Exit Sub
		oAudit = Request.Form("oAudit")
		conn.execute("update [Plugin_Finance_Outin] set oAudit = '"&oAudit&"' where Id = "&ID&" ")
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
	end if
		
End Sub

Sub AddBank()
%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.getElementById('bMoney').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '金额<%=alert04%>'});document.getElementById('bMoney').focus();return false;}
		}
		-->
		</script>
	<form name="Save" action="?action=SaveAddBank" method="post" onSubmit="return CheckInput();">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<col width="100"><col width="200"><col width="100">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="4"><B>新增存款</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" >日期</td>
						<td class="td_r_l"><input name="bTime" type="text" maxlength="10" id="bTime" class="Wdate int" size="15" onFocus="WdatePicker()" value="<%=EasyCrm.FormatDate(Now(),2)%>" /> </td>
						<td class="td_l_r title" >类型</td>
						<td class="td_r_l">
							<select name="bClass" class="int" style="width:150px;">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_Finance_Bank_Class&"",",")
								for i = 0 to ubound(str)
								response.Write "<option value="&str(i)&">"&str(i)&"</option>"
								next
								%>
							</select>
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" >银行</td>
						<td class="td_r_l" >
							<select name="bName" class="int" style="width:150px;">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_Finance_Bank_Name&"",",")
								for i = 0 to ubound(str)
								response.Write "<option value="&str(i)&">"&str(i)&"</option>"
								next
								%>
							</select>
						</td>
						<td class="td_l_r title" >开户行</td>
						<td class="td_r_l" ><input name="bOpening" type="text" id="bOpening" class="int" size="30" /></td>
					</tr>
					<tr>
						<td class="td_l_r title" >帐号</td>
						<td class="td_r_l" ><input name="bCard" type="text" id="bCard" class="int" size="30" /></td>
						<td class="td_l_r title" >票号</td>
						<td class="td_r_l" ><input name="bInvoice" type="text" id="bInvoice" class="int" size="30" /></td>
					</tr>
					<tr>
						<td class="td_l_r title" >摘要</td>
						<td class="td_r_l" ><input name="bDigest" type="text" id="bDigest" class="int" size="30" /></td>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 金额</td>
						<td class="td_r_l" >
							<input name="bType" type="radio" class="noborder" value="bDebit" > 借+　
							<input name="bType" type="radio" class="noborder" value="bCredit" checked> 贷-　
							<b style="color:red;">￥</b><input name="bMoney" type="text" id="bMoney" class="int" size="6" onkeyup='this.value=this.value.replace(/\D/gi,"")' /> RMB</td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" style="padding:10px;" colspan=3> <textarea name="bRemark" id="bRemark" style="width:99%;height:100px;"></textarea></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd ">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('bRemark',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
<%
End Sub

Sub SaveAddBank()

	bTime = Request.Form("bTime")
	bClass = Request.Form("bClass")
	bName = Request.Form("bName")
	bOpening = Request.Form("bOpening")
	bCard = Request.Form("bCard")
	bInvoice = Request.Form("bInvoice")
	bDigest = Request.Form("bDigest")
	
	if Request.Form("bType") = "bDebit" then
	bDebit = Request.Form("bMoney")
	bCredit = 0
	elseif Request.Form("bType") = "bCredit" then
	bDebit = 0
	bCredit = Request.Form("bMoney")
	end if
	
	bRemark = Request.Form("bRemark")
	
	conn.execute("insert into [Plugin_Finance_Bank] (bTime,bClass,bName,bOpening,bCard,bInvoice,bDigest,bDebit,bCredit,bRemark) values('"&bTime&"','"&bClass&"','"&bName&"','"&bOpening&"','"&bCard&"','"&bInvoice&"','"&bDigest&"','"&bDebit&"','"&bCredit&"','"&bRemark&"')")
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub

Sub InfoEditBank()
%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.getElementById('bMoney').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '金额<%=alert04%>'});document.getElementById('bMoney').focus();return false;}
		}
		-->
		</script>
	<form name="Save" action="?action=SaveEditBank" method="post" onSubmit="return CheckInput();">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<col width="100"><col width="200"><col width="100">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="4"><B>编辑存款</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" >日期</td>
						<td class="td_r_l"><input name="bTime" type="text" maxlength="10" id="bTime" class="Wdate int" size="15" onFocus="WdatePicker()" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Plugin_Finance_Bank","Id",""&id&"","bTime"),2)%>" /> </td>
						<td class="td_l_r title" >类型</td>
						<td class="td_r_l">
							<select name="bClass" class="int" style="width:150px;">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_Finance_Bank_Class&"",",")
								for i = 0 to ubound(str)
								if EasyCrm.getNewItem("Plugin_Finance_Bank","Id",""&id&"","bClass") = str(i) then
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
						<td class="td_l_r title" >银行</td>
						<td class="td_r_l" >
							<select name="bName" class="int" style="width:150px;">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_Finance_Bank_Name&"",",")
								for i = 0 to ubound(str)
								if EasyCrm.getNewItem("Plugin_Finance_Bank","Id",""&id&"","bName") = str(i) then
								response.Write "<option value="&str(i)&" selected>"&str(i)&"</option>"
								else
								response.Write "<option value="&str(i)&">"&str(i)&"</option>"
								end if
								next
								%>
							</select>
						</td>
						<td class="td_l_r title" >开户行</td>
						<td class="td_r_l" ><input name="bOpening" type="text" id="bOpening" class="int" size="30" value="<%=EasyCrm.getNewItem("Plugin_Finance_Bank","Id",""&id&"","bOpening")%>" /></td>
					</tr>
					<tr>
						<td class="td_l_r title" >帐号</td>
						<td class="td_r_l" ><input name="bCard" type="text" id="bCard" class="int" size="30" value="<%=EasyCrm.getNewItem("Plugin_Finance_Bank","Id",""&id&"","bCard")%>" /></td>
						<td class="td_l_r title" >票号</td>
						<td class="td_r_l" ><input name="bInvoice" type="text" id="bInvoice" class="int" size="30" value="<%=EasyCrm.getNewItem("Plugin_Finance_Bank","Id",""&id&"","bInvoice")%>" /></td>
					</tr>
					<tr>
						<td class="td_l_r title" >摘要</td>
						<td class="td_r_l" ><input name="bDigest" type="text" id="bDigest" class="int" size="30" value="<%=EasyCrm.getNewItem("Plugin_Finance_Bank","Id",""&id&"","bDigest")%>" /></td>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 金额</td>
						<td class="td_r_l" >
							<%
							bDebit = EasyCrm.getNewItem("Plugin_Finance_Bank","Id",""&id&"","bDebit")
							bCredit = EasyCrm.getNewItem("Plugin_Finance_Bank","Id",""&id&"","bCredit")
							bMoney = bDebit + bCredit
							%>
							<input name="bType" type="radio" class="noborder" value="bDebit" <%if bDebit>0 then%>checked<%end if%> > 借+　
							<input name="bType" type="radio" class="noborder" value="bCredit" <%if bCredit>0 then%>checked<%end if%>> 贷-　
							<b style="color:red;">￥</b><input name="bMoney" type="text" id="bMoney" class="int" size="6" onkeyup='this.value=this.value.replace(/\D/gi,"")' value="<%=bMoney%>" /> RMB</td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" style="padding:10px;" colspan=3> <textarea name="bRemark" id="bRemark" style="width:99%;height:100px;"><%=EasyCrm.getNewItem("Plugin_Finance_Bank","Id",""&id&"","bRemark")%></textarea></td>
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
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('bRemark',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
	
<%
End Sub

Sub SaveEditBank()
	id = Request("id")

	bTime = Request.Form("bTime")
	bClass = Request.Form("bClass")
	bName = Request.Form("bName")
	bOpening = Request.Form("bOpening")
	bCard = Request.Form("bCard")
	bInvoice = Request.Form("bInvoice")
	bDigest = Request.Form("bDigest")
	
	if Request.Form("bType") = "bDebit" then
	bDebit = Request.Form("bMoney")
	bCredit = 0
	elseif Request.Form("bType") = "bCredit" then
	bDebit = 0
	bCredit = Request.Form("bMoney")
	end if
	
	bRemark = Request.Form("bRemark")
	
	conn.execute "UPDATE Plugin_Finance_Bank SET bTime='"&bTime&"',bClass='"&bClass&"',bName='"&bName&"',bOpening='"&bOpening&"',bCard='"&bCard&"',bInvoice='"&bInvoice&"',bDigest='"&bDigest&"',bDebit='"&bDebit&"',bCredit='"&bCredit&"',bRemark='"&bRemark&"' Where id="&id
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub


Sub AddOutin()
%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.getElementById('oMoney').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '金额<%=alert04%>'});document.getElementById('oMoney').focus();return false;}
		}
		-->
		</script>
	<form name="Save" action="?action=SaveAddOutin" method="post" onSubmit="return CheckInput();">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<col width="100"><col width="200"><col width="100">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="4"><B>新增记录</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" >日期</td>
						<td class="td_r_l" colspan=3><input name="oTime" type="text" maxlength="10" id="oTime" class="Wdate int" size="15" onFocus="WdatePicker()" value="<%=EasyCrm.FormatDate(Now(),2)%>" /> </td>
					</tr>
					<tr>
						<td class="td_l_r title" >公司名称</td>
						<td class="td_r_l" ><input name="oCompany" type="text" id="oCompany" class="int" size="30" /></td>
						<td class="td_l_r title" >票号</td>
						<td class="td_r_l" ><input name="oInvoice" type="text" id="oInvoice" class="int" size="30" /></td>
					</tr>
					<tr>
						<td class="td_l_r title" >摘要</td>
						<td class="td_r_l"><input name="oDigest" type="text" id="oDigest" class="int" size="30" /></td>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 金额</td>
						<td class="td_r_l">
							<input name="osType" type="radio" class="noborder" value="oDebit" checked > 借+　
							<input name="osType" type="radio" class="noborder" value="oCredit" > 贷-　
							<b style="color:red;">￥</b><input name="oMoney" type="text" id="oMoney" class="int" size="6" onkeyup='this.value=this.value.replace(/\D/gi,"")' /> RMB</td>
					</tr>
					<tr>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 状态</td>
						<td class="td_r_l" colspan=3 >
							<input name="oState" type="radio" class="noborder" value="已完成" checked> 已完成　
							<input name="oState" type="radio" class="noborder" value="未完成" > 未完成　</td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" style="padding:10px;" colspan=3> <textarea name="oRemark" id="oRemark" style="width:99%;height:100px;"></textarea></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd ">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('oRemark',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
<%
End Sub

Sub SaveAddOutin()

	oTime = Request.Form("oTime")
	oCompany = Request.Form("oCompany")
	oInvoice = Request.Form("oInvoice")
	oDigest = Request.Form("oDigest")
	oState = Request.Form("oState")
	
	if Request.Form("osType") = "oDebit" then
	oDebit = Request.Form("oMoney")
	oCredit = 0
	elseif Request.Form("osType") = "oCredit" then
	oDebit = 0
	oCredit = Request.Form("oMoney")
	end if
	
	oRemark = Request.Form("oRemark")
	
	conn.execute("insert into [Plugin_Finance_Outin] (oTime,oCompany,oInvoice,oDigest,oState,oDebit,oCredit,oRemark) values('"&oTime&"','"&oCompany&"','"&oInvoice&"','"&oDigest&"','"&oState&"','"&oDebit&"','"&oCredit&"','"&oRemark&"')")
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub

Sub InfoEditOutin()
%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.getElementById('oMoney').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '金额<%=alert04%>'});document.getElementById('oMoney').focus();return false;}
		}
		-->
		</script>
	<form name="Save" action="?action=SaveEditOutin" method="post" onSubmit="return CheckInput();">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<col width="100"><col width="200"><col width="100">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="4"><B>编辑记录</B></td>
					</tr>
							<%
							oState = EasyCrm.getNewItem("Plugin_Finance_Outin","Id",""&id&"","oState")
							oDebit = EasyCrm.getNewItem("Plugin_Finance_Outin","Id",""&id&"","oDebit")
							oCredit = EasyCrm.getNewItem("Plugin_Finance_Outin","Id",""&id&"","oCredit")
							oMoney = oDebit + oCredit
							%>
					<tr>
						<td class="td_l_r title" >日期</td>
						<td class="td_r_l" colspan=3><input name="oTime" type="text" maxlength="10" id="oTime" class="Wdate int" size="15" onFocus="WdatePicker()" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Plugin_Finance_Outin","Id",""&id&"","oTime"),2)%>" /> </td>
					</tr>
					<tr>
						<td class="td_l_r title" >公司名称</td>
						<td class="td_r_l" ><input name="oCompany" type="text" id="oCompany" class="int" size="30" value="<%=EasyCrm.getNewItem("Plugin_Finance_Outin","Id",""&id&"","oCompany")%>" /></td>
						<td class="td_l_r title" >票号</td>
						<td class="td_r_l" ><input name="oInvoice" type="text" id="oInvoice" class="int" size="30" value="<%=EasyCrm.getNewItem("Plugin_Finance_Outin","Id",""&id&"","oInvoice")%>" /></td>
					</tr>
					<tr>
						<td class="td_l_r title" >摘要</td>
						<td class="td_r_l"><input name="oDigest" type="text" id="oDigest" class="int" size="30" value="<%=EasyCrm.getNewItem("Plugin_Finance_Outin","Id",""&id&"","oDigest")%>" /></td>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 金额</td>
						<td class="td_r_l">
							<input name="osType" type="radio" class="noborder" value="oDebit" <%if oDebit>0 then%>checked<%end if%> > 收入+　
							<input name="osType" type="radio" class="noborder" value="oCredit" <%if oCredit>0 then%>checked<%end if%> > 支出-　
							<b style="color:red;">￥</b><input name="oMoney" type="text" id="oMoney" class="int" size="6" onkeyup='this.value=this.value.replace(/\D/gi,"")' value="<%=oMoney%>" /> RMB</td>
					</tr>
					<tr>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 状态</td>
						<td class="td_r_l" colspan=3 >
							<input name="oState" type="radio" class="noborder" value="已完成" <%if oState="已完成" then %>checked<%end if%>> 已完成　
							<input name="oState" type="radio" class="noborder" value="未完成" <%if oState="未完成" then %>checked<%end if%>> 未完成　</td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" style="padding:10px;" colspan=3> <textarea name="oRemark" id="oRemark" style="width:99%;height:100px;"><%=EasyCrm.getNewItem("Plugin_Finance_Outin","Id",""&id&"","oRemark")%></textarea></td>
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
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('oRemark',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
	
<%
End Sub

Sub SaveEditOutin()
	id = Request("id")
	
	oTime = Request.Form("oTime")
	oCompany = Request.Form("oCompany")
	oInvoice = Request.Form("oInvoice")
	oDigest = Request.Form("oDigest")
	oState = Request.Form("oState")
	
	if Request.Form("osType") = "oDebit" then
	oDebit = Request.Form("oMoney")
	oCredit = 0
	elseif Request.Form("osType") = "oCredit" then
	oDebit = 0
	oCredit = Request.Form("oMoney")
	end if
	
	oRemark = Request.Form("oRemark")
	
	conn.execute "UPDATE Plugin_Finance_Outin SET oTime='"&oTime&"',oCompany='"&oCompany&"',oInvoice='"&oInvoice&"',oDigest='"&oDigest&"',oState='"&oState&"',oDebit='"&oDebit&"',oCredit='"&oCredit&"',oRemark='"&oRemark&"' Where id="&id
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub


%>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>