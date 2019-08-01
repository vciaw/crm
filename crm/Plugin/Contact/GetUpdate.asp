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
Case "InfoEdit"
    Call InfoEdit()
Case "SaveEdit"
    Call SaveEdit()
End Select

Sub Add()
%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.getElementById('cCompany').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '公司名称<%=alert04%>'});document.getElementById('cCompany').focus();return false;}
			if(document.getElementById('cLinkman').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '联系人<%=alert04%>'});document.getElementById('cLinkman').focus();return false;}
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
						<td class="td_l_l" COLSPAN="4"><B>新增联系人</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 公司名称</td>
						<td class="td_r_l" colspan=3 ><input name="cCompany" type="text" class="int" id="cCompany" size="40"> </td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">分类</td>
						<td class="td_r_l" colspan=3> 
						<%
							str = split(""&Plugin_contact_class&"",",")
							for i = 0 to ubound(str)
							response.Write "<input name=""cClass"" type=""radio"" class=""noborder"" value="""&str(i)&"""> "&str(i)&"　"
							next
						%> 
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 联系人</td>
						<td class="td_r_l" ><input name="cLinkman" type="text" id="cLinkman" class="int" size="15" /></td>
						<td class="td_l_r title" >电话</td>
						<td class="td_r_l" ><input name="cTel" type="text" id="cTel" class="int" size="25" /></td>
					</tr>
					<tr>
						<td class="td_l_r title" >职位</td>
						<td class="td_r_l" >
						<select name="cZhiwei" class="int">
							<option value="">请选择</option>
							<%
							str = split(""&Plugin_contact_zhiwei&"",",")
							for i = 0 to ubound(str)
							response.Write "<option value="&str(i)&">"&str(i)&"</option>"
							next
							%>
						</select>
					</td>
						<td class="td_l_r title" >部门</td>
						<td class="td_r_l" >
							<select name="cGroup" class="int">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_contact_group&"",",")
								for i = 0 to ubound(str)
								response.Write "<option value="&str(i)&">"&str(i)&"</option>"
								next
								%>
							</select>
						</td>
					</tr>
					<tr>
						<td class="td_l_r title">ＱＱ</td>
						<td class="td_r_l"><input name="cQQ" type="text" class="int" id="cQQ" size="20"></td>
						<td class="td_l_r title">主营</td>
						<td class="td_r_l"><input name="cProducts" type="text" class="int" id="cProducts" size="30"></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" style="padding:10px;" colspan=3> <textarea name="cInfo" id="cInfo" style="width:99%;height:70px;"></textarea></td>
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
	 new tqEditor('cInfo',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
<%
End Sub

Sub SaveAdd()
	cClass = Trim(Request("cClass"))
	cCompany = Trim(Request("cCompany"))
	cGroup = Trim(Request("cGroup"))
	cZhiwei = Trim(Request("cZhiwei"))
	cLinkman = Trim(Request("cLinkman"))
	cTel = Trim(Request("cTel"))
	cQQ = Trim(Request("cQQ"))
	cProducts = Trim(Request("cProducts"))
	cInfo = Trim(Request("cInfo"))
	cUser = Trim(Request("cUser"))
	conn.execute("insert into [Plugin_Contact] (cClass,cCompany,cGroup,cZhiwei,cLinkman,cTel,cQQ,cProducts,cInfo,cUser,cTime) values('"&cClass&"','"&cCompany&"','"&cGroup&"','"&cZhiwei&"','"&cLinkman&"','"&cTel&"','"&cQQ&"','"&cProducts&"','"&cInfo&"','"&cUser&"','"&Now()&"')")
	
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
			if(document.getElementById('cCompany').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '公司名称<%=alert04%>'});document.getElementById('cCompany').focus();return false;}
			if(document.getElementById('cLinkman').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '联系人<%=alert04%>'});document.getElementById('cLinkman').focus();return false;}
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
						<td class="td_l_l" COLSPAN="4"><B>修改联系人</B></td>
					</tr>
					<tr>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 公司名称</td>
						<td class="td_r_l" colspan=3 ><input name="cCompany" type="text" class="int" id="cCompany" size="40" value="<%=EasyCrm.getNewItem("Plugin_Contact","Id",""&id&"","cCompany")%>" ></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">分类</td>
						<td class="td_r_l" colspan=3>  
						<%
							str = split(""&Plugin_contact_class&"",",")
							for i = 0 to ubound(str)
							if EasyCrm.getNewItem("Plugin_Contact","Id",""&id&"","cClass") = str(i) then
							response.Write "<input name=""cClass"" type=""radio"" class=""noborder"" value="""&str(i)&""" checked> "&str(i)&"　"
							else
							response.Write "<input name=""cClass"" type=""radio"" class=""noborder"" value="""&str(i)&""" > "&str(i)&"　"
							end if
							next
						%>
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" ><font color="#color:#ff0000">*</font> 联系人</td>
						<td class="td_r_l" ><input name="cLinkman" type="text" id="cLinkman" class="int" size="15" value="<%=EasyCrm.getNewItem("Plugin_Contact","Id",""&id&"","cLinkman")%>" /></td>
						<td class="td_l_r title" >电话</td>
						<td class="td_r_l" ><input name="cTel" type="text" id="cTel" class="int" size="25" value="<%=EasyCrm.getNewItem("Plugin_Contact","Id",""&id&"","cTel")%>" /></td>
					</tr>
					<tr>
						<td class="td_l_r title" >职位</td>
						<td class="td_r_l" >
							<select name="cZhiwei" class="int">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_contact_zhiwei&"",",")
								for i = 0 to ubound(str)
								if EasyCrm.getNewItem("Plugin_Contact","Id",""&id&"","cZhiwei") = str(i) then
								response.Write "<option value="&str(i)&" selected>"&str(i)&"</option>"
								else
								response.Write "<option value="&str(i)&">"&str(i)&"</option>"
								end if
								next
								%>
							</select>
						</td>
						<td class="td_l_r title" >部门</td>
						<td class="td_r_l" >
							<select name="cGroup" class="int">
								<option value="">请选择</option>
								<%
								str = split(""&Plugin_contact_group&"",",")
								for i = 0 to ubound(str)
								if EasyCrm.getNewItem("Plugin_Contact","Id",""&id&"","cGroup") = str(i) then
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
						<td class="td_l_r title">ＱＱ</td>
						<td class="td_r_l"><input name="cQQ" type="text" class="int" id="cQQ" size="20" value="<%=EasyCrm.getNewItem("Plugin_Contact","Id",""&id&"","cQQ")%>" ></td>
						<td class="td_l_r title">主营</td>
						<td class="td_r_l"><input name="cProducts" type="text" class="int" id="cProducts" size="30" value="<%=EasyCrm.getNewItem("Plugin_Contact","Id",""&id&"","cProducts")%>" ></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top">备注</td>
						<td class="td_r_l" style="padding:10px;" colspan=3> <textarea name="cInfo" id="cInfo" style="width:99%;height:70px;"><%=EasyCrm.getNewItem("Plugin_Contact","Id",""&id&"","cInfo")%></textarea></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
					<input name="id" type="hidden" id="id" value="<% = Id %>">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('cInfo',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
<%
End Sub

Sub SaveEdit()
	id = Request("id")
	cClass = Trim(Request("cClass"))
	cCompany = Trim(Request("cCompany"))
	cGroup = Trim(Request("cGroup"))
	cZhiwei = Trim(Request("cZhiwei"))
	cLinkman = Trim(Request("cLinkman"))
	cTel = Trim(Request("cTel"))
	cQQ = Trim(Request("cQQ"))
	cProducts = Trim(Request("cProducts"))
	cInfo = Trim(Request("cInfo"))
	cUser = Trim(Request("cUser"))
	conn.execute "UPDATE Plugin_Contact SET cClass='"&cClass&"',cCompany='"&cCompany&"',cGroup='"&cGroup&"',cZhiwei='"&cZhiwei&"',cLinkman='"&cLinkman&"',cTel='"&cTel&"',cQQ='"&cQQ&"',cProducts='"&cProducts&"',cInfo='"&cInfo&"' Where id="&id
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
End Sub
%>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>