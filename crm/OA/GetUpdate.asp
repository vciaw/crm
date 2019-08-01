<!--#include file="../Data/Conn.asp"--><!--#include file="../UpLoad/UpLoad.asp"--><!--#include file="../data/EasyCrm.asp"-->
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
Case "Setting"
    Call Setting()
Case "Notice"
    Call Notice()
Case "Receiver"
    Call Receiver()
Case "Report"
    Call Report()
Case "Calendar"
    Call Calendar()
Case "Soft"
    Call Soft()
End Select

Sub Notice()
	if sType = "Add" then
	%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.getElementById('ONclass').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Notice_ONclass & alert04%>'});document.getElementById('ONclass').focus();return false;}
			if(document.getElementById('ONtitle').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Notice_ONtitle & alert04%>'});document.getElementById('ONtitle').focus();return false;}
		}
		-->
		</script>
	<form name="Save" action="?action=Notice&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="2"><B><%=L_Top_Notice_add%></B></td>
					</tr>
					<tr>
						<td class="td_l_r title" width="100"> <%=L_Notice_ONclass%></td>
						<td class="td_r_l"> <% = EasyCrm.getRadio("SelectData","Select_NoticeClass","ONclass","") %></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top"><font color="#FF0000">*</font> <%=L_Notice_ONtitle%></td>
						<td class="td_r_l"> <input type="text" class="int" name="ONtitle" id="ONtitle" size="50" maxlength="50">　
						<input name="ONStar" type="radio" class="noborder" value="1"> <img src="<%=SiteUrl&Skinurl%>images/ico/star.png" border=0>　
						<input name="ONStar" type="radio" class="noborder" value="0" checked> <img src="<%=SiteUrl&Skinurl%>images/ico/starno.png" border=0>
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top"><%=L_Notice_ONcontent%></td>
						<td class="td_r_l" style="padding:10px;"> <textarea name="ONcontent" id="ONcontent" style="width:99%;height:240px;"></textarea></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
					<input name="ONuser" type="hidden" value="<%=Session("CRM_name")%>">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('ONcontent',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
	<%
	elseIF sType="SaveAdd" then

		ONclass = Request.Form("ONclass")
		ONtitle = Request.Form("ONtitle")
		ONStar = Request.Form("ONStar")
		ONcontent = EasyCrm.htmlEncode2(Request.Form("ONcontent"))
		ONuser = Request.Form("ONuser")
		
		conn.execute("insert into [OA_Notice] (ONclass,ONtitle,ONStar,ONcontent,ONuser,ONaddtime,ONedittime) values('"&ONclass&"','"&ONtitle&"','"&ONStar&"','"&ONcontent&"','"&ONuser&"','"&Now()&"','"&Now()&"')")
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	elseIF sType = "Edit" then
	%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.getElementById('ONclass').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Notice_ONclass & alert04%>'});document.getElementById('ONclass').focus();return false;}
			if(document.getElementById('ONtitle').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Notice_ONtitle & alert04%>'});document.getElementById('ONtitle').focus();return false;}
		}
		-->
		</script>
	<form name="Save" action="?action=Notice&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="2"><B><%=L_Top_Notice_edit%></B></td>
					</tr>
					<tr>
						<td class="td_l_r title" width="100"><%=L_Notice_ONclass%></td>
						<td class="td_r_l"> <% = EasyCrm.getRadio("SelectData","Select_NoticeClass","ONclass",EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONclass")) %></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top"><font color="#FF0000">*</font> <%=L_Notice_ONtitle%></td>
						<td class="td_r_l"> <input type="text" class="int" name="ONtitle" id="ONtitle" size="50" maxlength="50" value="<%=EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONtitle")%>">　
						<input name="ONStar" type="radio" class="noborder" value="1" <%if EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONStar") = 1 then%>checked<%end if%> > <img src="<%=SiteUrl&Skinurl%>images/ico/star.png" border=0>　
						<input name="ONStar" type="radio" class="noborder" value="0" <%if EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONStar") = 0 then%>checked<%end if%> > <img src="<%=SiteUrl&Skinurl%>images/ico/starno.png" border=0>
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top"><%=L_Notice_ONcontent%></td>
						<td class="td_r_l" style="padding:10px;"> <textarea name="ONcontent" id="ONcontent" style="width:99%;height:240px;"><%=EasyCrm.htmlEncode3(EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONcontent"))%></textarea></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
					<input name="ONid" type="hidden" value="<%=Id%>">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('ONcontent',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
	<%
	elseIF sType="SaveEdit" then

		ONid = Request.Form("ONid")
		ONclass = Request.Form("ONclass")
		ONtitle = Request.Form("ONtitle")
		ONStar = Request.Form("ONStar")
		ONcontent = EasyCrm.htmlEncode2(Request.Form("ONcontent"))
		
		conn.execute "UPDATE [OA_Notice] SET ONclass='"&ONclass&"',ONtitle='"&ONtitle&"',ONStar='"&ONStar&"',ONcontent='"&ONcontent&"',ONedittime='"&Now()&"' Where ONid="&ONid
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	elseIF sType = "View" then
	%>
	<style>body{padding:0 0 55px 0;}</style>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<tr class="tr_t"> 
						<td class="td_l_l"><img src="<%=SiteUrl&Skinurl%>images/ico/star<%if EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONStar") = 0 then%>no<%end if%>.png" border=0> <B><%=EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONclass")%></B> : <%=EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONtitle")%></td>
					</tr>
					<tr>
						<td class="td_r_l" style="padding:10px;line-height:2em;"> <%=EasyCrm.htmlEncode3(EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONcontent"))%></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
					<span class="Bottom_pd r fontnobold">最后更新：<%=EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONedittime")%></span>
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('ONcontent',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
	<%
	end if
End Sub

Sub Calendar()
	if sType = "Add" then
	%>
	<style>body{padding:0 0 55px 0;}</style>
	<form name="Save" action="?action=Calendar&sType=SaveAdd" method="post">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
			<textarea name="calendarText" id="calendarText" style="width:99%;height:210px;"></textarea>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
					<input name="currentDate" type="hidden" value="<%=Request("currentDate")%>">
					<input name="calendaruser" type="hidden" value="<%=Session("CRM_name")%>">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<%
	elseIF sType="SaveAdd" then

		calendarText = Request.Form("calendarText")
		calendaruser = Request.Form("calendaruser")
		currentDate = Request.Form("currentDate")
		if calendarText <> "" then
		conn.execute("insert into [calendar] (calendarDate,calendarText,calendaruser) values('"&currentDate&"','"&calendarText&"','"&calendaruser&"')")
		end if
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	elseIF sType = "Edit" then
	%>
	<style>body{padding:0 0 55px 0;}</style>
	<form name="Save" action="?action=Calendar&sType=SaveEdit" method="post">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
			<textarea name="calendarText" id="calendarText" style="width:99%;height:210px;"><%=EasyCrm.getNewItem("calendar","Id",""&Id&"","calendarText")%></textarea>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd ">
					<input name="Id" type="hidden" value="<%=Id%>">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<%
	elseIF sType="SaveEdit" then


		Id = Request.Form("Id")
		calendarText = Request.Form("calendarText")
		if calendarText <> "" then
		conn.execute "UPDATE [calendar] SET calendarText='"&calendarText&"' Where id="&id
		else
		conn.execute "DELETE FROM [calendar] Where id="&id
		end if
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
		
		ONid = Request.Form("ONid")
		ONclass = Request.Form("ONclass")
		ONtitle = Request.Form("ONtitle")
		ONStar = Request.Form("ONStar")
		ONcontent = EasyCrm.htmlEncode2(Request.Form("ONcontent"))
		
		conn.execute "UPDATE [OA_Notice] SET ONclass='"&ONclass&"',ONtitle='"&ONtitle&"',ONStar='"&ONStar&"',ONcontent='"&ONcontent&"',ONedittime='"&Now()&"' Where ONid="&ONid
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	elseIF sType="Del" then
		conn.execute "DELETE FROM [calendar] Where id="&id
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	elseIF sType = "View" then
	%>
	<style>body{padding:0 0 55px 0;}</style>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<tr class="tr_t"> 
						<td class="td_l_l"><img src="<%=SiteUrl&Skinurl%>images/ico/star<%if EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONStar") = 0 then%>no<%end if%>.png" border=0> <B><%=EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONclass")%></B> : <%=EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONtitle")%></td>
					</tr>
					<tr>
						<td class="td_r_l" style="padding:10px;line-height:2em;"> <%=EasyCrm.htmlEncode3(EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONcontent"))%></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
					<span class="Bottom_pd r fontnobold">最后更新：<%=EasyCrm.getNewItem("OA_Notice","ONId",""&Id&"","ONedittime")%></span>
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('ONcontent',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
	<%
	end if
End Sub

Sub Soft()
	if sType = "Add" then
	%>
	<script language="JavaScript">
	<!-- 附件记录必填项提示
	function CheckInput()
	{
		if(document.getElementById('s_class').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '未选择分类！'});document.getElementById('s_class').focus();return false;}
		if(document.getElementById('s_file').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '未选择要上传的附件！'});document.getElementById('s_file').focus();return false;}
	}
	-->
	</script>
		<form name="Save" action="?action=Soft&sType=SaveAdd" method="post" enctype="multipart/form-data" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>上传文件</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> 分类</td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_SoftClass","s_class","") %></td>
							</tr>
							<tr> 
								<td class="td_l_r title">共享</td>
								<td class="td_r_l"> 
									<input name="s_share" type="radio" class="noborder" value="1"> 是　
									<input name="s_share" type="radio" class="noborder" value="0" checked> 否
								</td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> 附件</td>
								<td class="td_r_l" style="padding:5px 10px;"> <input name="s_file" type="file" id="s_file" value="" class="int"></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_File_fContent%></td>
								<td class="td_r_l" style="padding:5px 10px;"> <textarea name="s_content" rows="4" id="s_content" class="int" style="height:50px;width:98%;"></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10"> 
						<div style="float:left;padding:10px 0 0;">
							<input name="s_title" type="hidden" id="s_title" value="">
							<input name="s_user" type="hidden" value="<%=Session("CRM_name")%>">
							<input type="submit" name="Submit" class="button45" onClick="s_title.value=/[^\\]+\.\w+$/.exec(s_file.value)[0]" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
						</div>
					</td> 
				</tr>
			</table>
		</form>
	<%
	elseif sType="SaveAdd" then '保存添加
	
		dim nTime : nTime = Timer()
		dim request,lngUpSize
		Set request=new UpLoadClass
		request.TotalSize= 104857600
		request.MaxSize  = 100000*1024
		request.FileType = ""&uploadtype&""
		request.Savepath = "../soft/"&Session("CRM_account")&"/"
		lngUpSize = request.Open()
		    
		Dim s_file,s_class,s_title,s_share,s_content,s_user
		s_file = request.Savepath & Request.Form("s_file")
		if s_file = request.Savepath then s_file=""
		s_class = Request.Form("s_class")
		s_title = Request.Form("s_title")
		s_share = Request.Form("s_share")
		s_content = Request.Form("s_content")
		s_user = Request.Form("s_user")
		
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [OA_soft] ",conn,3,2
		rs.AddNew
		rs("s_title") = s_title
		rs("s_class") = s_class
		if s_file<>"" then
		rs("s_file") = s_file
		end if
		rs("s_share") = s_share
		rs("s_user") = s_user
		rs("s_content") = s_content
		rs("s_time") = Now()
    	rs.Update
    	rs.Close
    	Set rs = Nothing 
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	elseIF sType = "Edit" then
	%>
		<form name="Save" action="?action=Soft&sType=SaveEdit" method="post" enctype="multipart/form-data" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0"> 
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
						<col width="120" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>上传文件</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> 分类</td>
								<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_SoftClass","s_class",EasyCrm.getNewItem("OA_soft","sId",""&Id&"","s_class")) %></td>
							</tr>
							<tr> 
								<td class="td_l_r title">共享</td>
								<td class="td_r_l"> 
									<input name="s_share" type="radio" class="noborder" value="1" <%if EasyCrm.getNewItem("OA_soft","sId",""&Id&"","s_share")=1 then%>checked<%end if%> > 是　
									<input name="s_share" type="radio" class="noborder" value="0" <%if EasyCrm.getNewItem("OA_soft","sId",""&Id&"","s_share")=0 then%>checked<%end if%> > 否
								</td>
							</tr>
							<tr> 
								<td class="td_l_r title"><font color="#FF0000">*</font> 附件</td>
								<td class="td_r_l" style="padding:5px 10px;"> <input name="s_file" type="file" id="s_file" value="" class="int"> <span class="info_help help01" onmouseover="tip.start(this)" tips="不修改请留空">&nbsp;</span></td>
							</tr>
							<tr> 
								<td class="td_l_r title"><%=L_File_fContent%></td>
								<td class="td_r_l" style="padding:5px 10px;"> <textarea name="s_content" rows="4" id="s_content" class="int" style="height:50px;width:98%;"><%=EasyCrm.getNewItem("OA_soft","sId",""&Id&"","s_content")%></textarea></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" class="td_n pdl10 pdr10"> 
						<div style="float:left;padding:10px 0 0;">
							<input name="sId" type="hidden" id="sId" value="<%=Id%>">
							<input name="s_title" type="hidden" id="s_title" value="">
							<input name="fileold" type="hidden" id="fileold" value="<%=EasyCrm.getNewItem("OA_soft","sId",""&Id&"","s_file")%>">
							<input name="s_user" type="hidden" value="<%=Session("CRM_name")%>">
							<input type="submit" name="Submit" class="button45" onClick="s_title.value=/[^\\]+\.\w+$/.exec(s_file.value)[0]" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
						</div>
					</td> 
				</tr>
			</table>
		</form>
	<%
	elseIF sType="SaveEdit" then
	
		nTime = Timer()
		Set request=new UpLoadClass
		request.TotalSize= 104857600
		request.MaxSize  = 100000*1024
		request.FileType = ""&uploadtype&""
		request.Savepath = "../soft/"&Session("CRM_account")&"/"
		lngUpSize = request.Open()
		    
		sId = CLng(ABS(Request.Form("sId")))
		s_file = request.Savepath & Request.Form("s_file")
		if s_file = request.Savepath then s_file=""
		fileold = Request.Form("fileold")
		s_class = Request.Form("s_class")
		s_share = Request.Form("s_share")
		s_title = Request.Form("s_title")
		s_content = Request.Form("s_content")
		If s_file <> "" and fileold<>"" Then '上传新附件同时删除原附件
			Set fso = CreateObject("Scripting.FileSystemObject")
			IF fso.FileExists(server.MapPath(fileold)) Then
			fso.DeleteFile(server.MapPath(fileold))
			End IF
		End If
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From OA_soft Where sId = " & sId,conn,3,2
		if s_file<>"" then
		rs("s_file") = s_file
		rs("s_title") = s_title
		end if
		rs("s_class") = s_class
		rs("s_share") = s_share
		rs("s_content") = s_content
		rs.Update
		rs.Close
		Set rs = Nothing
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	end if
End Sub

Sub Receiver() '选择收件人

	if tipinfo<>"" then
		Response.Write("<script>art.dialog({title: '提示',time: 1,icon: 'warning',content: '"&tipinfo&"'});</script>")
	end if

	if sType="Receiver" then
	%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<col width="100" />
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="2">选择收件人</td>
					</tr>
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
							<input type="checkbox" name="Receive" onclick="Choose()" value= '<%=rsm("uName")%>'> <%=rsm("uName")%>　
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
			<td valign="top" class="td_n pdl10 pdr10 pdb10"> 
				<div style="float:left;padding:10px 0;">
					<input type="hidden" name="button" id="sReceiver" value="">
					<input type="button" name="button" class="button45" onclick="javascript:$.dialog.open.origin.$('#oReceiver').val(document.getElementById('sReceiver').value);art.dialog.close();" value="确认选择">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</div>
			</td> 
		</tr>
	</table>
	<script LANGUAGE="Javascript">
	eval(function(p,a,c,k,e,r){e=String;if('0'.replace(0,e)==0){while(c--)r[e(c)]=k[c];k=[function(e){return r[e]||e}];e=function(){return'[24-7]'};c=1};while(c--)if(k[c])p=p.replace(new RegExp('\\b'+e(c)+'\\b','g'),k[c]);return p}('function Choose(){4 s=5.getElementsByName("Receive");4 2="";for(4 i=0;i<s.6;i++){if(s[i].checked){2=2+s[i].7+\',\'}}2=2.substr(0,2.6-1);5.getElementById("sReceiver").7=2}',[],8,'||s2||var|document|length|value'.split('|'),0,{}))
	</script>
	<%
	end if
End Sub

Sub Report()

	if tipinfo<>"" then
		Response.Write("<script>art.dialog({title: '提示',time: 1,icon: 'warning',content: '"&tipinfo&"'});</script>")
	end if

	if sType="Add" then
	%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>

		<script language="JavaScript">
		<!-- 必填项提示
		function CheckInput()
		{
			if(document.getElementById('oReport').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Report_oReport & alert04%>'});document.getElementById('oReport').focus();return false;}
			if(document.getElementById('oPlan').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Report_oPlan & alert04%>'});document.getElementById('oPlan').focus();return false;}
		}
		-->
		</script>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<form name="Save" action="?action=Report&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
					<tr class="tr_t"> 
						<td class="td_l_l" COLSPAN="2"><B><%=L_Page_Report_add%></B></td>
					</tr>
					<tr>
						<td class="td_l_r title" width="100">选择对象</td>
						<td class="td_r_l"> 
							<%
								Set rsm = Server.CreateObject("ADODB.Recordset")
								rsm.Open "Select * From [user] where uGroup="&Session("CRM_group")&" and ulevel > "&Session("CRM_level")&" or ulevel=9 ",conn,1,1
								Do While Not rsm.BOF And Not rsm.EOF
							%>
							<input type="checkbox" name="oSendto" value="<%=rsm("uName")%>" <%if rsm("ulevel")=9 then%>checked<%end if%> > <%=rsm("uName")%>　
							<%
								rsm.MoveNext
								Loop
								rsm.Close
								Set rsm = Nothing
							%>
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" width="100"><%=L_Report_oClass%></td>
						<td class="td_r_l"> 
							<input name="oClass" type="radio" class="noborder" value="<%=L_Ribao%>" checked> <%=L_Ribao%>　
							<input name="oClass" type="radio" class="noborder" value="<%=L_Zhoubao%>"> <%=L_Zhoubao%>　
							<input name="oClass" type="radio" class="noborder" value="<%=L_Yuebao%>"> <%=L_Yuebao%>　
							<input name="oClass" type="radio" class="noborder" value="<%=L_Jibao%>"> <%=L_Jibao%>　
							<input name="oClass" type="radio" class="noborder" value="<%=L_Nianbao%>"> <%=L_Nianbao%>　
						</td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top"><%=L_Report_oReport%></td>
						<td class="td_r_l" style="padding:10px;"> <textarea name="oReport" id="oReport" style="width:99%;height:100px;"></textarea></td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top"><%=L_Report_oPlan%></td>
						<td class="td_r_l" style="padding:10px;"> <textarea name="oPlan" id="oPlan" style="width:99%;height:100px;"></textarea></td>
					</tr>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
					<span class="r" style="line-height:30px;color:#f00;">★ 提交后无法修改，请认真填写！</span>
					<input name="oUser" type="hidden" value="<%=Session("CRM_name")%>">
					<input type="submit" name="Submit" class="button45" value="保存">　
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
		</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('oReport',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	 new tqEditor('oPlan',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
	<%
	elseIF sType="SaveAdd" then

		oSendto = Request.Form("oSendto")
		oClass = Request.Form("oClass")
		oReport = EasyCrm.htmlEncode2(Request.Form("oReport"))
		oPlan = EasyCrm.htmlEncode2(Request.Form("oPlan"))
		oUser = Request.Form("oUser")
		
		'选择多人循环插入数据库 
		'arroSendto=split(oSendto,",")
		'for i=0 to ubound(arroSendto)
		conn.execute("insert into [OA_Report] (oSendto,oClass,oTitle,oReport,oPlan,oUser,oIsread,oTime) values('"&oSendto&"','"&oClass&"','"&oUser&" "&L_whoswork&oClass&"','"&oReport&"','"&oPlan&"','"&oUser&"',0,'"&Now()&"')")
		'next
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
	elseIF sType="Reply" then
	id = Request("id")
	'接收报告者阅读后，工作报告设为已读
	if Session("CRM_level") = 9 then
	conn.execute "UPDATE OA_Report SET oIsread='1' Where id="&id
	else
	conn.execute "UPDATE OA_Report SET oIsread='1' Where oSendto like '%"&Session("CRM_name")&"%' and id="&id
	end if
	%>
	<style>body{padding:0 0 55px 0;}</style>
	<script src="<%=SiteUrl&skinurl%>TQEditor/TQEditor.js?skin=silver" type="text/javascript"></script>
	<script language="JavaScript">
	function CheckInput(){
		if(document.getElementById('oReply').value == ""){art.dialog({title: 'Error',time: 1,icon: 'warning',content: '<%=L_Report_oReply & alert04%>'});document.getElementById('oReply').focus();return false;}
	}
	</script>
	<form name="Save" action="?action=Report&sType=SaveReply" method="post" onSubmit="return CheckInput();">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
				<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="100" />
					<tr class="tr_t"> 
						<td class="td_l_l" style="border-right:0;"><B><%=L_Page_Report_view%> <%=oTitle%> </B> </td>
						<td class="td_l_r"><%=EasyCrm.FormatDate(EasyCrm.getNewItem("OA_Report","ID",""&Id&"","oTime"),1)%></td>
					</tr>
					<tr>
						<td class="td_l_r title" width="100" valign="top"><%=L_Report_oReport%> </td>
						<td class="td_r_l" style="padding:5px 10px;"> <%=EasyCrm.htmlEncode3(EasyCrm.getNewItem("OA_Report","ID",""&Id&"","oReport"))%> </td>
					</tr>
					<tr>
						<td class="td_l_r title" valign="top"><%=L_Report_oPlan%> </td>
						<td class="td_r_l" style="padding:5px 10px;"> <%=EasyCrm.htmlEncode3(EasyCrm.getNewItem("OA_Report","ID",""&Id&"","oPlan"))%></td>
					</tr>
					<% If mid(Session("CRM_qx"), 68, 1) = 1 Then %>
					<tr>
						<td class="td_l_r title" valign="top"> <font color="#FF0000">*</font> <%=L_Report_oReply%></td>
						<td class="td_r_l" style="padding:10px;"> <textarea name="oReply" id="oReply" style="width:99%;height:140px;"><%=EasyCrm.getNewItem("OA_Report","ID",""&Id&"","oReply")%></textarea></td>
					</tr>
					<%else%>
					<tr>
						<td class="td_l_r title"><%=L_Report_oReply%></td>
						<td class="td_r_l"> <%if EasyCrm.getNewItem("OA_Report","ID",""&Id&"","oReply")<>"" then%><%=EasyCrm.getNewItem("OA_Report","ID",""&Id&"","oReply")%><%else%><%=L_Wu%><%end if%></td>
					</tr>
					<%end if%>
				</table>
			</td> 
		</tr>
	</table>
	<div class="fixed_bg_B">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td valign="top" class="td_n Bottom_pd "> 
					<input name="id" type="hidden" value="<%=id%>">
					<input name="oReplyOld" type="hidden" value="<%=EasyCrm.getNewItem("OA_Report","ID",""&Id&"","oReply")%>">
					<% If mid(Session("CRM_qx"), 68, 1) = 1 Then %>
					<input type="submit" name="Submit" class="button45" value="保存">　
					<%end if%>
					<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
				</td>
			</tr>
		</table>
	</div>
	</form>
	<script type="text/javascript" defer="true"> 
	 new tqEditor('oReply',{toolbar: 'crm',
	imageUploadUrl: '<%=SiteUrl&skinurl%>TQEditor/UpLoad.asp?oUpLoadType=Images',
	imageFileTypes: '*.jpg;*.gif;*.png;*.jpeg',auto_clean:true});
	</script>
	<%
	elseIF sType="SaveReply" then
	id = Request("id")
	oReply = Request.Form("oReply")
	oReplyOld = Request.Form("oReplyOld")
	conn.execute "UPDATE [OA_Report] SET oReply='"&oReply&"' Where id="&id
	
	'如果批复内容改变，则站内信通知提交者
	if oReply<>oReplyOld then
	conn.execute("insert into [OA_mms_Receive] (oReceiver,oSender,oTitle,oContent,oIsread,oTime) values('"&EasyCrm.getNewItem("OA_Report","ID",""&Id&"","oUser")&"','系统消息','您提交的工作报告已批阅！','请进入【工作报告】栏目查看详情！',0,'"&Now()&"')")
	end if
	
	Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
	Response.End()
	%>
	<%
	end if
End Sub
%>
<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>