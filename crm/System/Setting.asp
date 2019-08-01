<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 5, 1) = 1 Then %>
<%
	'获取get值
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	if otype="" then otype="main"
%>
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
<body style="padding-top:35px;padding-bottom:55px;">

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：系统管理 > 全局设置</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>
<%
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

If request("action")="Edit" then
title = replace(Trim(Request.Form("title")),CHR(34),"'")
SiteUrl = replace(Trim(Request.Form("SiteUrl")),CHR(34),"'")
Skinurl = replace(Trim(Request.Form("Skinurl")),CHR(34),"'")
DataPageSize = replace(Trim(Request.Form("DataPageSize")),CHR(34),"'")
YNalert = replace(Trim(Request.Form("YNalert")),CHR(34),"'")
CRTypeEnd = replace(Trim(Request.Form("CRTypeEnd")),CHR(34),"'")
YnUserLog = replace(Trim(Request.Form("YnUserLog")),CHR(34),"'")
YnDDNum = replace(Trim(Request.Form("YnDDNum")),CHR(34),"'")
YnHTNum = replace(Trim(Request.Form("YnHTNum")),CHR(34),"'")
YnDelReason = replace(Trim(Request.Form("YnDelReason")),CHR(34),"'")
SaveOldUser = replace(Trim(Request.Form("SaveOldUser")),CHR(34),"'")
YNRecycler = replace(Trim(Request.Form("YNRecycler")),CHR(34),"'")
ClientOnly = ""
	for i = 1 to 3
		if Request("ClientOnly" & i) = "1" then
			ClientOnly = ClientOnly & "1"
		else
			ClientOnly = ClientOnly & "0"
		end if
	next
SelectCharset = replace(Trim(Request.Form("SelectCharset")),CHR(34),"'")
language = replace(Trim(Request.Form("language")),CHR(34),"'")
uploadtype = replace(Trim(Request.Form("uploadtype")),CHR(34),"'")
Keeponline = replace(Trim(Request.Form("Keeponline")),CHR(34),"'")
CookieKey = replace(Trim(Request.Form("CookieKey")),CHR(34),"'")
gdzy = replace(Trim(Request.Form("gdzy")),CHR(34),"'")

Dim n,TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "Dim title,SiteUrl,Skinurl,DataPageSize,YNalert,CRTypeEnd,YnUserLog,YnDDNum,YnHTNum,YnDelReason,SaveOldUser,YNRecycler,ClientOnly,SelectCharset,language,uploadtype,Keeponline,CookieKey" & VbCrLf
	
	TempStr = TempStr & "'全局配置" & VbCrLf
	TempStr = TempStr & "title="& Chr(34) & title & Chr(34) &" '系统名称" & VbCrLf
	TempStr = TempStr & "SiteUrl="& Chr(34) & SiteUrl & Chr(34) &" '安装目录" & VbCrLf
	TempStr = TempStr & "Skinurl="& Chr(34) & Skinurl & Chr(34) &" '风格路径" & VbCrLf
	TempStr = TempStr & "DataPageSize="& DataPageSize &" '分页数量" & VbCrLf
	TempStr = TempStr & "YNalert="& YNalert &" '操作提示" & VbCrLf
	TempStr = TempStr & "CRTypeEnd="& Chr(34) & CRTypeEnd & Chr(34) &" '跟进流程结束状态" & VbCrLf
	TempStr = TempStr & "YnUserLog="& YnUserLog &" '记录登录日志" & VbCrLf
	TempStr = TempStr & "YnDDNum="& YnDDNum &" '订单编号生成方式" & VbCrLf
	TempStr = TempStr & "YnHTNum="& YnHTNum &" '合同编号生成方式" & VbCrLf
	TempStr = TempStr & "YnDelReason="& YnDelReason &" '删除客户档案是否需要填写原因" & VbCrLf
	TempStr = TempStr & "SaveOldUser="& SaveOldUser &" '客户转移后，是否保留原有业务员" & VbCrLf
	TempStr = TempStr & "YNRecycler="& YNRecycler &" '公海申请客户是否需要审核" & VbCrLf
	TempStr = TempStr & "ClientOnly="& Chr(34) & ClientOnly & Chr(34) &" '判断客户唯一的标准" & VbCrLf
	TempStr = TempStr & "SelectCharset="& SelectCharset & " '处理乱码" & VbCrLf
	TempStr = TempStr & "language="& Chr(34) & language & Chr(34) &" '系统语言" & VbCrLf 
	TempStr = TempStr & "uploadtype="& Chr(34) & uploadtype & Chr(34) &" '上传文件后缀" & VbCrLf
	TempStr = TempStr & "Keeponline="& Keeponline &" '保持在线" & VbCrLf
	TempStr = TempStr & "CookieKey="& Chr(34) & CookieKey & Chr(34) & " '识别码" & VbCrLf & VbCrLf
	TempStr = TempStr & "gdzy="& Chr(34) & gdzy & Chr(34) & " '跟单转移" & VbCrLf & VbCrLf
	TempStr = TempStr & "%" & chr(62) & VbCrLf
	
	    sqlgh="UPDATE Client SET cYn = 0 WHERE  cuser  in (select uname from [user] where uLevel<>9 ) and cType<>'"&CRTypeEnd&"' and (DATEDIFF(d, cLastUpdated, { fn NOW() }) >" &gdzy&")"
        conn.execute(sqlgh)
		sqlgs="UPDATE Client SET cYn = 1 WHERE cType='"&CRTypeEnd&"'"
		conn.execute(sqlgs)
		ADODB_SaveToFile TempStr,"../data/Config.asp"
	If GBL_CHK_TempStr = "" Then
		if ""&YNalert&"" = 1 then
		Response.Write("<script language=javascript>alert('"&alert2&"');this.location.href='Setting.asp';</script>")
		else
		Response.Write("<script language=javascript>this.location.href='Setting.asp';</script>")
		end if
	Else
		Response.Write("<script language=javascript>alert('"&alert01&"');this.location.href='Setting.asp';</script>")
	End If
End if
%><form action="?Action=Edit" method="post">
<table width="100%" border="0" cellpadding="0" cellspacing="0" id="TableMain">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdt10"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="100" />
        <tr class="tr_t"> 
			<td class="td_l_l" COLSPAN="2"><B>全局设置</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">公司名称</td>
			<td class="td_r_l"> <input name="title" type="text" id="title" class="int" value="<%=title%>" size="40">　<span class="info_help help01">例：XX公司-客户管理系统</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">系统目录</td>
			<td class="td_r_l"> <input name="SiteUrl" type="text" id="SiteUrl" class="int" value="<%=SiteUrl%>" size="5">　<span class="info_help help01">以"/"结尾，例 根目录：/ 　子目录：/ECRM/ </span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">风格路径</td>
			<td class="td_r_l"> <input name="Skinurl" type="text" id="Skinurl" class="int" value="<%=Skinurl%>" size="40">　<span class="info_help help01">默认：Skin/default/</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">分页数量</td>
			<td class="td_r_l"> <input name="DataPageSize" type="text" id="DataPageSize" class="int" value="<%=DataPageSize%>" size="5">　<span class="info_help help01">每页显示多少条数据</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">操作提示</td>
			<td class="td_r_l"> 
			<input name="YNalert" type="radio" class="noborder" value="1"<%IF ""&YNalert&"" = 1 then Response.Write("  checked") end if%>> 是　
			<input name="YNalert" type="radio" class="noborder" value="0"<%IF ""&YNalert&"" = 0 then Response.Write("  checked") end if%>> 否　<span class="info_help help01">操作执行后是否弹出确认窗口，隐藏会减少操作步骤</span>
			</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">跟进结束</td>
			<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_Type","CRTypeEnd",""&CRTypeEnd&"") %>　
			<span class="info_help help01">客户跟进流程结束时的状态</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">记录日志</td>
			<td class="td_r_l"> 
			<input name="YnUserLog" type="radio" class="noborder" value="1"<%IF ""&YnUserLog&""=1 then Response.Write("checked") end if%>> 是　
			<input name="YnUserLog" type="radio" class="noborder" value="0"<%IF ""&YnUserLog&""=0 then Response.Write("checked") end if%>> 否　
			<span class="info_help help01">登录日志和操作记录</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">订单编号</td>
			<td class="td_r_l"> 
			<input name="YnDDNum" type="radio" class="noborder" value="1"<%IF ""&YnDDNum&""=1 then Response.Write("checked") end if%>> 自动　
			<input name="YnDDNum" type="radio" class="noborder" value="0"<%IF ""&YnDDNum&""=0 then Response.Write("checked") end if%>> 手动　
			<span class="info_help help01">自动格式为：DD20140606120000001</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">合同编号</td>
			<td class="td_r_l"> 
			<input name="YnHTNum" type="radio" class="noborder" value="1"<%IF ""&YnHTNum&""=1 then Response.Write("checked") end if%>> 自动　
			<input name="YnHTNum" type="radio" class="noborder" value="0"<%IF ""&YnHTNum&""=0 then Response.Write("checked") end if%>> 手动　
			<span class="info_help help01">自动格式为：HT20140606120000001</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">删除原因</td>
			<td class="td_r_l"> 
			<input name="YnDelReason" type="radio" class="noborder" value="1"<%IF ""&YnDelReason&""=1 then Response.Write("checked") end if%>> 是　
			<input name="YnDelReason" type="radio" class="noborder" value="0"<%IF ""&YnDelReason&""=0 then Response.Write("checked") end if%>> 否　
			<span class="info_help help01">删除客户档案是否需要填写原因</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">转移保留</td>
			<td class="td_r_l"> 
			<input name="SaveOldUser" type="radio" class="noborder" value="1"<%IF ""&SaveOldUser&""=1 then Response.Write("checked") end if%>> 是　
			<input name="SaveOldUser" type="radio" class="noborder" value="0"<%IF ""&SaveOldUser&""=0 then Response.Write("checked") end if%>> 否　
			<span class="info_help help01">客户转移后，是否保留原有业务员</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">公海审核</td>
			<td class="td_r_l"> 
			<input name="YNRecycler" type="radio" class="noborder" value="1"<%IF ""&YNRecycler&""=1 then Response.Write("checked") end if%>> 是　
			<input name="YNRecycler" type="radio" class="noborder" value="0"<%IF ""&YNRecycler&""=0 then Response.Write("checked") end if%>> 否　
			<span class="info_help help01">公海申请客户是否需要审核</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">客户唯一</td>
			<td class="td_r_l"> 
			<input name="ClientOnly1" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 1, 1) = 1 then%> checked<%end if%> > 公司名称 + 
			<input name="ClientOnly2" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 2, 1) = 1 then%> checked<%end if%> > 联系人 + 
			<input name="ClientOnly3" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 3, 1) = 1 then%> checked<%end if%> > 手机号码　
			<span class="info_help help01">判断客户唯一的标准</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">处理乱码</td>
			<td class="td_r_l"> 
				<input name="SelectCharset" type="radio" class="noborder" value="1"<%IF ""&SelectCharset&""=1 then Response.Write("  checked") end if%>> 是　
				<input name="SelectCharset" type="radio" class="noborder" value="0"<%IF ""&SelectCharset&""=0 then Response.Write("  checked") end if%>> 否　<span class="info_help help01">不同服务器环境可能会出现乱码，请改变设置</span>
			</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">默认语言</td>
			<td class="td_r_l"> 
				<input name="language" type="radio" class="noborder" value="zh-cn"<%IF ""&language&""="zh-cn" then Response.Write("  checked") end if%>> 中文　<span class="info_help help01">风格包存放路径：/lang/zh-cn/</span>
			</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">保持在线</td>
			<td class="td_r_l"> 
				<input name="Keeponline" type="radio" class="noborder" value="1"<%IF ""&Keeponline&""=1 then Response.Write("  checked") end if%>> 是　
				<input name="Keeponline" type="radio" class="noborder" value="0"<%IF ""&Keeponline&""=0 then Response.Write("  checked") end if%>> 否　<span class="info_help help01">混合Session和Cookies保存登录状态，有一定安全隐患</span>
			</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">Cookie值</td>
			<td class="td_r_l"> 
				<input name="CookieKey" type="text" id="CookieKey" class="int" value="<%=CookieKey%>" size="40">　<span class="info_help help01">6-16位数字+字母，重新登录生效</span>
			</td>
        </tr>
		<tr class="tr"> 
            <td class="td_l_r title">跟单转移</td>
            <td class="td_r_l"><input name="gdzy" type="text" id="gdzy" class="int" value="<%=gdzy%>" size="15">
              最后跟单超多少天客户信息自动进入公海</td>
          </tr>
        <tr class="tr"> 
			<td class="td_l_r title">上传类型</td>
			<td class="td_r_l"> <input name="uploadtype" type="text" id="uploadtype" class="int" value="<%=uploadtype%>" size="40">　<span class="info_help help01">例：gif/jpg/doc/xls/rar</span>
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
			<input name="Submit" type="submit" class="button45" id="Submit" value=" 保存 ">
		</td>
	</tr>
</table>
</div>
</form>

<%
else
Response.write"<script>alert("""&alert31&""");location.href=""../"";</script>"
end if
%>
<%
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
%><% Set EasyCrm = nothing %>