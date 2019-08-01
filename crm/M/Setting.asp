<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
If Session("CRM_level") < 9 Then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
end if
	'获取get值
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	if otype="" then otype="main"
%><%=Header%>

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
	
	TempStr = TempStr & "%" & chr(62) & VbCrLf
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
%>
<!-- start header -->
    <div id="header">
         <a href="System.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start page -->
    <div class="page">
	<div class="simplebox">
            	<h1 class="titleh">全局设置</h1>
                <div class="content">
                	
                <form action="?Action=Edit" method="post">
                    <div class="form-line">
                   	  <label class="st-label">公司名称</label>
					  <input name="title" type="text" id="title" class="int" value="<%=title%>" ><BR><span class="info_help help01">例：XX公司-客户管理系统</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">系统目录</label>
					   <input name="SiteUrl" type="text" id="SiteUrl" class="int" value="<%=SiteUrl%>" size="5"><BR><span class="info_help help01">以"/"结尾，例 根目录：/ 　子目录：/ECRM/ </span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">风格路径</label>
					   <input name="Skinurl" type="text" id="Skinurl" class="int" value="<%=Skinurl%>"><BR><span class="info_help help01">默认：Skin/default/</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">分页数量</label>
					   <input name="DataPageSize" type="text" id="DataPageSize" class="int" value="<%=DataPageSize%>" size="5"><BR><span class="info_help help01">每页显示多少条数据</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">操作提示</label>
						<input name="YNalert" type="radio" class="noborder" value="1"<%IF ""&YNalert&"" = 1 then Response.Write("  checked") end if%>> 是　
						<input name="YNalert" type="radio" class="noborder" value="0"<%IF ""&YNalert&"" = 0 then Response.Write("  checked") end if%>> 否　
						<BR><span class="info_help help01">操作是否弹出确认窗口，隐藏会减少操作步骤</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">跟进结束</label>
						<% = EasyCrm.getSelect("SelectData","Select_Type","CRTypeEnd",""&CRTypeEnd&"") %>
						<BR><span class="info_help help01">客户跟进流程结束时的状态</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">记录日志</label>
						<input name="YnUserLog" type="radio" class="noborder" value="1"<%IF ""&YnUserLog&""=1 then Response.Write("checked") end if%>> 是　
						<input name="YnUserLog" type="radio" class="noborder" value="0"<%IF ""&YnUserLog&""=0 then Response.Write("checked") end if%>> 否　
						<BR><span class="info_help help01">登录日志和操作记录</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">订单编号</label>
						<input name="YnDDNum" type="radio" class="noborder" value="1"<%IF ""&YnDDNum&""=1 then Response.Write("checked") end if%>> 自动　
						<input name="YnDDNum" type="radio" class="noborder" value="0"<%IF ""&YnDDNum&""=0 then Response.Write("checked") end if%>> 手动　
						<BR><span class="info_help help01">自动格式为：DD20140606120000001</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">合同编号</label>
						<input name="YnHTNum" type="radio" class="noborder" value="1"<%IF ""&YnHTNum&""=1 then Response.Write("checked") end if%>> 自动　
						<input name="YnHTNum" type="radio" class="noborder" value="0"<%IF ""&YnHTNum&""=0 then Response.Write("checked") end if%>> 手动　
						<BR><span class="info_help help01">自动格式为：HT20140606120000001</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">删除原因</label>
						<input name="YnDelReason" type="radio" class="noborder" value="1"<%IF ""&YnDelReason&""=1 then Response.Write("checked") end if%>> 是　
						<input name="YnDelReason" type="radio" class="noborder" value="0"<%IF ""&YnDelReason&""=0 then Response.Write("checked") end if%>> 否　
						<BR><span class="info_help help01">删除客户档案是否需要填写原因</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">转移保留</label>
						<input name="SaveOldUser" type="radio" class="noborder" value="1"<%IF ""&SaveOldUser&""=1 then Response.Write("checked") end if%>> 是　
						<input name="SaveOldUser" type="radio" class="noborder" value="0"<%IF ""&SaveOldUser&""=0 then Response.Write("checked") end if%>> 否　
						<BR><span class="info_help help01">客户转移后，是否保留原有业务员</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">公海审核</label>
						<input name="YNRecycler" type="radio" class="noborder" value="1"<%IF ""&YNRecycler&""=1 then Response.Write("checked") end if%>> 是　
						<input name="YNRecycler" type="radio" class="noborder" value="0"<%IF ""&YNRecycler&""=0 then Response.Write("checked") end if%>> 否　
						<BR><span class="info_help help01">公海申请客户是否需要审核</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">客户唯一</label>
						<input name="ClientOnly1" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 1, 1) = 1 then%> checked<%end if%> > 公司名称 + 
						<input name="ClientOnly2" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 2, 1) = 1 then%> checked<%end if%> > 联系人 + 
						<input name="ClientOnly3" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 3, 1) = 1 then%> checked<%end if%> > 手机
						<BR><span class="info_help help01">判断客户唯一的标准</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">处理乱码</label>
						<input name="SelectCharset" type="radio" class="noborder" value="1"<%IF ""&SelectCharset&""=1 then Response.Write("  checked") end if%>> 是　
						<input name="SelectCharset" type="radio" class="noborder" value="0"<%IF ""&SelectCharset&""=0 then Response.Write("  checked") end if%>> 否　
						<BR><span class="info_help help01">不同服务器环境可能会出现乱码，请改变设置</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">保持在线</label>
						<input name="Keeponline" type="radio" class="noborder" value="1"<%IF ""&Keeponline&""=1 then Response.Write("  checked") end if%>> 是　
						<input name="Keeponline" type="radio" class="noborder" value="0"<%IF ""&Keeponline&""=0 then Response.Write("  checked") end if%>> 否　
						<BR><span class="info_help help01">混合保存登录状态，有一定安全隐患</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">Cookie值</label>
						<input name="CookieKey" type="text" id="CookieKey" class="int" value="<%=CookieKey%>">　
						<BR><span class="info_help help01">6-16位数字+字母，重新登录生效</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">上传类型</label>
						<input name="uploadtype" type="text" id="uploadtype" class="int" value="<%=uploadtype%>">　
						<BR><span class="info_help help01">例：gif/jpg/doc/xls/rar</span>
                    </div>
                    
                    <div class="form-line">
                    <input type="submit" name="button" id="button" value="&nbsp;保 存&nbsp;" class="submit-button" />
                    <input type="reset" name="button" id="button2" value="&nbsp;重 置&nbsp;" class="reset-button" />
                    </div>

                  </form>
                </div>
			</div>
		<%=Footer%>
            
    
    <div class="clear"></div>
    </div>
    <!-- end page -->
<script type="text/javascript" src="js/frame.js"></script>
</body>
</html>
