<!--#include file="../data/conn.asp" --><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><% If mid(Session("CRM_qx"), 5, 1) = 1 Then %>
<%
	'��ȡgetֵ
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
		<td class="top_left td_t_n td_r_n">��ǰλ�ã�ϵͳ���� > ȫ������</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="ˢ��" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="����" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="ǰ��" onClick=window.location.href="javascript:history.go(1);" />
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
			GBL_CHK_TempStr = "����������֧��ADODB.Stream���޷���ɲ�������ʹ��FTP�ȹ��ܣ���<font color=Red >data/config.asp</font>�ļ������滻�ɿ�������"
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
	
	TempStr = TempStr & "'ȫ������" & VbCrLf
	TempStr = TempStr & "title="& Chr(34) & title & Chr(34) &" 'ϵͳ����" & VbCrLf
	TempStr = TempStr & "SiteUrl="& Chr(34) & SiteUrl & Chr(34) &" '��װĿ¼" & VbCrLf
	TempStr = TempStr & "Skinurl="& Chr(34) & Skinurl & Chr(34) &" '���·��" & VbCrLf
	TempStr = TempStr & "DataPageSize="& DataPageSize &" '��ҳ����" & VbCrLf
	TempStr = TempStr & "YNalert="& YNalert &" '������ʾ" & VbCrLf
	TempStr = TempStr & "CRTypeEnd="& Chr(34) & CRTypeEnd & Chr(34) &" '�������̽���״̬" & VbCrLf
	TempStr = TempStr & "YnUserLog="& YnUserLog &" '��¼��¼��־" & VbCrLf
	TempStr = TempStr & "YnDDNum="& YnDDNum &" '����������ɷ�ʽ" & VbCrLf
	TempStr = TempStr & "YnHTNum="& YnHTNum &" '��ͬ������ɷ�ʽ" & VbCrLf
	TempStr = TempStr & "YnDelReason="& YnDelReason &" 'ɾ���ͻ������Ƿ���Ҫ��дԭ��" & VbCrLf
	TempStr = TempStr & "SaveOldUser="& SaveOldUser &" '�ͻ�ת�ƺ��Ƿ���ԭ��ҵ��Ա" & VbCrLf
	TempStr = TempStr & "YNRecycler="& YNRecycler &" '��������ͻ��Ƿ���Ҫ���" & VbCrLf
	TempStr = TempStr & "ClientOnly="& Chr(34) & ClientOnly & Chr(34) &" '�жϿͻ�Ψһ�ı�׼" & VbCrLf
	TempStr = TempStr & "SelectCharset="& SelectCharset & " '��������" & VbCrLf
	TempStr = TempStr & "language="& Chr(34) & language & Chr(34) &" 'ϵͳ����" & VbCrLf 
	TempStr = TempStr & "uploadtype="& Chr(34) & uploadtype & Chr(34) &" '�ϴ��ļ���׺" & VbCrLf
	TempStr = TempStr & "Keeponline="& Keeponline &" '��������" & VbCrLf
	TempStr = TempStr & "CookieKey="& Chr(34) & CookieKey & Chr(34) & " 'ʶ����" & VbCrLf & VbCrLf
	TempStr = TempStr & "gdzy="& Chr(34) & gdzy & Chr(34) & " '����ת��" & VbCrLf & VbCrLf
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
			<td class="td_l_l" COLSPAN="2"><B>ȫ������</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">��˾����</td>
			<td class="td_r_l"> <input name="title" type="text" id="title" class="int" value="<%=title%>" size="40">��<span class="info_help help01">����XX��˾-�ͻ�����ϵͳ</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">ϵͳĿ¼</td>
			<td class="td_r_l"> <input name="SiteUrl" type="text" id="SiteUrl" class="int" value="<%=SiteUrl%>" size="5">��<span class="info_help help01">��"/"��β���� ��Ŀ¼��/ ����Ŀ¼��/ECRM/ </span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">���·��</td>
			<td class="td_r_l"> <input name="Skinurl" type="text" id="Skinurl" class="int" value="<%=Skinurl%>" size="40">��<span class="info_help help01">Ĭ�ϣ�Skin/default/</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">��ҳ����</td>
			<td class="td_r_l"> <input name="DataPageSize" type="text" id="DataPageSize" class="int" value="<%=DataPageSize%>" size="5">��<span class="info_help help01">ÿҳ��ʾ����������</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">������ʾ</td>
			<td class="td_r_l"> 
			<input name="YNalert" type="radio" class="noborder" value="1"<%IF ""&YNalert&"" = 1 then Response.Write("  checked") end if%>> �ǡ�
			<input name="YNalert" type="radio" class="noborder" value="0"<%IF ""&YNalert&"" = 0 then Response.Write("  checked") end if%>> ��<span class="info_help help01">����ִ�к��Ƿ񵯳�ȷ�ϴ��ڣ����ػ���ٲ�������</span>
			</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">��������</td>
			<td class="td_r_l"> <% = EasyCrm.getSelect("SelectData","Select_Type","CRTypeEnd",""&CRTypeEnd&"") %>��
			<span class="info_help help01">�ͻ��������̽���ʱ��״̬</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">��¼��־</td>
			<td class="td_r_l"> 
			<input name="YnUserLog" type="radio" class="noborder" value="1"<%IF ""&YnUserLog&""=1 then Response.Write("checked") end if%>> �ǡ�
			<input name="YnUserLog" type="radio" class="noborder" value="0"<%IF ""&YnUserLog&""=0 then Response.Write("checked") end if%>> ��
			<span class="info_help help01">��¼��־�Ͳ�����¼</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">�������</td>
			<td class="td_r_l"> 
			<input name="YnDDNum" type="radio" class="noborder" value="1"<%IF ""&YnDDNum&""=1 then Response.Write("checked") end if%>> �Զ���
			<input name="YnDDNum" type="radio" class="noborder" value="0"<%IF ""&YnDDNum&""=0 then Response.Write("checked") end if%>> �ֶ���
			<span class="info_help help01">�Զ���ʽΪ��DD20140606120000001</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">��ͬ���</td>
			<td class="td_r_l"> 
			<input name="YnHTNum" type="radio" class="noborder" value="1"<%IF ""&YnHTNum&""=1 then Response.Write("checked") end if%>> �Զ���
			<input name="YnHTNum" type="radio" class="noborder" value="0"<%IF ""&YnHTNum&""=0 then Response.Write("checked") end if%>> �ֶ���
			<span class="info_help help01">�Զ���ʽΪ��HT20140606120000001</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">ɾ��ԭ��</td>
			<td class="td_r_l"> 
			<input name="YnDelReason" type="radio" class="noborder" value="1"<%IF ""&YnDelReason&""=1 then Response.Write("checked") end if%>> �ǡ�
			<input name="YnDelReason" type="radio" class="noborder" value="0"<%IF ""&YnDelReason&""=0 then Response.Write("checked") end if%>> ��
			<span class="info_help help01">ɾ���ͻ������Ƿ���Ҫ��дԭ��</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">ת�Ʊ���</td>
			<td class="td_r_l"> 
			<input name="SaveOldUser" type="radio" class="noborder" value="1"<%IF ""&SaveOldUser&""=1 then Response.Write("checked") end if%>> �ǡ�
			<input name="SaveOldUser" type="radio" class="noborder" value="0"<%IF ""&SaveOldUser&""=0 then Response.Write("checked") end if%>> ��
			<span class="info_help help01">�ͻ�ת�ƺ��Ƿ���ԭ��ҵ��Ա</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">�������</td>
			<td class="td_r_l"> 
			<input name="YNRecycler" type="radio" class="noborder" value="1"<%IF ""&YNRecycler&""=1 then Response.Write("checked") end if%>> �ǡ�
			<input name="YNRecycler" type="radio" class="noborder" value="0"<%IF ""&YNRecycler&""=0 then Response.Write("checked") end if%>> ��
			<span class="info_help help01">��������ͻ��Ƿ���Ҫ���</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">�ͻ�Ψһ</td>
			<td class="td_r_l"> 
			<input name="ClientOnly1" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 1, 1) = 1 then%> checked<%end if%> > ��˾���� + 
			<input name="ClientOnly2" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 2, 1) = 1 then%> checked<%end if%> > ��ϵ�� + 
			<input name="ClientOnly3" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 3, 1) = 1 then%> checked<%end if%> > �ֻ����롡
			<span class="info_help help01">�жϿͻ�Ψһ�ı�׼</span></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">��������</td>
			<td class="td_r_l"> 
				<input name="SelectCharset" type="radio" class="noborder" value="1"<%IF ""&SelectCharset&""=1 then Response.Write("  checked") end if%>> �ǡ�
				<input name="SelectCharset" type="radio" class="noborder" value="0"<%IF ""&SelectCharset&""=0 then Response.Write("  checked") end if%>> ��<span class="info_help help01">��ͬ�������������ܻ�������룬��ı�����</span>
			</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">Ĭ������</td>
			<td class="td_r_l"> 
				<input name="language" type="radio" class="noborder" value="zh-cn"<%IF ""&language&""="zh-cn" then Response.Write("  checked") end if%>> ���ġ�<span class="info_help help01">�������·����/lang/zh-cn/</span>
			</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">��������</td>
			<td class="td_r_l"> 
				<input name="Keeponline" type="radio" class="noborder" value="1"<%IF ""&Keeponline&""=1 then Response.Write("  checked") end if%>> �ǡ�
				<input name="Keeponline" type="radio" class="noborder" value="0"<%IF ""&Keeponline&""=0 then Response.Write("  checked") end if%>> ��<span class="info_help help01">���Session��Cookies�����¼״̬����һ����ȫ����</span>
			</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_r title">Cookieֵ</td>
			<td class="td_r_l"> 
				<input name="CookieKey" type="text" id="CookieKey" class="int" value="<%=CookieKey%>" size="40">��<span class="info_help help01">6-16λ����+��ĸ�����µ�¼��Ч</span>
			</td>
        </tr>
		<tr class="tr"> 
            <td class="td_l_r title">����ת��</td>
            <td class="td_r_l"><input name="gdzy" type="text" id="gdzy" class="int" value="<%=gdzy%>" size="15">
              ��������������ͻ���Ϣ�Զ����빫��</td>
          </tr>
        <tr class="tr"> 
			<td class="td_l_r title">�ϴ�����</td>
			<td class="td_r_l"> <input name="uploadtype" type="text" id="uploadtype" class="int" value="<%=uploadtype%>" size="40">��<span class="info_help help01">����gif/jpg/doc/xls/rar</span>
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
			<input name="Submit" type="submit" class="button45" id="Submit" value=" ���� ">
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