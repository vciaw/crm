<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
If Session("CRM_level") < 9 Then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
end if
	'��ȡgetֵ
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
            	<h1 class="titleh">ȫ������</h1>
                <div class="content">
                	
                <form action="?Action=Edit" method="post">
                    <div class="form-line">
                   	  <label class="st-label">��˾����</label>
					  <input name="title" type="text" id="title" class="int" value="<%=title%>" ><BR><span class="info_help help01">����XX��˾-�ͻ�����ϵͳ</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">ϵͳĿ¼</label>
					   <input name="SiteUrl" type="text" id="SiteUrl" class="int" value="<%=SiteUrl%>" size="5"><BR><span class="info_help help01">��"/"��β���� ��Ŀ¼��/ ����Ŀ¼��/ECRM/ </span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">���·��</label>
					   <input name="Skinurl" type="text" id="Skinurl" class="int" value="<%=Skinurl%>"><BR><span class="info_help help01">Ĭ�ϣ�Skin/default/</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">��ҳ����</label>
					   <input name="DataPageSize" type="text" id="DataPageSize" class="int" value="<%=DataPageSize%>" size="5"><BR><span class="info_help help01">ÿҳ��ʾ����������</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">������ʾ</label>
						<input name="YNalert" type="radio" class="noborder" value="1"<%IF ""&YNalert&"" = 1 then Response.Write("  checked") end if%>> �ǡ�
						<input name="YNalert" type="radio" class="noborder" value="0"<%IF ""&YNalert&"" = 0 then Response.Write("  checked") end if%>> ��
						<BR><span class="info_help help01">�����Ƿ񵯳�ȷ�ϴ��ڣ����ػ���ٲ�������</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">��������</label>
						<% = EasyCrm.getSelect("SelectData","Select_Type","CRTypeEnd",""&CRTypeEnd&"") %>
						<BR><span class="info_help help01">�ͻ��������̽���ʱ��״̬</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">��¼��־</label>
						<input name="YnUserLog" type="radio" class="noborder" value="1"<%IF ""&YnUserLog&""=1 then Response.Write("checked") end if%>> �ǡ�
						<input name="YnUserLog" type="radio" class="noborder" value="0"<%IF ""&YnUserLog&""=0 then Response.Write("checked") end if%>> ��
						<BR><span class="info_help help01">��¼��־�Ͳ�����¼</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">�������</label>
						<input name="YnDDNum" type="radio" class="noborder" value="1"<%IF ""&YnDDNum&""=1 then Response.Write("checked") end if%>> �Զ���
						<input name="YnDDNum" type="radio" class="noborder" value="0"<%IF ""&YnDDNum&""=0 then Response.Write("checked") end if%>> �ֶ���
						<BR><span class="info_help help01">�Զ���ʽΪ��DD20140606120000001</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">��ͬ���</label>
						<input name="YnHTNum" type="radio" class="noborder" value="1"<%IF ""&YnHTNum&""=1 then Response.Write("checked") end if%>> �Զ���
						<input name="YnHTNum" type="radio" class="noborder" value="0"<%IF ""&YnHTNum&""=0 then Response.Write("checked") end if%>> �ֶ���
						<BR><span class="info_help help01">�Զ���ʽΪ��HT20140606120000001</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">ɾ��ԭ��</label>
						<input name="YnDelReason" type="radio" class="noborder" value="1"<%IF ""&YnDelReason&""=1 then Response.Write("checked") end if%>> �ǡ�
						<input name="YnDelReason" type="radio" class="noborder" value="0"<%IF ""&YnDelReason&""=0 then Response.Write("checked") end if%>> ��
						<BR><span class="info_help help01">ɾ���ͻ������Ƿ���Ҫ��дԭ��</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">ת�Ʊ���</label>
						<input name="SaveOldUser" type="radio" class="noborder" value="1"<%IF ""&SaveOldUser&""=1 then Response.Write("checked") end if%>> �ǡ�
						<input name="SaveOldUser" type="radio" class="noborder" value="0"<%IF ""&SaveOldUser&""=0 then Response.Write("checked") end if%>> ��
						<BR><span class="info_help help01">�ͻ�ת�ƺ��Ƿ���ԭ��ҵ��Ա</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">�������</label>
						<input name="YNRecycler" type="radio" class="noborder" value="1"<%IF ""&YNRecycler&""=1 then Response.Write("checked") end if%>> �ǡ�
						<input name="YNRecycler" type="radio" class="noborder" value="0"<%IF ""&YNRecycler&""=0 then Response.Write("checked") end if%>> ��
						<BR><span class="info_help help01">��������ͻ��Ƿ���Ҫ���</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">�ͻ�Ψһ</label>
						<input name="ClientOnly1" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 1, 1) = 1 then%> checked<%end if%> > ��˾���� + 
						<input name="ClientOnly2" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 2, 1) = 1 then%> checked<%end if%> > ��ϵ�� + 
						<input name="ClientOnly3" type="checkbox" class="noborder" value="1"<%if mid(ClientOnly, 3, 1) = 1 then%> checked<%end if%> > �ֻ�
						<BR><span class="info_help help01">�жϿͻ�Ψһ�ı�׼</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">��������</label>
						<input name="SelectCharset" type="radio" class="noborder" value="1"<%IF ""&SelectCharset&""=1 then Response.Write("  checked") end if%>> �ǡ�
						<input name="SelectCharset" type="radio" class="noborder" value="0"<%IF ""&SelectCharset&""=0 then Response.Write("  checked") end if%>> ��
						<BR><span class="info_help help01">��ͬ�������������ܻ�������룬��ı�����</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">��������</label>
						<input name="Keeponline" type="radio" class="noborder" value="1"<%IF ""&Keeponline&""=1 then Response.Write("  checked") end if%>> �ǡ�
						<input name="Keeponline" type="radio" class="noborder" value="0"<%IF ""&Keeponline&""=0 then Response.Write("  checked") end if%>> ��
						<BR><span class="info_help help01">��ϱ����¼״̬����һ����ȫ����</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">Cookieֵ</label>
						<input name="CookieKey" type="text" id="CookieKey" class="int" value="<%=CookieKey%>">��
						<BR><span class="info_help help01">6-16λ����+��ĸ�����µ�¼��Ч</span>
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">�ϴ�����</label>
						<input name="uploadtype" type="text" id="uploadtype" class="int" value="<%=uploadtype%>">��
						<BR><span class="info_help help01">����gif/jpg/doc/xls/rar</span>
                    </div>
                    
                    <div class="form-line">
                    <input type="submit" name="button" id="button" value="&nbsp;�� ��&nbsp;" class="submit-button" />
                    <input type="reset" name="button" id="button2" value="&nbsp;�� ��&nbsp;" class="reset-button" />
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
