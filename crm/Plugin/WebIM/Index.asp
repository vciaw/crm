<!--#include file="../../data/conn.asp" --><!--#include file="../../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<%
Dim wconn,wMDBPath
set wrs=server.CreateObject("adodb.recordset")
Set wconn = Server.CreateObject("ADODB.Connection")
wMDBPath = Server.MapPath("DataBase/#WebIMdata.mdb")
wconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & wMDBPath

Dim UserName,UserEmail
UserName  = Session("CRM_name")
UID  = Session("CRM_uID")
UserEmail = EasyCrm.getNewItem("User","uAccount","'"&Session("CRM_account")&"'","uEmail")
Userlevel = Session("CRM_level")
	
'��һ�ε�¼�Զ������û���Ϣ�Ͷ�Ӧ�Ĺ�ϵ��
	wsql = "select * From [user] where UserName = '"&UserName&"' "
	wRs.Open wsql,wconn,1,1
	If wRs.RecordCount = 0 Then '���û�
		Set mRs=wconn.Execute("select max(UserID)+1 As maxuserid From [user]",1,1) '��ȡ���ID
		if mRs("maxuserid")<>"" then 
			wconn.execute ("insert into [user](UserName,UserPass,UserID,UserEmail,UserFace,UserSign,UserStatus,LastOnlineTime,UserGender,UserPower,SysCode) values('"&UserName&"','','"&mRs("maxuserid")&"','"&UserEmail&"','default.gif','','7','"&now()&"','1','2','1')")	'д���û���Ϣ
		else	
			wconn.execute ("insert into [user](UserName,UserPass,UserID,UserEmail,UserFace,UserSign,UserStatus,LastOnlineTime,UserGender,UserPower,SysCode) values('"&UserName&"','','10000','"&UserEmail&"','default.gif','','0','"&now()&"','1','0','1')")	'д���ʼ�˻�
		end if
		Response.Write "<script>location.href='index.asp';</script>" 
	else
		set ucRs=server.CreateObject("adodb.recordset")
		ucRs.Open "select * From [UserConfig] where UserID = "&wRs("UserID")&" ",wconn,1,1
		if ucRs.RecordCount = 0 then '���û�
		wconn.execute ("update [UserNum] set IsOK = 2 where Num = "&wRs("UserID")&" ") '�����û�
		wconn.execute ("insert into [UserConfig](UserID,DisType,OrderType,ChatSide,MsgSendKey,ShowFocus,MsgShowTime) values('"&wRs("UserID")&"','1','1','1','1','2','1')")	'�����û�������Ϣ
		end if
		ucRs.Close
	End If
		wconn.execute ("update [user] set UserEmail = '"&UserEmail&"',UserStatus = 7 where UserName = '"&UserName&"' ") '�����ʼ���ַ
		wconn.execute ("delete from [user] where UserName='' or UserName is null ")	'ɾ�������˻�
		wconn.execute ("delete from [UserMsg] where MsgAddTime <= #" & now()-30 & "# ") '��Ϣ����30��
	wRs.Close
	
	wRs.Open "Select * From [user] Where UserName <> '"&UserName&"' ",wconn,1,1
	Do While Not wRs.BOF And Not wRs.EOF
		Set uRs=wconn.Execute("select UserID From [user] where UserName = '"&UserName&"' ",1,1) '��ȡ��ǰ�û���UserID
		set ufRs=server.CreateObject("adodb.recordset")
		ufRs.Open "select * From [UserFriend] where UserID = "&uRs("UserID")&" and FriendID = "&wRs("UserID")&" ",wconn,1,1
		if ufRs.RecordCount = 0 then
		'ѭ��д���û���ϵ
		wconn.execute ("insert into [UserFriend](UserID,FriendID,GroupID,CustomName,IsBlocked) values('"&uRs("UserID")&"','"&wRs("UserID")&"','1','','2')")
		else
		end if
		ufRs.Close
		wRs.MoveNext
	Loop
	wRs.Close
	Set wRs = Nothing	
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<link href="styles/webim.css" type="text/css" rel="stylesheet" media="all">
<script type="text/javascript" src="js/webimhelper.js?v=102"></script>
<script type="text/javascript" src="js/webim.js?v=102"></script>
<!--[if IE 6]>
<script type="text/javascript" src="<%=SiteUrl&skinurl%>Js/fixpng.js"></script>
<script>DD_belatedPNG.fix('img,background');</script>
<![endif]-->
<title>EasyCrm ������</title>
</head>
<body onload= "javascript:TempChatMain() " style="background: #fff;">
<%if UserEmail<>"" then%>
<script type="text/javascript">
Other.SetCookie("stremail", "<%=UserEmail%>");Other.SetCookie("saveemail", "1");Other.SetCookie("savepass", "1");Other.SetCookie("autologin", "1");
</script>
<%else
	Response.Write "<script>alert('�����ַ����Ϊ��');location.href='../../system/User_info.asp?uid="&UID&"';</script>" 
end if%>
</body>
</html><% Set EasyCrm = nothing %>
