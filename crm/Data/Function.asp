<% 
Response.Addheader "Content-Type","text/html; charset=gb2312"
Dim action,subAction,arrList,otype
Dim strNormal,strAdmin,strCounter,strToPrint
Dim conn,connstr,MDBPath
Accsql=""&Accsql&""  ' 0 为access数据库 ，1 为mssql数据库
set rs=server.CreateObject("adodb.recordset")
Set conn = Server.CreateObject("ADODB.Connection")
MDBPath = Server.MapPath(""&Data_MDBPath&"")
if Accsql="0" then
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MDBPath
	'conn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ="+MDBPath
elseif Accsql="1" then
	Conn.open "Provider=SQLOLEDB;Data Source="&Data_Source&";User ID="&Data_User&";Password="&Data_Password&";Initial Catalog="&Data_Catalog&""
end if

if Keeponline = 1 then
	Session("CRM_account") = Request.Cookies(CookieKey)("CRM_account")
	Session("CRM_name") = Request.Cookies(CookieKey)("CRM_name")
	Session("CRM_uId") = Request.Cookies(CookieKey)("CRM_uId")
	Session("CRM_level") = Request.Cookies(CookieKey)("CRM_level")
	Session("CRM_group") = Request.Cookies(CookieKey)("CRM_group")
	Session("CRM_qx") = Request.Cookies(CookieKey)("CRM_qx")
	Session("CRM_MR") = Request.Cookies(CookieKey)("CRM_MR")
	Session("CRM_Accsql") = Request.Cookies(CookieKey)("CRM_Accsql")
	Session("Data_Source") = Request.Cookies(CookieKey)("Data_Source")
	Session("Data_User") = Request.Cookies(CookieKey)("Data_User")
	Session("Data_Password") = Request.Cookies(CookieKey)("Data_Password")
	Session("Data_Catalog") = Request.Cookies(CookieKey)("Data_Catalog")
	Session("Data_MDBPath") = Request.Cookies(CookieKey)("Data_MDBPath")
	Session("CRM_url") = Request.Cookies(CookieKey)("CRM_url")
end if 
	
if Session("CRM_level")<>"" then
	arrUser = getUserList(Session("CRM_level"),Session("CRM_group"),Session("CRM_MR"))
end if

'路径过滤
Dim url1,url2,url3,httpurl
url1=Request.Servervariables("url")
url2=InstrRev(url1,"/")
url3=len(url1)
httpurl=Right(url1,url3-url2)

'权限范围
Function getUserList(intLevel,intGroup,inManagerange)
    Dim rs,strUserList
	Set rs = Server.CreateObject("ADODB.Recordset")
	if intLevel<>"" and intGroup<>"" then
	arrManagerange = Replace(Replace(inManagerange," ",""),",", "','")
	rs.Open "Select * From [user] where uName In ( Select uName From [user] Where uLevel < "&Session("CRM_level")&" And uGroup = "&Session("CRM_group")&" or uName in ( '"&arrManagerange&"' ) )",conn,1,1
	else
	Response.write"<script>location.href=""../main/login.asp"";</script>"
	end if
			strUserList = "'"&Session("CRM_name")&"'" '添加自己
	Do While Not rs.BOF And Not rs.EOF
		if rs("uName") <> Session("CRM_name") then  '跳过自己
		    strUserList = strUserList & ",'" & rs("uName") & "'"
		end if
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	getUserList = strUserList
End Function

'防SQL注入，防止外部提交
Dim GetFlag
Dim ErrorSql
Dim RequestKey
Dim ForI
ErrorSql = "'~;~(~)~exec~update~*~%~chr~mid~master~truncate~char~declare~srcipt"
ErrorSql = split(ErrorSql,"~")
If Request.ServerVariables("REQUEST_METHOD")="GET" Then
GetFlag=True
Else
GetFlag=False
End If
If GetFlag Then
For Each RequestKey In Request.QueryString
For ForI=0 To Ubound(ErrorSql)
If Instr(LCase(Request.QueryString(RequestKey)),ErrorSql(ForI))<>0 Then
response.write "<script>location.href=""index.asp"";</script>"
Response.End
End If
Next
Next
End If
%>