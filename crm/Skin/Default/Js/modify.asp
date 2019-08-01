<!--#include file="../../../data/mssql.asp"-->
<% 
Dim conn,MDBPath
Accsql=""&Accsql&""
set rs=server.CreateObject("adodb.recordset")
Set conn = Server.CreateObject("ADODB.Connection")
MDBPath = Server.MapPath(""&Data_MDBPath&"")
if Accsql="0" then
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MDBPath
elseif Accsql="1" then
	Conn.open "Provider=SQLOLEDB;Data Source="&Data_Source&";User ID="&Data_User&";Password="&Data_Password&";Initial Catalog="&Data_Catalog&""
end if

'id为记录ID,v为修改后的值,t1为表名,fname为字段名,gname为ID字段名，自动编号的那个字段的字段名
id=cstr(request("id"))
v=cstr(request("v"))
t1=request("t1")
fname=request("fname")
gname=request("gname")
conn.execute "update "&t1&" set "&fname&"='"&v&"' where "&gname&"="&id
response.write 1
%>
