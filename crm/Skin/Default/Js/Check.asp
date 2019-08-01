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

 Response.ContentType="text/xml"
 item= Replace(Trim(request("item")),"'","") 
 if item<>"" then 
 sql="select top 10 cCompany from [client] where cCompany like '%"&item&"%' order by cid desc"
 rs.open sql,conn,1,1  
 str="<?xml version=""1.0"" encoding=""gb2312""?>"&vbnewline
  str=str&"<root>"&vbnewline
  If rs.RecordCount = 0 Then
   str=str&"<message id=""1"">"&vbnewline  
   str=str&"  <text>ÔÊĞíÂ¼Èë</text>"&vbnewline
   str=str&"</message>"&vbnewline
  else
   str=str&"<message id=""1"">"&vbnewline  
   str=str&"  <text>¹Ø±Õ</text>"&vbnewline
   str=str&"</message>"&vbnewline
  end if
 If rs.eof Then  
 Else
  i=1
  Do While Not rs.eof
   str=str&"<message id="""&i&""">"&vbnewline  
   str=str&"  <text>"&rs(0)&"</text>"&vbnewline
   str=str&"</message>"&vbnewline
  i=i+1
  rs.movenext
  loop
  End If  
  str=str&"</root>"
  rs.close
  response.write str
  end if
%>