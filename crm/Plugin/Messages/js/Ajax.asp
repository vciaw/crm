<!--#include file="../../../Data/Mssql.asp"-->
<%
Response.Addheader "Content-Type","text/html; charset=gb2312" 
Dim conn,connstr,MDBPath
Accsql=""&Accsql&""
set rs=server.CreateObject("adodb.recordset")
Set conn = Server.CreateObject("ADODB.Connection")
MDBPath = Server.MapPath(""&DataPath&"")
if Accsql="0" then
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MDBPath
elseif Accsql="1" then
	Conn.open "Provider=SQLOLEDB;Data Source="&Data_Source&";User ID="&Data_User&";Password="&Data_Password&";Initial Catalog="&Data_Catalog&""
end if

Dim action
action = Trim(Request("action"))

Select Case action
Case "Area"
    Call Area()
Case "Trade"
    Call Trade()
End Select

Sub Area()
Dim rs,sql
AreaData = Request.QueryString("AreaData")

	Response.Write"<select name=""Squares"" class=""int"" onchange=""getSquare(options[selectedIndex]);"">"
	Response.Write"<option value="""">请选择</option>"
	sql = "Select aName From AreaData Where aFId = '"&AreaData&"' "
	sql = sql &" Order by aId asc"
	Set rs = Conn.Execute(sql)
	If (rs.EOF And rs.BOF) Then
		Response.Write"<option value="""">暂无小类</option>"
	Else
		Do While Not rs.EOF
			Response.Write"<option value="""&rs(0)&""">"&rs(0)&"</option>"
		rs.MoveNext
		Loop
	End If
	Response.Write "</select>"
	Set rs = Nothing

End Sub

Sub Trade()
Dim rs,sql
TradeData = Request.QueryString("TradeData")

	Response.Write"<select name=""Strades"" class=""int"" onchange=""getStrade(options[selectedIndex]);"">"
	Response.Write"<option value="""">请选择</option>"
	sql = "Select pClassname From ProductClass Where pClassFid = '"&TradeData&"' "
	sql = sql &" Order by pClassId asc"
	Set rs = Conn.Execute(sql)
	If (rs.EOF And rs.BOF) Then
		Response.Write"<option value="""">暂无小类</option>"
	Else
		Do While Not rs.EOF
			Response.Write"<option value="""&rs(0)&""">"&rs(0)&"</option>"
		rs.MoveNext
		Loop
	End If
	Response.Write "</select>"
	Set rs = Nothing

End Sub

%>