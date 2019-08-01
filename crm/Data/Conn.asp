<!--#include file="Mssql.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="Md5.asp" -->
<!--#include file="Function.asp"-->
<!--#include file="CheckInput.asp"-->
<!--#include file="../lang/zh-cn/lang.asp"-->
<!--#include file="Config/Show_Listall.asp"-->
<!--#include file="Config/Show_Client.asp"-->
<!--#include file="Config/Must_Client.asp"-->
<!--#include file="Config/Show_Must_Linkmans.asp"-->
<!--#include file="Config/Show_Must_Records.asp"-->
<!--#include file="Config/Show_Must_Order.asp"-->
<!--#include file="Config/Show_Must_Order_Products.asp"-->
<!--#include file="Config/Show_Must_Products.asp"-->
<!--#include file="Config/Show_Must_Hetong.asp"-->
<!--#include file="Config/Show_Must_Service.asp"-->
<!--#include file="Config/Show_Must_Expense.asp"-->
<%	
if httpurl <> "login.asp" then
	If Session("CRM_account") <> "" And Session("CRM_name") <> "" And IsNumeric(Session("CRM_level")) Then 
		url = "../main/"  
	else
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('../index.asp');</script>"
	end if
end if
%>