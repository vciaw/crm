﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file = "config.asp"-->
<!--#include file = "function.asp"-->
<!--#include file = "cmd.asp"-->
<%
	Response.Expires = WebCachTime
	Response.ContentType = "text/xml"
	Response.Charset="utf-8"
	id = Request.QueryString("id")
	If id = "" Then Response.End
	Response.Write("<?xml version=""1.0"" encoding=""utf-8""?>")
	Response.Write("<list>")
	Call CheckLogin()
	Call DataBegin()
	sql = "select a.isblocked,a.groupid as gid,a.customname as cname,b.* from userfriend a inner join [user] b on a.friendid = b.userid where b.userid ="&id&" and a.userid = "&Session("userid")
	oRs.Open sql,oConn,1,1
	If Not(oRs.Bof And oRs.Eof) Then
		Response.Write("<item>")
		Call OutNode("f",oRs("userface"))
		Call OutNode("id",oRs("userid"))
		Call OutNode("n",oRs("username"))
		Call OutNode("e",oRs("useremail"))
		Call OutNode("sn",oRs("usersign"))
		Call OutNode("s",oRs("userstatus"))
		Call OutNode("g",oRs("gid"))
		Call OutNode("b",oRs("isblocked"))
		Call OutNode("cn",oRs("cname"))
		Call OutNode("u",oRs("usergender"))
		Response.Write("</item>")
	End If
	oRs.Close()
	Set oRs = Nothing
	Call DataEnd()
	Response.Write("</list>")
%>