<!--#include file="../data/config.asp" --><%
Response.Cookies(CookieKey)("CRM_account") = ""
Response.Cookies(CookieKey)("CRM_name") = ""
Response.Cookies(CookieKey)("CRM_uId") = ""
Response.Cookies(CookieKey)("CRM_level") = ""
Response.Cookies(CookieKey)("CRM_group") = ""
Response.Cookies(CookieKey)("CRM_qx") = ""
Response.Cookies(CookieKey)("CRM_MR") = ""
Response.Cookies(CookieKey)("CRM_Accsql") = ""
Response.Cookies(CookieKey)("Data_Source") = ""
Response.Cookies(CookieKey)("Data_User") = ""
Response.Cookies(CookieKey)("Data_Password") = ""
Response.Cookies(CookieKey)("Data_Catalog") = ""
Response.Cookies(CookieKey)("Data_MDBPath") = ""
Response.Cookies(CookieKey)("CRM_url") = ""
Response.Cookies(CookieKey)("CRM_LoginTime") = ""
Session.Abandon()
Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('login.asp');</script>"
Response.End
%>
