<%
set regex = new regexp 
regex.ignorecase = true 
regex.global = true 
regex.pattern = "mini 9.5|vx1000|lge |m800|e860|u940|ux840|compal|wireless| mobi|ahong|lg380|lgku|lgu900|lg210|lg47|lg920|lg840|lg370|sam-r|mg50|s55|g83|t66|vx400|mk99|d615|d763|el370|sl900|mp500|samu3|samu4|vx10|xda_|samu5|samu6|samu7|samu9|a615|b832|m881|s920|n210|s700|c-810|_h797|mob-x|sk16d|848b|mowser|s580|r800|471x|v120|rim8|c500foma:|160x|x160|480x|x640|t503|w839|i250|sprint|w398samr810|m5252|c7100|mt126|x225|s5330|s820|htil-g1|fly v71|s302|-x113|novarra|k610i|-three|8325rc|8352rc|sanyo|vx54|c888|nx250|n120|mtk |c5588|s710|t880|c5005|i;458x|p404i|s210|c5100|teleca|s940|c500|s590|foma|samsu|vx8|vx9|a1000|_mms|myx|a700|gu1100|bc831|e300|ems100|me701|me702m-three|sd588|s800|8325rc|ac831|mw200|brew |d88|htc\/|htc_touch|355x|m50|km100|d736|p-9521|telco|sl74|ktouch|m4u\/|me702|8325rc|kddi|phone|lg |sonyericsson|samsung|240x|x320vx10|nokia|sony cmd|motorola|up.browser|up.link|mmp|symbian|smartphone|midp|wap|vodafone|o2|pocket|kindle|mobile|psp|treo|iris|3g_t|windows ce|opera mobi|windows ce; smartphone;|windows ce; iemobile|ipod|iphone|android|opera mini|blackberry|palm os|palm|hiptop|avantgo|fennec|plucker|xiino|blazer|elaine|iris|3g_t|windows ce|opera mobi|windows ce; smartphone;|windows ce; iemobile" 

agent = request.ServerVariables("HTTP_USER_AGENT") & "" 
if agent <> "" then 
if not regex.test(agent) then

s= Server.MapPath("Install.asp")     
Set fso = CreateObject("Scripting.FileSystemObject")     
If fso.FileExists(s) Then     
Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('Install.asp');</script>" 
else 
Session("CRM_isMobile")=0
If Session("CRM_account") <> "" And Session("CRM_name") <> "" And IsNumeric(Session("CRM_level")) Then 
Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('main/index.asp');</script>"
else
Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('main/login.asp');</script>"
end if
End If     
Set fso = Nothing


else 

Session("CRM_isMobile")=1
If Session("CRM_account") <> "" And Session("CRM_name") <> "" And IsNumeric(Session("CRM_level")) Then 
Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('m/index.asp');</script>"
else
Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('m/login.asp');</script>"
end if

end if 
end if 
%>