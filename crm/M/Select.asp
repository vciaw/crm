<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
If Session("CRM_level") < 9 Then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
end if
	'获取get值
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	if otype="" then otype="Select_Type"
%><%=Header%>
<script language="JavaScript">
<!--
function checkInput(o)
{
    var oo = eval("document.all." + o);
    var num = oo.length;
    for(var i=0;i<num;i++){
	    if(oo[i].value == ""){
		    alert("不能为空！");
			oo[i].focus();
			return false
			break;
		}
	}
}
-->
</script>
<!-- start header -->
    <div id="header">
         <a href="System.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
         <a style="cursor:pointer" class="button list"><img src="img/list-button.png" width="15" height="16" alt="icon" /></a>
         <a onClick=window.location.href="javascript:window.location.reload();" class="button create"><img src="img/reload-button.png" width="15" height="16" alt="icon" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start page -->
    <div class="page">
	<div class="simplebox">
	
<div class="listbox" style="margin-bottom:10px;border:1px solid #C1D6E6;">
<input type="button" class="reset-button" value="客户类型" onclick="location.href='?otype=Select_Type'" style="cursor:pointer" />
<input type="button" class="reset-button" value="客户来源" onclick="location.href='?otype=Select_Source'" style="cursor:pointer" />
<input type="button" class="reset-button" value="客户级别" onclick="location.href='?otype=Select_Star'" style="cursor:pointer" />
<input type="button" class="reset-button" value="所属职位" onclick="location.href='?otype=Select_Zhiwei'" style="cursor:pointer" />
<input type="button" class="reset-button" value="跟单类型" onclick="location.href='?otype=Select_Records'" style="cursor:pointer" />
<input type="button" class="reset-button" value="合同分类" onclick="location.href='?otype=Select_Hetong'" style="cursor:pointer" />
<input type="button" class="reset-button" value="售后类型" onclick="location.href='?otype=Select_Service'" style="cursor:pointer" />
<input type="button" class="reset-button" value="收入类型" onclick="location.href='?otype=Select_ExpenseIN'" style="cursor:pointer" />
<input type="button" class="reset-button" value="支出类型" onclick="location.href='?otype=Select_ExpenseOUT'" style="cursor:pointer" />
<input type="button" class="reset-button" value="文件分类" onclick="location.href='?otype=Select_SoftClass'" style="cursor:pointer" />
<input type="button" class="reset-button" value="公文分类" onclick="location.href='?otype=Select_NoticeClass'" style="cursor:pointer" />
<input type="button" class="reset-button" value="性别" onclick="location.href='?otype=Select_Sex'" style="cursor:pointer" />
<input type="button" class="reset-button" value="是/否" onclick="location.href='?otype=Select_YN'" style="cursor:pointer" />
</div>
<%

Select Case action
	Case "add"
		Call addOrEdit()
	Case "save"
		Call saveData()
	Case "edit"
		Call addOrEdit()
	Case "restore"
		Call restore()
	Case "delete"
		Call deleteData()
	Case Else
		Call addOrEdit()
End Select
	
Sub addOrEdit()
    Dim otypedata,otypedataOld,strOut,strAction
	If action = "edit" Then
	    Dim rs
		otypedataOld = Trim(Request("otypedataOld"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [SelectData] Where "&otype&" = '" & otypedataOld & "'",conn,3,1
		If rs.RecordCount = 1 Then
			otypedata = rs(""&otype&"")
		End If
		rs.Close
		Set rs = Nothing
		strOut = ""&L_Edit&""
		strAction = "?action=restore&otype="&otype&""
	Else
	    strOut = ""&L_Add&""
		strAction = "?action=save&otype="&otype&""
	End If		    
%>

            	<h1 class="titleh"><% = strOut %></h1>
                <div class="content">
				<form name="menuForm" action="<% = strAction %>" method="post" onSubmit="return checkInput('menuForm');">
                    <div class="form-line">
                   	  <label class="st-label">参数</label>
						<% If action = "edit" Then %>
							<input name="otypedataOld" type="hidden" id="otypedataOld" value="<% = otypedataOld %>">
						<% End If %>
						<input name="otypedata" type="text" class="int" id="otypedata" size="30" maxlength="16" value="<% = otypedata %>">
                    </div>
                    
                    <div class="form-line">
                    <input type="submit" name="button" id="button" value="&nbsp;保 存&nbsp;" class="submit-button" />
					<% If action = "edit" Then %><input name="Back" type="button" class="reset-button" id="Back" value="&nbsp;取 消&nbsp;" onClick="location.href='?otype=<%=otype%>';"><% End If %>
                    </div>
                </form>
                </div>
<%
End Sub
%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col ><col width="140">
						<tbody> 
                        <tr> 
							<td >类型</td> 
							<td>管理</td> 
                        </tr> 
						<% = list() %>
                    </table> 
			</div>
		<%=Footer%>
            
<%
Function list()
    Dim strToPrint
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select * From [SelectData] where "&otype&" <>'' and "&otype&" <> 'null'  ",conn,3,1
    Do While Not rs.BOF And Not rs.EOF
    	strToPrint = strToPrint & "        <tr>" & VBCrlf
    	strToPrint = strToPrint & "          <td><a href=""?action=edit&otype="&otype&"&otypedataOld=" & rs(""&otype&"") & """>" & rs(""&otype&"") & "</a></td>" & VBCrlf
    	strToPrint = strToPrint & "          <td><input type=""button"" class=""submit-button"" value="""&L_Edit&""" title="""&L_Edit&""" onClick=""window.location.href='?action=edit&otype="&otype&"&otypedataOld=" & rs(""&otype&"") & "'"" /> " & VBCrlf
    	strToPrint = strToPrint & "          <input type=""button"" class=""reset-button"" value="""&L_Del&"""  title="""&L_Del&""" onClick=""window.location.href='?action=delete&otype="&otype&"&otypedataOld=" & rs(""&otype&"") & "'"" onClick=""return confirm('"&Alert_del_YN&"');"" /></td>" & VBCrlf
    	strToPrint = strToPrint & "          </tr>" & VBCrlf
    	rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
	list = strToPrint
End Function

Sub saveData() '保存数据
    Dim otypedata
	otypedata = Trim(Request.Form("otypedata"))
	If otypedata = "" Then
        Response.Write("<script>alert("""&alert03&""");history.back(1);</script>")
		Exit Sub
	End If
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select  * From [SelectData] Where "&otype&" = '" & otypedata & "'",conn,3,2
	If rs.RecordCount > 0 Then
        Response.Write("<script>alert("""&alert02&""");history.back(1);</script>")
		rs.Close
		Set rs = Nothing
		Exit Sub
	Else
	    rs.AddNew
		rs(""&otype&"") = otypedata
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("?otype="&otype&"")
	End If
End Sub

Sub restore() '修改数据
    Dim otypedataOld,otypedata
	otypedataOld = Trim(Request.Form("otypedataOld"))
	otypedata = Trim(Request.Form("otypedata"))
	If otypedataOld = "" Or otypedata = "" Then
        Response.Write("<script>alert("""&alert03&""");history.back(1);</script>")
		Exit Sub
	End If
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [SelectData] Where "&otype&" <> '" & otypedataOld & "'",conn,3,1
	Do While Not rs.BOF And Not rs.EOF
	    If rs(""&otype&"") = otypedata Then
        Response.Write("<script>alert("""&alert02&""");history.back(1);</script>")
		    rs.Close
		    Set rs = Nothing
		    Exit Sub
		End If
		rs.MoveNext
	Loop
	rs.Close
	
	rs.Open "Select * From [SelectData] Where "&otype&" = '" & otypedataOld & "'",conn,3,2
	If rs.RecordCount = 1 Then
	    otypedataOld = rs(""&otype&"")
		rs(""&otype&"") = otypedata
		rs.Update
	End If
	Set rs = Nothing
	Response.Redirect("?otype="&otype&"")
End Sub

Sub deleteData() '删除数据
    Dim otypedataOld
	otypedataOld = Trim(Request("otypedataOld"))
        'Response.Write("<script>alert("""&otype&""");</script>")
		'Response.end
	If otypedataOld = "" Then
        Response.Write("<script>alert("""&alert03&""");history.back(1);</script>")
		Exit Sub
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [SelectData] Where "&otype&" = '" & otypedataOld & "'",conn,3,2
	If rs.RecordCount > 0 Then
		rs.Delete
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
	Response.Redirect("?otype="&otype&"")
End Sub

%>
    <div class="clear"></div>
    </div>
    <!-- end page -->
<script type="text/javascript" src="js/frame.js"></script>
</body>
</html>
