<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
If Session("CRM_level") < 9 Then
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">window.top.location.replace('index.asp');</script>"
end if
	'获取get值
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
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
         	<a href="#" class="button list"><img src="img/create.png" width="16" height="16" alt="icon"/></a>
         <a onClick=window.location.href="javascript:window.location.reload();" class="button create"><img src="img/reload-button.png" width="15" height="16" alt="icon" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
    <!-- start page -->
    <div class="page">
	<div class="simplebox">
                <div class="content listbox" style="margin-bottom:10px;">
				<form name="Save" action="?action=SaveAdd" method="post">
                    <div class="form-line">
                   	  <label class="st-label">数据表</label>
						<select name="cTable" id="cTable" ><option value="">请选择</option><option value="Client">客户档案</option><option value="Records">跟单记录</option><option value="Order">订单记录</option><option value="Hetong">合同记录</option><option value="Service">售后记录</option></select>　
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">显示名</label>
						<input name="cTitle" type="text" id="cTitle" class="int" size="20" /> 例 : 开户行
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">字段名</label>
						<input name="cName" type="text" id="cName" class="int" size="20" /> 例 : BANK
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">字段类型</label>
						<select name="cType" id="cType" ><option value="">请选择</option><option value="text">文本</option><option value="time">时间日期</option><option value="select">下拉框</option><option value="checkbox">多选框</option><option value="radio">单选框</option></select>　
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">表单长度</label>
						<input name="cWidth" type="text" id="cWidth" class="int" size="20" /> 单位 : PX　
                    </div>
					
                    <div class="form-line">
                   	  <label class="st-label">备注</label>
						<textarea name="cContent" id="cContent" class="int" style="width:80%;"></textarea>
                    </div>
                    
                    <div class="form-line">
                    <input type="submit" name="button" id="button" value="&nbsp;保 存&nbsp;" class="submit-button" />
                    </div>
                </form>
                </div>

            	<h1 class="titleh">自定义字段</h1>
					<table class="tabledata"> 
						<tbody> 
                        <tr> 
							<td >数据表</td> 
							<td>字段名</td> 
							<td>类型</td> 
							<td>管理</td> 
                        </tr> 
						<%
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.Open "Select * From [CustomField] order by Id asc ",conn,1,1
							Do While Not rs.BOF And Not rs.EOF
						%>
								<tr class="tr">
									<td class="td_l_c">[<%=rs("cTable")%>]</td>
									<td class="td_l_c"><%=rs("cTitle")%></td>
									<td class="td_l_c"><%=rs("cType")%></td>
									<td class="td_l_c"><a onClick=window.location.href="?action=delete&Id=<%=rs("Id")%>" style="cursor:pointer" />删除</a></td>
								</tr>
						<%
							rs.MoveNext
							Loop
							rs.Close
							Set rs = Nothing
						%>
                    </table> 
			</div>
		<%=Footer%>
            
<%

Select Case action
Case "SaveAdd" '添加
    Call SaveAdd()
Case "delete" '删除
    Call deleteData()
End Select

Sub SaveAdd() '删除
		cTable = Request("cTable")
		cTitle = Request("cTitle")
		cName = Request("cName")
		cTypeS = Request("cType")
		cWidth = Request("cWidth")
		cContent = Request("cContent")
		If cName = "" Then
			Response.Write("<script>alert("""&alert01&""");history.back(1);</script>")
			Exit Sub
		End If
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [CustomField] ",conn,3,2
		rs.AddNew
		rs("cTable") = cTable
		rs("cTitle") = cTitle
		rs("cName") = cName
		rs("cType") = cTypeS
		rs("cWidth") = cWidth
		rs("cContent") = cContent
    	rs.Update
    	rs.Close
    	Set rs = Nothing
	Response.Redirect("?")
End Sub

Sub deleteData() '删除
	Id = Request("Id")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [CustomField] Where Id = " & Id,conn,3,2
	If rs.RecordCount > 0 Then
		rs.Delete
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
	Response.Redirect("?")
End Sub

%>
    <div class="clear"></div>
    </div>
    <!-- end page -->
<script type="text/javascript" src="js/frame.js"></script>
</body>
</html>
