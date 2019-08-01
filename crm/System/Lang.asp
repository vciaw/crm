<!--#include file="../data/conn.asp" -->
<%
	'获取get值
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	tipinfo = 	Request.QueryString("tipinfo")
	if otype="" then otype="Client"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>
<body style="padding-top:35px;">

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：系统管理 > 语言包</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
			<input type="button" class="button_top_help" value=" " title="帮助说明" onclick='Lang_Help()' style="cursor:pointer" />
        </td>
	</tr>
</table>
<script>function Lang_Help() {art.dialog({ title: '帮助说明',icon: 'question', content: '仅提供【客户数据表】及【附属表】的字段显示修改<BR>如有高级修改需求，请用记事本编辑以下文件<BR>【/lang/zh-cn/lang.asp】',drag: false,resize: false}); };</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
            <div class="MenuboxS">
              <ul>
                <li <%if otype="Client" or otype="" then%>class="hover"<%end if%>><span><a href="?otype=Client">客户档案</a></span></li>
                <li <%if otype="Linkmans" then%>class="hover"<%end if%>><span><a href="?otype=Linkmans">联系人</a></span></li>
                <li <%if otype="Records" then%>class="hover"<%end if%>><span><a href="?otype=Records">跟单记录</a></span></li>
                <li <%if otype="Order" then%>class="hover"<%end if%>><span><a href="?otype=Order">订单记录</a></span></li>
                <li <%if otype="Hetong" then%>class="hover"<%end if%>><span><a href="?otype=Hetong">合同记录</a></span></li>
                <li <%if otype="Service" then%>class="hover"<%end if%>><span><a href="?otype=Service">售后记录</a></span></li>
                <li <%if otype="Expense" then%>class="hover"<%end if%>><span><a href="?otype=Expense">费用记录</a></span></li>
                <li <%if otype="File" then%>class="hover"<%end if%>><span><a href="?otype=File">附件记录</a></span></li>
                <li <%if otype="ManageLog" then%>class="hover"<%end if%>><span><a href="?otype=ManageLog">操作记录</a></span></li>
              </ul> 
            </div>
		</td>
	</tr>
</table>
<%
if tipinfo<>"" then
	Response.Write("<script>art.dialog({title: '提示',time: 1,icon: 'warning',content: '"&tipinfo&"'});</script>")
end if

If otype="Client" then '读取客户字段
%>
<form action="?otype=Client&action=SaveClient" method="post">
<table width="100%" border="0" cellpadding="0" cellspacing="0" id="TableMain">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdt10"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cId</td>
			<td class="td_l_l"> <input name="L_Client_cId" type="text" id="L_Client_cId" class="int" value="<%=L_Client_cId%>" size="60"></td>
			<td class="td_l_l">客户编号</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cDate</td>
			<td class="td_l_l"> <input name="L_Client_cDate" type="text" id="L_Client_cDate" class="int" value="<%=L_Client_cDate%>" size="60"></td>
			<td class="td_l_l">录入日期</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cCompany</td>
			<td class="td_l_l"> <input name="L_Client_cCompany" type="text" id="L_Client_cCompany" class="int" value="<%=L_Client_cCompany%>" size="60"></td>
			<td class="td_l_l">公司名称</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cArea</td>
			<td class="td_l_l"> <input name="L_Client_cArea" type="text" id="L_Client_cArea" class="int" value="<%=L_Client_cArea%>" size="60"></td>
			<td class="td_l_l">省份</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cSquare</td>
			<td class="td_l_l"> <input name="L_Client_cSquare" type="text" id="L_Client_cSquare" class="int" value="<%=L_Client_cSquare%>" size="60"></td>
			<td class="td_l_l">地区</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cAddress</td>
			<td class="td_l_l"> <input name="L_Client_cAddress" type="text" id="L_Client_cAddress" class="int" value="<%=L_Client_cAddress%>" size="60"></td>
			<td class="td_l_l">详细地址</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cZip</td>
			<td class="td_l_l"> <input name="L_Client_cZip" type="text" id="L_Client_cZip" class="int" value="<%=L_Client_cZip%>" size="60"></td>
			<td class="td_l_l">邮编</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cLinkman</td>
			<td class="td_l_l"> <input name="L_Client_cLinkman" type="text" id="L_Client_cLinkman" class="int" value="<%=L_Client_cLinkman%>" size="60"></td>
			<td class="td_l_l">主要联系人</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cZhiwei</td>
			<td class="td_l_l"> <input name="L_Client_cZhiwei" type="text" id="L_Client_cZhiwei" class="int" value="<%=L_Client_cZhiwei%>" size="60"></td>
			<td class="td_l_l">职位</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cMobile</td>
			<td class="td_l_l"> <input name="L_Client_cMobile" type="text" id="L_Client_cMobile" class="int" value="<%=L_Client_cMobile%>" size="60"></td>
			<td class="td_l_l">手机号码</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cTel</td>
			<td class="td_l_l"> <input name="L_Client_cTel" type="text" id="L_Client_cTel" class="int" value="<%=L_Client_cTel%>" size="60"></td>
			<td class="td_l_l">联系电话</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cFax</td>
			<td class="td_l_l"> <input name="L_Client_cFax" type="text" id="L_Client_cFax" class="int" value="<%=L_Client_cFax%>" size="60"></td>
			<td class="td_l_l">传真号码</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cHomepage</td>
			<td class="td_l_l"> <input name="L_Client_cHomepage" type="text" id="L_Client_cHomepage" class="int" value="<%=L_Client_cHomepage%>" size="60"></td>
			<td class="td_l_l">企业网站</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cEmail</td>
			<td class="td_l_l"> <input name="L_Client_cEmail" type="text" id="L_Client_cEmail" class="int" value="<%=L_Client_cEmail%>" size="60"></td>
			<td class="td_l_l">电子邮件</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cTrade</td>
			<td class="td_l_l"> <input name="L_Client_cTrade" type="text" id="L_Client_cTrade" class="int" value="<%=L_Client_cTrade%>" size="60"></td>
			<td class="td_l_l">产品大类</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cStrade</td>
			<td class="td_l_l"> <input name="L_Client_cStrade" type="text" id="L_Client_cStrade" class="int" value="<%=L_Client_cStrade%>" size="60"></td>
			<td class="td_l_l">产品小类</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cType</td>
			<td class="td_l_l"> <input name="L_Client_cType" type="text" id="L_Client_cType" class="int" value="<%=L_Client_cType%>" size="60"></td>
			<td class="td_l_l">客户类型</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cStart</td>
			<td class="td_l_l"> <input name="L_Client_cStart" type="text" id="L_Client_cStart" class="int" value="<%=L_Client_cStart%>" size="60"></td>
			<td class="td_l_l">客户级别</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cSource</td>
			<td class="td_l_l"> <input name="L_Client_cSource" type="text" id="L_Client_cSource" class="int" value="<%=L_Client_cSource%>" size="60"></td>
			<td class="td_l_l">客户来源</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cInfo</td>
			<td class="td_l_l"> <input name="L_Client_cInfo" type="text" id="L_Client_cInfo" class="int" value="<%=L_Client_cInfo%>" size="60"></td>
			<td class="td_l_l">主营项目</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cBeizhu</td>
			<td class="td_l_l"> <input name="L_Client_cBeizhu" type="text" id="L_Client_cBeizhu" class="int" value="<%=L_Client_cBeizhu%>" size="60"></td>
			<td class="td_l_l">备注其它</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cUser</td>
			<td class="td_l_l"> <input name="L_Client_cUser" type="text" id="L_Client_cUser" class="int" value="<%=L_Client_cUser%>" size="60"></td>
			<td class="td_l_l">业务员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cGroup</td>
			<td class="td_l_l"> <input name="L_Client_cGroup" type="text" id="L_Client_cGroup" class="int" value="<%=L_Client_cGroup%>" size="60"></td>
			<td class="td_l_l">部门</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cOldUser</td>
			<td class="td_l_l"> <input name="L_Client_cOldUser" type="text" id="L_Client_cOldUser" class="int" value="<%=L_Client_cOldUser%>" size="60"></td>
			<td class="td_l_l">申请人</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cLastUpdated</td>
			<td class="td_l_l"> <input name="L_Client_cLastUpdated" type="text" id="L_Client_cLastUpdated" class="int" value="<%=L_Client_cLastUpdated%>" size="60"></td>
			<td class="td_l_l">最后更新时间</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cShare</td>
			<td class="td_l_l"> <input name="L_Client_cShare" type="text" id="L_Client_cShare" class="int" value="<%=L_Client_cShare%>" size="60"></td>
			<td class="td_l_l">是否共享</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cRNextTime</td>
			<td class="td_l_l"> <input name="L_Client_cRNextTime" type="text" id="L_Client_cRNextTime" class="int" value="<%=L_Client_cRNextTime%>" size="60"></td>
			<td class="td_l_l">下次跟进</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cOEDate</td>
			<td class="td_l_l"> <input name="L_Client_cOEDate" type="text" id="L_Client_cOEDate" class="int" value="<%=L_Client_cOEDate%>" size="60"></td>
			<td class="td_l_l">交付订单</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cHEdate</td>
			<td class="td_l_l"> <input name="L_Client_cHEdate" type="text" id="L_Client_cHEdate" class="int" value="<%=L_Client_cHEdate%>" size="60"></td>
			<td class="td_l_l">合同到期</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cHMoney</td>
			<td class="td_l_l"> <input name="L_Client_cHMoney" type="text" id="L_Client_cHMoney" class="int" value="<%=L_Client_cHMoney%>" size="60"></td>
			<td class="td_l_l">总金额</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cHOwed</td>
			<td class="td_l_l"> <input name="L_Client_cHOwed" type="text" id="L_Client_cHOwed" class="int" value="<%=L_Client_cHOwed%>" size="60"></td>
			<td class="td_l_l">总欠款</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Client_cSNum</td>
			<td class="td_l_l"> <input name="L_Client_cSNum" type="text" id="L_Client_cSNum" class="int" value="<%=L_Client_cSNum%>" size="60"></td>
			<td class="td_l_l">售后次数</td>
        </tr>
      </table>
    </td> 
  </tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdb10 "> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
			<tr> 
				<td class="td_r_l" style="border-top:0;" COLSPAN="6"><input name="Submit" type="submit" class="button45" id="Submit" value=" <%=L_Submit%> "></td>
			</tr>
		</table>
	</td>
  </tr>
</table>
</form>
<% '保存客户字段
	If action="SaveClient" then
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "'客户档案表 Client" & VbCrLf
		TempStr = TempStr & "L_Client_cId="& Chr(34) & replace(Trim(Request.Form("L_Client_cId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cDate="& Chr(34) & replace(Trim(Request.Form("L_Client_cDate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cCompany="& Chr(34) & replace(Trim(Request.Form("L_Client_cCompany")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cArea="& Chr(34) & replace(Trim(Request.Form("L_Client_cArea")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cSquare="& Chr(34) & replace(Trim(Request.Form("L_Client_cSquare")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cAddress="& Chr(34) & replace(Trim(Request.Form("L_Client_cAddress")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cZip="& Chr(34) & replace(Trim(Request.Form("L_Client_cZip")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cLinkman="& Chr(34) & replace(Trim(Request.Form("L_Client_cLinkman")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cZhiwei="& Chr(34) & replace(Trim(Request.Form("L_Client_cZhiwei")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cMobile="& Chr(34) & replace(Trim(Request.Form("L_Client_cMobile")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cTel="& Chr(34) & replace(Trim(Request.Form("L_Client_cTel")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cFax="& Chr(34) & replace(Trim(Request.Form("L_Client_cFax")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cHomepage="& Chr(34) & replace(Trim(Request.Form("L_Client_cHomepage")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cEmail="& Chr(34) & replace(Trim(Request.Form("L_Client_cEmail")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cTrade="& Chr(34) & replace(Trim(Request.Form("L_Client_cTrade")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cStrade="& Chr(34) & replace(Trim(Request.Form("L_Client_cStrade")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cType="& Chr(34) & replace(Trim(Request.Form("L_Client_cType")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cStart="& Chr(34) & replace(Trim(Request.Form("L_Client_cStart")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cSource="& Chr(34) & replace(Trim(Request.Form("L_Client_cSource")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cInfo="& Chr(34) & replace(Trim(Request.Form("L_Client_cInfo")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cBeizhu="& Chr(34) & replace(Trim(Request.Form("L_Client_cBeizhu")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cUser="& Chr(34) & replace(Trim(Request.Form("L_Client_cUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cGroup="& Chr(34) & replace(Trim(Request.Form("L_Client_cGroup")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cOldUser="& Chr(34) & replace(Trim(Request.Form("L_Client_cOldUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cLastUpdated="& Chr(34) & replace(Trim(Request.Form("L_Client_cLastUpdated")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cShare="& Chr(34) & replace(Trim(Request.Form("L_Client_cShare")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cRNextTime="& Chr(34) & replace(Trim(Request.Form("L_Client_cRNextTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cOEDate="& Chr(34) & replace(Trim(Request.Form("L_Client_cOEDate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cHEdate="& Chr(34) & replace(Trim(Request.Form("L_Client_cHEdate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cHMoney="& Chr(34) & replace(Trim(Request.Form("L_Client_cHMoney")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cHOwed="& Chr(34) & replace(Trim(Request.Form("L_Client_cHOwed")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Client_cSNum="& Chr(34) & replace(Trim(Request.Form("L_Client_cSNum")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "%" & chr(62) & VbCrLf
			ADODB_SaveToFile TempStr,"../Lang/zh-cn/Client.asp"
		If GBL_CHK_TempStr = "" Then
			if ""&YNalert&"" = 1 then
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&alert2&"';</script>")
			else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"';</script>")
			end if
		Else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&GBL_CHK_TempStr&"';</script>")
		End If
	End if
%>
<%
elseIf otype="Linkmans" then '读取联系人字段
%>
<form action="?otype=Linkmans&action=SaveLinkmans" method="post">
<table width="100%" border="0" cellpadding="0" cellspacing="0" id="TableMain">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdt10"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lId</td>
			<td class="td_l_l"> <input name="L_Linkmans_lId" type="text" id="L_Linkmans_lId" class="int" value="<%=L_Linkmans_lId%>" size="60"></td>
			<td class="td_l_l">编号</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_cId</td>
			<td class="td_l_l"> <input name="L_Linkmans_cId" type="text" id="L_Linkmans_cId" class="int" value="<%=L_Linkmans_cId%>" size="60"></td>
			<td class="td_l_l">公司名称</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lName</td>
			<td class="td_l_l"> <input name="L_Linkmans_lName" type="text" id="L_Linkmans_lName" class="int" value="<%=L_Linkmans_lName%>" size="60"></td>
			<td class="td_l_l">联系人</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lSex</td>
			<td class="td_l_l"> <input name="L_Linkmans_lSex" type="text" id="L_Linkmans_lSex" class="int" value="<%=L_Linkmans_lSex%>" size="60"></td>
			<td class="td_l_l">性别</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lZhiwei</td>
			<td class="td_l_l"> <input name="L_Linkmans_lZhiwei" type="text" id="L_Linkmans_lZhiwei" class="int" value="<%=L_Linkmans_lZhiwei%>" size="60"></td>
			<td class="td_l_l">职位</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lQQ</td>
			<td class="td_l_l"> <input name="L_Linkmans_lQQ" type="text" id="L_Linkmans_lQQ" class="int" value="<%=L_Linkmans_lQQ%>" size="60"></td>
			<td class="td_l_l">腾讯QQ</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lMSN</td>
			<td class="td_l_l"> <input name="L_Linkmans_lMSN" type="text" id="L_Linkmans_lMSN" class="int" value="<%=L_Linkmans_lMSN%>" size="60"></td>
			<td class="td_l_l">MSN</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lMobile</td>
			<td class="td_l_l"> <input name="L_Linkmans_lMobile" type="text" id="L_Linkmans_lMobile" class="int" value="<%=L_Linkmans_lMobile%>" size="60"></td>
			<td class="td_l_l">手机号码</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lALWW</td>
			<td class="td_l_l"> <input name="L_Linkmans_lALWW" type="text" id="L_Linkmans_lALWW" class="int" value="<%=L_Linkmans_lALWW%>" size="60"></td>
			<td class="td_l_l">阿里旺旺</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lTel</td>
			<td class="td_l_l"> <input name="L_Linkmans_lTel" type="text" id="L_Linkmans_lTel" class="int" value="<%=L_Linkmans_lTel%>" size="60"></td>
			<td class="td_l_l">电话</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lEmail</td>
			<td class="td_l_l"> <input name="L_Linkmans_lEmail" type="text" id="L_Linkmans_lEmail" class="int" value="<%=L_Linkmans_lEmail%>" size="60"></td>
			<td class="td_l_l">电子邮件</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lBirthday</td>
			<td class="td_l_l"> <input name="L_Linkmans_lBirthday" type="text" id="L_Linkmans_lBirthday" class="int" value="<%=L_Linkmans_lBirthday%>" size="60"></td>
			<td class="td_l_l">生日</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lContent</td>
			<td class="td_l_l"> <input name="L_Linkmans_lContent" type="text" id="L_Linkmans_lContent" class="int" value="<%=L_Linkmans_lContent%>" size="60"></td>
			<td class="td_l_l">详情备注</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lUser</td>
			<td class="td_l_l"> <input name="L_Linkmans_lUser" type="text" id="L_Linkmans_lUser" class="int" value="<%=L_Linkmans_lUser%>" size="60"></td>
			<td class="td_l_l">业务员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Linkmans_lTime</td>
			<td class="td_l_l"> <input name="L_Linkmans_lTime" type="text" id="L_Linkmans_lTime" class="int" value="<%=L_Linkmans_lTime%>" size="60"></td>
			<td class="td_l_l">录入时间</td>
        </tr>
      </table>
    </td> 
  </tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdb10 "> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
			<tr> 
				<td class="td_r_l" style="border-top:0;" COLSPAN="6"><input name="Submit" type="submit" class="button45" id="Submit" value=" <%=L_Submit%> "></td>
			</tr>
		</table>
	</td>
  </tr>
</table>
</form>
<%  '保存联系人字段
	If action="SaveLinkmans" then
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "'联系人表 Linkmans" & VbCrLf
		TempStr = TempStr & "L_Linkmans_lId="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_cId="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_cId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lName="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lName")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lSex="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lSex")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lZhiwei="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lZhiwei")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lQQ="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lQQ")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lMSN="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lMSN")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lMobile="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lMobile")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lALWW="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lALWW")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lTel="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lTel")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lEmail="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lEmail")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lBirthday="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lBirthday")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lContent="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lContent")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lUser="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Linkmans_lTime="& Chr(34) & replace(Trim(Request.Form("L_Linkmans_lTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "%" & chr(62) & VbCrLf
			ADODB_SaveToFile TempStr,"../Lang/zh-cn/Linkmans.asp"
		If GBL_CHK_TempStr = "" Then
			if ""&YNalert&"" = 1 then
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&alert2&"';</script>")
			else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"';</script>")
			end if
		Else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&GBL_CHK_TempStr&"';</script>")
		End If
	End if
%>
<%
elseIf otype="Records" then '读取跟单字段
%>
<form action="?otype=Records&action=SaveRecords" method="post">
<table width="100%" border="0" cellpadding="0" cellspacing="0" id="TableMain">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdt10"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Records_rId</td>
			<td class="td_l_l"> <input name="L_Records_rId" type="text" id="L_Records_rId" class="int" value="<%=L_Records_rId%>" size="60"></td>
			<td class="td_l_l">编号</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Records_cId</td>
			<td class="td_l_l"> <input name="L_Records_cId" type="text" id="L_Records_cId" class="int" value="<%=L_Records_cId%>" size="60"></td>
			<td class="td_l_l">公司名称</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Records_rType</td>
			<td class="td_l_l"> <input name="L_Records_rType" type="text" id="L_Records_rType" class="int" value="<%=L_Records_rType%>" size="60"></td>
			<td class="td_l_l">跟单类型</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Records_rState</td>
			<td class="td_l_l"> <input name="L_Records_rState" type="text" id="L_Records_rState" class="int" value="<%=L_Records_rState%>" size="60"></td>
			<td class="td_l_l">跟单进度</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Records_rLinkman</td>
			<td class="td_l_l"> <input name="L_Records_rLinkman" type="text" id="L_Records_rLinkman" class="int" value="<%=L_Records_rLinkman%>" size="60"></td>
			<td class="td_l_l">跟单对象</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Records_rNextTime</td>
			<td class="td_l_l"> <input name="L_Records_rNextTime" type="text" id="L_Records_rNextTime" class="int" value="<%=L_Records_rNextTime%>" size="60"></td>
			<td class="td_l_l">下次联系</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Records_rRemind</td>
			<td class="td_l_l"> <input name="L_Records_rRemind" type="text" id="L_Records_rRemind" class="int" value="<%=L_Records_rRemind%>" size="60"></td>
			<td class="td_l_l">提醒时间</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Records_rContent</td>
			<td class="td_l_l"> <input name="L_Records_rContent" type="text" id="L_Records_rContent" class="int" value="<%=L_Records_rContent%>" size="60"></td>
			<td class="td_l_l">详情备注</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Records_rUser</td>
			<td class="td_l_l"> <input name="L_Records_rUser" type="text" id="L_Records_rUser" class="int" value="<%=L_Records_rUser%>" size="60"></td>
			<td class="td_l_l">业务员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Records_rTime</td>
			<td class="td_l_l"> <input name="L_Records_rTime" type="text" id="L_Records_rTime" class="int" value="<%=L_Records_rTime%>" size="60"></td>
			<td class="td_l_l">录入时间</td>
        </tr>
      </table>
    </td> 
  </tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdb10 "> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
			<tr> 
				<td class="td_r_l" style="border-top:0;" COLSPAN="6"><input name="Submit" type="submit" class="button45" id="Submit" value=" <%=L_Submit%> "></td>
			</tr>
		</table>
	</td>
  </tr>
</table>
</form>
<% '保存跟单字段
	If action="SaveRecords" then
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "'跟单记录表 Records" & VbCrLf
		TempStr = TempStr & "L_Records_rId="& Chr(34) & replace(Trim(Request.Form("L_Records_rId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Records_cId="& Chr(34) & replace(Trim(Request.Form("L_Records_cId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Records_rType="& Chr(34) & replace(Trim(Request.Form("L_Records_rType")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Records_rState="& Chr(34) & replace(Trim(Request.Form("L_Records_rState")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Records_rLinkman="& Chr(34) & replace(Trim(Request.Form("L_Records_rLinkman")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Records_rNextTime="& Chr(34) & replace(Trim(Request.Form("L_Records_rNextTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Records_rRemind="& Chr(34) & replace(Trim(Request.Form("L_Records_rRemind")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Records_rContent="& Chr(34) & replace(Trim(Request.Form("L_Records_rContent")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Records_rUser="& Chr(34) & replace(Trim(Request.Form("L_Records_rUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Records_rTime="& Chr(34) & replace(Trim(Request.Form("L_Records_rTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "%" & chr(62) & VbCrLf
			ADODB_SaveToFile TempStr,"../Lang/zh-cn/Records.asp"
		If GBL_CHK_TempStr = "" Then
			if ""&YNalert&"" = 1 then
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&alert2&"';</script>")
			else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"';</script>")
			end if
		Else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&GBL_CHK_TempStr&"';</script>")
		End If
	End if
%>
<%
elseIf otype="Order" then '读取订单字段
%>
<form action="?otype=Order&action=SaveOrder" method="post">
<table width="100%" border="0" cellpadding="0" cellspacing="0" id="TableMain">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdt10"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oId</td>
			<td class="td_l_l"> <input name="L_Order_oId" type="text" id="L_Order_oId" class="int" value="<%=L_Order_oId%>" size="60"></td>
			<td class="td_l_l">编号</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_cId</td>
			<td class="td_l_l"> <input name="L_Order_cId" type="text" id="L_Order_cId" class="int" value="<%=L_Order_cId%>" size="60"></td>
			<td class="td_l_l">公司名称</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oLinkman</td>
			<td class="td_l_l"> <input name="L_Order_oLinkman" type="text" id="L_Order_oLinkman" class="int" value="<%=L_Order_oLinkman%>" size="60"></td>
			<td class="td_l_l">联系人</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oCode</td>
			<td class="td_l_l"> <input name="L_Order_oCode" type="text" id="L_Order_oCode" class="int" value="<%=L_Order_oCode%>" size="60"></td>
			<td class="td_l_l">订单编号</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oSDate</td>
			<td class="td_l_l"> <input name="L_Order_oSDate" type="text" id="L_Order_oSDate" class="int" value="<%=L_Order_oSDate%>" size="60"></td>
			<td class="td_l_l">下单日期</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oEDate</td>
			<td class="td_l_l"> <input name="L_Order_oEDate" type="text" id="L_Order_oEDate" class="int" value="<%=L_Order_oEDate%>" size="60"></td>
			<td class="td_l_l">交单日期</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oDeposit</td>
			<td class="td_l_l"> <input name="L_Order_oDeposit" type="text" id="L_Order_oDeposit" class="int" value="<%=L_Order_oDeposit%>" size="60"></td>
			<td class="td_l_l">预付款</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oMoney</td>
			<td class="td_l_l"> <input name="L_Order_oMoney" type="text" id="L_Order_oMoney" class="int" value="<%=L_Order_oMoney%>" size="60"></td>
			<td class="td_l_l">订单金额</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oState</td>
			<td class="td_l_l"> <input name="L_Order_oState" type="text" id="L_Order_oState" class="int" value="<%=L_Order_oState%>" size="60"></td>
			<td class="td_l_l">订单状态</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oContent</td>
			<td class="td_l_l"> <input name="L_Order_oContent" type="text" id="L_Order_oContent" class="int" value="<%=L_Order_oContent%>" size="60"></td>
			<td class="td_l_l">详情备注</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oUser</td>
			<td class="td_l_l"> <input name="L_Order_oUser" type="text" id="L_Order_oUser" class="int" value="<%=L_Order_oUser%>" size="60"></td>
			<td class="td_l_l">业务员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_oTime</td>
			<td class="td_l_l"> <input name="L_Order_oTime" type="text" id="L_Order_oTime" class="int" value="<%=L_Order_oTime%>" size="60"></td>
			<td class="td_l_l">录入时间</td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin-top:10px;">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_osId</td>
			<td class="td_l_l"> <input name="L_Order_Products_osId" type="text" id="L_Order_Products_osId" class="int" value="<%=L_Order_Products_osId%>" size="60"></td>
			<td class="td_l_l">编号</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oId</td>
			<td class="td_l_l"> <input name="L_Order_Products_oId" type="text" id="L_Order_Products_oId" class="int" value="<%=L_Order_Products_oId%>" size="60"></td>
			<td class="td_l_l">订单编号</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_cId</td>
			<td class="td_l_l"> <input name="L_Order_Products_cId" type="text" id="L_Order_Products_cId" class="int" value="<%=L_Order_Products_cId%>" size="60"></td>
			<td class="td_l_l">公司名称</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_ProId</td>
			<td class="td_l_l"> <input name="L_Order_Products_ProId" type="text" id="L_Order_Products_ProId" class="int" value="<%=L_Order_Products_ProId%>" size="60"></td>
			<td class="td_l_l">产品编号</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oProTitle</td>
			<td class="td_l_l"> <input name="L_Order_Products_oProTitle" type="text" id="L_Order_Products_oProTitle" class="int" value="<%=L_Order_Products_oProTitle%>" size="60"></td>
			<td class="td_l_l">产品名称</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oProItemA</td>
			<td class="td_l_l"> <input name="L_Order_Products_oProItemA" type="text" id="L_Order_Products_oProItemA" class="int" value="<%=L_Order_Products_oProItemA%>" size="60"></td>
			<td class="td_l_l">自定义属性A</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oProItemB</td>
			<td class="td_l_l"> <input name="L_Order_Products_oProItemB" type="text" id="L_Order_Products_oProItemB" class="int" value="<%=L_Order_Products_oProItemB%>" size="60"></td>
			<td class="td_l_l">自定义属性B</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oProItemC</td>
			<td class="td_l_l"> <input name="L_Order_Products_oProItemC" type="text" id="L_Order_Products_oProItemC" class="int" value="<%=L_Order_Products_oProItemC%>" size="60"></td>
			<td class="td_l_l">自定义属性C</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oProItemD</td>
			<td class="td_l_l"> <input name="L_Order_Products_oProItemD" type="text" id="L_Order_Products_oProItemD" class="int" value="<%=L_Order_Products_oProItemD%>" size="60"></td>
			<td class="td_l_l">自定义属性D</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oProItemE</td>
			<td class="td_l_l"> <input name="L_Order_Products_oProItemE" type="text" id="L_Order_Products_oProItemE" class="int" value="<%=L_Order_Products_oProItemE%>" size="60"></td>
			<td class="td_l_l">自定义属性E</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oProPrice</td>
			<td class="td_l_l"> <input name="L_Order_Products_oProPrice" type="text" id="L_Order_Products_oProPrice" class="int" value="<%=L_Order_Products_oProPrice%>" size="60"></td>
			<td class="td_l_l">产品单价</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oProNum</td>
			<td class="td_l_l"> <input name="L_Order_Products_oProNum" type="text" id="L_Order_Products_oProNum" class="int" value="<%=L_Order_Products_oProNum%>" size="60"></td>
			<td class="td_l_l">数量</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oDiscount</td>
			<td class="td_l_l"> <input name="L_Order_Products_oDiscount" type="text" id="L_Order_Products_oDiscount" class="int" value="<%=L_Order_Products_oDiscount%>" size="60"></td>
			<td class="td_l_l">折扣金额</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oMoney</td>
			<td class="td_l_l"> <input name="L_Order_Products_oMoney" type="text" id="L_Order_Products_oMoney" class="int" value="<%=L_Order_Products_oMoney%>" size="60"></td>
			<td class="td_l_l">小计金额</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oContent</td>
			<td class="td_l_l"> <input name="L_Order_Products_oContent" type="text" id="L_Order_Products_oContent" class="int" value="<%=L_Order_Products_oContent%>" size="60"></td>
			<td class="td_l_l">详情备注</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oUser</td>
			<td class="td_l_l"> <input name="L_Order_Products_oUser" type="text" id="L_Order_Products_oUser" class="int" value="<%=L_Order_Products_oUser%>" size="60"></td>
			<td class="td_l_l">业务员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Order_Products_oTime</td>
			<td class="td_l_l"> <input name="L_Order_Products_oTime" type="text" id="L_Order_Products_oTime" class="int" value="<%=L_Order_Products_oTime%>" size="60"></td>
			<td class="td_l_l">录入时间</td>
        </tr>
      </table>
    </td> 
  </tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdb10 "> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
			<tr> 
				<td class="td_r_l" style="border-top:0;" COLSPAN="6">
				<input name="Submit" type="submit" class="button45" id="Submit" value=" <%=L_Submit%> ">
				</td>
			</tr>
		</table>
	</td>
  </tr>
</table>
</form>
<%
	If action="SaveOrder" then
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "'订单报价记录 Order" & VbCrLf
		TempStr = TempStr & "L_Order_oId="& Chr(34) & replace(Trim(Request.Form("L_Order_oId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_cId="& Chr(34) & replace(Trim(Request.Form("L_Order_cId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_oCode="& Chr(34) & replace(Trim(Request.Form("L_Order_oCode")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_oLinkman="& Chr(34) & replace(Trim(Request.Form("L_Order_oLinkman")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_oSDate="& Chr(34) & replace(Trim(Request.Form("L_Order_oSDate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_oEDate="& Chr(34) & replace(Trim(Request.Form("L_Order_oEDate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_oDeposit="& Chr(34) & replace(Trim(Request.Form("L_Order_oDeposit")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_oMoney="& Chr(34) & replace(Trim(Request.Form("L_Order_oMoney")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_oState="& Chr(34) & replace(Trim(Request.Form("L_Order_oState")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_oContent="& Chr(34) & replace(Trim(Request.Form("L_Order_oContent")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_oUser="& Chr(34) & replace(Trim(Request.Form("L_Order_oUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_oTime="& Chr(34) & replace(Trim(Request.Form("L_Order_oTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf & VbCrLf
		TempStr = TempStr & "L_Order_Products_osId="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_osId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oId="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_cId="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_cId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_ProId="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_ProId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oProTitle="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oProTitle")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oProItemA="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oProItemA")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oProItemB="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oProItemB")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oProItemC="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oProItemC")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oProItemD="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oProItemD")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oProItemE="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oProItemE")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oProPrice="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oProPrice")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oProNum="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oProNum")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oDiscount="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oDiscount")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oMoney="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oMoney")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oContent="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oContent")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oUser="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Order_Products_oTime="& Chr(34) & replace(Trim(Request.Form("L_Order_Products_oTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "%" & chr(62) & VbCrLf
			ADODB_SaveToFile TempStr,"../Lang/zh-cn/Order.asp"
		If GBL_CHK_TempStr = "" Then
			if ""&YNalert&"" = 1 then
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&alert2&"';</script>")
			else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"';</script>")
			end if
		Else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&GBL_CHK_TempStr&"';</script>")
		End If
	End if
%>
<%
elseIf otype="Hetong" then '读取合同字段
%>
<form action="?otype=Hetong&action=SaveHetong" method="post">
<table width="100%" border="0" cellpadding="0" cellspacing="0" id="TableMain">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdt10"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hId</td>
			<td class="td_l_l"> <input name="L_Hetong_hId" type="text" id="L_Hetong_hId" class="int" value="<%=L_Hetong_hId%>" size="60"></td>
			<td class="td_l_l">编号</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Hetong_cId</td>
			<td class="td_l_l"> <input name="L_Hetong_cId" type="text" id="L_Hetong_cId" class="int" value="<%=L_Hetong_cId%>" size="60"></td>
			<td class="td_l_l">公司名称</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Hetong_oId</td>
			<td class="td_l_l"> <input name="L_Hetong_oId" type="text" id="L_Hetong_oId" class="int" value="<%=L_Hetong_oId%>" size="60"></td>
			<td class="td_l_l">订单编号</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Hetong_hNum</td>
			<td class="td_l_l"> <input name="L_Hetong_hNum" type="text" id="L_Hetong_hNum" class="int" value="<%=L_Hetong_hNum%>" size="60"></td>
			<td class="td_l_l">合同编号（自定义）</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hSdate</td>
			<td class="td_l_l"> <input name="L_Hetong_hSdate" type="text" id="L_Hetong_hSdate" class="int" value="<%=L_Hetong_hSdate%>" size="60"></td>
			<td class="td_l_l">起始时间</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hEdate</td>
			<td class="td_l_l"> <input name="L_Hetong_hEdate" type="text" id="L_Hetong_hEdate" class="int" value="<%=L_Hetong_hEdate%>" size="60"></td>
			<td class="td_l_l">截至时间</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hType</td>
			<td class="td_l_l"> <input name="L_Hetong_hType" type="text" id="L_Hetong_hType" class="int" value="<%=L_Hetong_hType%>" size="60"></td>
			<td class="td_l_l">合同分类</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hMoney</td>
			<td class="td_l_l"> <input name="L_Hetong_hMoney" type="text" id="L_Hetong_hMoney" class="int" value="<%=L_Hetong_hMoney%>" size="60"></td>
			<td class="td_l_l">总金额</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hRevenue</td>
			<td class="td_l_l"> <input name="L_Hetong_hRevenue" type="text" id="L_Hetong_hRevenue" class="int" value="<%=L_Hetong_hRevenue%>" size="60"></td>
			<td class="td_l_l">已收款</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hOwed</td>
			<td class="td_l_l"> <input name="L_Hetong_hOwed" type="text" id="L_Hetong_hOwed" class="int" value="<%=L_Hetong_hOwed%>" size="60"></td>
			<td class="td_l_l">欠款</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hInvoice</td>
			<td class="td_l_l"> <input name="L_Hetong_hInvoice" type="text" id="L_Hetong_hInvoice" class="int" value="<%=L_Hetong_hInvoice%>" size="60"></td>
			<td class="td_l_l">含发票</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hTax</td>
			<td class="td_l_l"> <input name="L_Hetong_hTax" type="text" id="L_Hetong_hTax" class="int" value="<%=L_Hetong_hTax%>" size="60"></td>
			<td class="td_l_l">含税</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hState</td>
			<td class="td_l_l"> <input name="L_Hetong_hState" type="text" id="L_Hetong_hState" class="int" value="<%=L_Hetong_hState%>" size="60"></td>
			<td class="td_l_l">合同状态</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hState_0</td>
			<td class="td_l_l"> <input name="L_Hetong_hState_0" type="text" id="L_Hetong_hState_0" class="int" value="<%=L_Hetong_hState_0%>" size="60"></td>
			<td class="td_l_l">新增</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hState_1</td>
			<td class="td_l_l"> <input name="L_Hetong_hState_1" type="text" id="L_Hetong_hState_1" class="int" value="<%=L_Hetong_hState_1%>" size="60"></td>
			<td class="td_l_l">审核中</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hState_2</td>
			<td class="td_l_l"> <input name="L_Hetong_hState_2" type="text" id="L_Hetong_hState_2" class="int" value="<%=L_Hetong_hState_2%>" size="60"></td>
			<td class="td_l_l">合同有效</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hState_3</td>
			<td class="td_l_l"> <input name="L_Hetong_hState_3" type="text" id="L_Hetong_hState_3" class="int" value="<%=L_Hetong_hState_3%>" size="60"></td>
			<td class="td_l_l">合同无效</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hState_4</td>
			<td class="td_l_l"> <input name="L_Hetong_hState_4" type="text" id="L_Hetong_hState_4" class="int" value="<%=L_Hetong_hState_4%>" size="60"></td>
			<td class="td_l_l">通过</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hState_5</td>
			<td class="td_l_l"> <input name="L_Hetong_hState_5" type="text" id="L_Hetong_hState_5" class="int" value="<%=L_Hetong_hState_5%>" size="60"></td>
			<td class="td_l_l">驳回</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hState_6</td>
			<td class="td_l_l"> <input name="L_Hetong_hState_6" type="text" id="L_Hetong_hState_6" class="int" value="<%=L_Hetong_hState_6%>" size="60"></td>
			<td class="td_l_l">转待审</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hContent</td>
			<td class="td_l_l"> <input name="L_Hetong_hContent" type="text" id="L_Hetong_hContent" class="int" value="<%=L_Hetong_hContent%>" size="60"></td>
			<td class="td_l_l">详情备注</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hAudit</td>
			<td class="td_l_l"> <input name="L_Hetong_hAudit" type="text" id="L_Hetong_hAudit" class="int" value="<%=L_Hetong_hAudit%>" size="60"></td>
			<td class="td_l_l">审核人员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hAuditTime</td>
			<td class="td_l_l"> <input name="L_Hetong_hAuditTime" type="text" id="L_Hetong_hAuditTime" class="int" value="<%=L_Hetong_hAuditTime%>" size="60"></td>
			<td class="td_l_l">审核时间</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hAuditReasons</td>
			<td class="td_l_l"> <input name="L_Hetong_hAuditReasons" type="text" id="L_Hetong_hAuditReasons" class="int" value="<%=L_Hetong_hAuditReasons%>" size="60"></td>
			<td class="td_l_l">审核原因</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hUser</td>
			<td class="td_l_l"> <input name="L_Hetong_hUser" type="text" id="L_Hetong_hUser" class="int" value="<%=L_Hetong_hUser%>" size="60"></td>
			<td class="td_l_l">业务员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_hTime</td>
			<td class="td_l_l"> <input name="L_Hetong_hTime" type="text" id="L_Hetong_hTime" class="int" value="<%=L_Hetong_hTime%>" size="60"></td>
			<td class="td_l_l">录入时间</td>
        </tr>
      </table>
	  
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1" style="margin-top:10px;">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rId</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rId" type="text" id="L_Hetong_Renew_rId" class="int" value="<%=L_Hetong_Renew_rId%>" size="60"></td>
			<td class="td_l_l">编号</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Hetong_Renew_hID</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_hID" type="text" id="L_Hetong_Renew_hID" class="int" value="<%=L_Hetong_Renew_hID%>" size="60"></td>
			<td class="td_l_l">合同编号（自动）</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Hetong_Renew_rEdate</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rEdate" type="text" id="L_Hetong_Renew_rEdate" class="int" value="<%=L_Hetong_Renew_rEdate%>" size="60"></td>
			<td class="td_l_l">到期时间</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rMoney</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rMoney" type="text" id="L_Hetong_Renew_rMoney" class="int" value="<%=L_Hetong_Renew_rMoney%>" size="60"></td>
			<td class="td_l_l">总金额</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rRevenue</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rRevenue" type="text" id="L_Hetong_Renew_rRevenue" class="int" value="<%=L_Hetong_Renew_rRevenue%>" size="60"></td>
			<td class="td_l_l">已收款</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rState</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rState" type="text" id="L_Hetong_Renew_rState" class="int" value="<%=L_Hetong_Renew_rState%>" size="60"></td>
			<td class="td_l_l">续费状态</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rState_0</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rState_0" type="text" id="L_Hetong_Renew_rState_0" class="int" value="<%=L_Hetong_Renew_rState_0%>" size="60"></td>
			<td class="td_l_l">新增</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rState_1</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rState_1" type="text" id="L_Hetong_Renew_rState_1" class="int" value="<%=L_Hetong_Renew_rState_1%>" size="60"></td>
			<td class="td_l_l">审核中</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rState_2</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rState_2" type="text" id="L_Hetong_Renew_rState_2" class="int" value="<%=L_Hetong_Renew_rState_2%>" size="60"></td>
			<td class="td_l_l">续费成功</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rState_3</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rState_3" type="text" id="L_Hetong_Renew_rState_3" class="int" value="<%=L_Hetong_Renew_rState_3%>" size="60"></td>
			<td class="td_l_l">续费无效</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rState_4</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rState_4" type="text" id="L_Hetong_Renew_rState_4" class="int" value="<%=L_Hetong_Renew_rState_4%>" size="60"></td>
			<td class="td_l_l">通过</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rState_5</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rState_5" type="text" id="L_Hetong_Renew_rState_5" class="int" value="<%=L_Hetong_Renew_rState_5%>" size="60"></td>
			<td class="td_l_l">驳回</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rState_6</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rState_6" type="text" id="L_Hetong_Renew_rState_6" class="int" value="<%=L_Hetong_Renew_rState_6%>" size="60"></td>
			<td class="td_l_l">转待审</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rContent</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rContent" type="text" id="L_Hetong_Renew_rContent" class="int" value="<%=L_Hetong_Renew_rContent%>" size="60"></td>
			<td class="td_l_l">详情备注</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rAudit</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rAudit" type="text" id="L_Hetong_Renew_rAudit" class="int" value="<%=L_Hetong_Renew_rAudit%>" size="60"></td>
			<td class="td_l_l">审核人员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rAuditTime</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rAuditTime" type="text" id="L_Hetong_Renew_rAuditTime" class="int" value="<%=L_Hetong_Renew_rAuditTime%>" size="60"></td>
			<td class="td_l_l">审核时间</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rAuditReasons</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rAuditReasons" type="text" id="L_Hetong_Renew_rAuditReasons" class="int" value="<%=L_Hetong_Renew_rAuditReasons%>" size="60"></td>
			<td class="td_l_l">审核原因</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rUser</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rUser" type="text" id="L_Hetong_Renew_rUser" class="int" value="<%=L_Hetong_Renew_rUser%>" size="60"></td>
			<td class="td_l_l">业务员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Hetong_Renew_rTime</td>
			<td class="td_l_l"> <input name="L_Hetong_Renew_rTime" type="text" id="L_Hetong_Renew_rTime" class="int" value="<%=L_Hetong_Renew_rTime%>" size="60"></td>
			<td class="td_l_l">录入时间</td>
        </tr>
      </table>
    </td> 
  </tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdb10 "> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
			<tr> 
				<td class="td_r_l" style="border-top:0;" COLSPAN="6"><input name="Submit" type="submit" class="button45" id="Submit" value=" <%=L_Submit%> "></td>
			</tr>
		</table>
	</td>
  </tr>
</table>
</form>
<%
	If action="SaveHetong" then
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "'合同记录 Hetong" & VbCrLf
		TempStr = TempStr & "L_Hetong_hId="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_cId="& Chr(34) & replace(Trim(Request.Form("L_Hetong_cId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_oId="& Chr(34) & replace(Trim(Request.Form("L_Hetong_oId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hNum="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hNum")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hSdate="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hSdate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hEdate="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hEdate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hType="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hType")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hMoney="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hMoney")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hRevenue="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hRevenue")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hOwed="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hOwed")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hInvoice="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hInvoice")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hTax="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hTax")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hState="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hState")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hState_0="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hState_0")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hState_1="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hState_1")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hState_2="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hState_2")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hState_3="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hState_3")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hState_4="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hState_4")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hState_5="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hState_5")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hState_6="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hState_6")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hContent="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hContent")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hAudit="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hAudit")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hAuditTime="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hAuditTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hAuditReasons="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hAuditReasons")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hUser="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_hTime="& Chr(34) & replace(Trim(Request.Form("L_Hetong_hTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		
		TempStr = TempStr & "L_Hetong_Renew_rId="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_hID="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_hID")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rEdate="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rEdate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rMoney="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rMoney")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rRevenue="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rRevenue")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rState="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rState")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rState_0="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rState_0")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rState_1="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rState_1")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rState_2="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rState_2")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rState_3="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rState_3")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rState_4="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rState_4")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rState_5="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rState_5")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rState_6="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rState_6")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rContent="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rContent")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rAudit="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rAudit")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rAuditTime="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rAuditTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rAuditReasons="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rAuditReasons")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rUser="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Hetong_Renew_rTime="& Chr(34) & replace(Trim(Request.Form("L_Hetong_Renew_rTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		
		TempStr = TempStr & "%" & chr(62) & VbCrLf
			ADODB_SaveToFile TempStr,"../Lang/zh-cn/Hetong.asp"
		If GBL_CHK_TempStr = "" Then
			if ""&YNalert&"" = 1 then
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&alert2&"';</script>")
			else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"';</script>")
			end if
		Else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&GBL_CHK_TempStr&"';</script>")
		End If
	End if
%>
<%
elseIf otype="Service" then
%>

<form action="?otype=Service&action=SaveService" method="post">
<table width="100%" border="0" cellpadding="0" cellspacing="0" id="TableMain">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdt10"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sId</td>
			<td class="td_l_l"> <input name="L_Service_sId" type="text" id="L_Service_sId" class="int" value="<%=L_Service_sId%>" size="60"></td>
			<td class="td_l_l">编号</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Service_cId</td>
			<td class="td_l_l"> <input name="L_Service_cId" type="text" id="L_Service_cId" class="int" value="<%=L_Service_cId%>" size="60"></td>
			<td class="td_l_l">公司名称</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_ProId</td>
			<td class="td_l_l"> <input name="L_Service_ProId" type="text" id="L_Service_ProId" class="int" value="<%=L_Service_ProId%>" size="60"></td>
			<td class="td_l_l">相关产品</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sLinkman</td>
			<td class="td_l_l"> <input name="L_Service_sLinkman" type="text" id="L_Service_sLinkman" class="int" value="<%=L_Service_sLinkman%>" size="60"></td>
			<td class="td_l_l">联系人</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Service_sTitle</td>
			<td class="td_l_l"> <input name="L_Service_sTitle" type="text" id="L_Service_sTitle" class="int" value="<%=L_Service_sTitle%>" size="60"></td>
			<td class="td_l_l">反馈主题</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sType</td>
			<td class="td_l_l"> <input name="L_Service_sType" type="text" id="L_Service_sType" class="int" value="<%=L_Service_sType%>" size="60"></td>
			<td class="td_l_l">反馈分类</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sContent</td>
			<td class="td_l_l"> <input name="L_Service_sContent" type="text" id="L_Service_sContent" class="int" value="<%=L_Service_sContent%>" size="60"></td>
			<td class="td_l_l">详情备注</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sSolve</td>
			<td class="td_l_l"> <input name="L_Service_sSolve" type="text" id="L_Service_sSolve" class="int" value="<%=L_Service_sSolve%>" size="60"></td>
			<td class="td_l_l">是否解决</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sSolve_0</td>
			<td class="td_l_l"> <input name="L_Service_sSolve_0" type="text" id="L_Service_sSolve_0" class="int" value="<%=L_Service_sSolve_0%>" size="60"></td>
			<td class="td_l_l">未解决（转客服处理）</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sSolve_1</td>
			<td class="td_l_l"> <input name="L_Service_sSolve_1" type="text" id="L_Service_sSolve_1" class="int" value="<%=L_Service_sSolve_1%>" size="60"></td>
			<td class="td_l_l">已解决</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sInfo</td>
			<td class="td_l_l"> <input name="L_Service_sInfo" type="text" id="L_Service_sInfo" class="int" value="<%=L_Service_sInfo%>" size="60"></td>
			<td class="td_l_l">处理结果</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sUser</td>
			<td class="td_l_l"> <input name="L_Service_sUser" type="text" id="L_Service_sUser" class="int" value="<%=L_Service_sUser%>" size="60"></td>
			<td class="td_l_l">业务员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sSDate</td>
			<td class="td_l_l"> <input name="L_Service_sSDate" type="text" id="L_Service_sSDate" class="int" value="<%=L_Service_sSDate%>" size="60"></td>
			<td class="td_l_l">反馈日期</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sEDate</td>
			<td class="td_l_l"> <input name="L_Service_sEDate" type="text" id="L_Service_sEDate" class="int" value="<%=L_Service_sEDate%>" size="60"></td>
			<td class="td_l_l">结束日期</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Service_sTime</td>
			<td class="td_l_l"> <input name="L_Service_sTime" type="text" id="L_Service_sTime" class="int" value="<%=L_Service_sTime%>" size="60"></td>
			<td class="td_l_l">录入时间</td>
        </tr>
      </table>
    </td> 
  </tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdb10 "> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
			<tr> 
				<td class="td_r_l" style="border-top:0;" COLSPAN="6"><input name="Submit" type="submit" class="button45" id="Submit" value=" <%=L_Submit%> "></td>
			</tr>
		</table>
	</td>
  </tr>
</table>
</form>
<%
	If action="SaveService" then
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "'服务记录 Service" & VbCrLf
		TempStr = TempStr & "L_Service_sId="& Chr(34) & replace(Trim(Request.Form("L_Service_sId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_cId="& Chr(34) & replace(Trim(Request.Form("L_Service_cId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_ProId="& Chr(34) & replace(Trim(Request.Form("L_Service_ProId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sLinkman="& Chr(34) & replace(Trim(Request.Form("L_Service_sLinkman")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sTitle="& Chr(34) & replace(Trim(Request.Form("L_Service_sTitle")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sType="& Chr(34) & replace(Trim(Request.Form("L_Service_sType")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sContent="& Chr(34) & replace(Trim(Request.Form("L_Service_sContent")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sSolve="& Chr(34) & replace(Trim(Request.Form("L_Service_sSolve")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sSolve_0="& Chr(34) & replace(Trim(Request.Form("L_Service_sSolve_0")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sSolve_1="& Chr(34) & replace(Trim(Request.Form("L_Service_sSolve_1")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sInfo="& Chr(34) & replace(Trim(Request.Form("L_Service_sInfo")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sUser="& Chr(34) & replace(Trim(Request.Form("L_Service_sUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sSDate="& Chr(34) & replace(Trim(Request.Form("L_Service_sSDate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sEDate="& Chr(34) & replace(Trim(Request.Form("L_Service_sEDate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Service_sTime="& Chr(34) & replace(Trim(Request.Form("L_Service_sTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "%" & chr(62) & VbCrLf
			ADODB_SaveToFile TempStr,"../Lang/zh-cn/Service.asp"
		If GBL_CHK_TempStr = "" Then
			if ""&YNalert&"" = 1 then
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&alert2&"';</script>")
			else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"';</script>")
			end if
		Else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&GBL_CHK_TempStr&"';</script>")
		End If
	End if
%>
<%
elseIf otype="Expense" then
%>

<form action="?otype=Expense&action=SaveExpense" method="post">
<table width="100%" border="0" cellpadding="0" cellspacing="0" id="TableMain">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdt10"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Expense_eId</td>
			<td class="td_l_l"> <input name="L_Expense_eId" type="text" id="L_Expense_eId" class="int" value="<%=L_Expense_eId%>" size="60"></td>
			<td class="td_l_l">编号</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Expense_cId</td>
			<td class="td_l_l"> <input name="L_Expense_cId" type="text" id="L_Expense_cId" class="int" value="<%=L_Expense_cId%>" size="60"></td>
			<td class="td_l_l">公司名称</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Expense_eDate</td>
			<td class="td_l_l"> <input name="L_Expense_eDate" type="text" id="L_Expense_eDate" class="int" value="<%=L_Expense_eDate%>" size="60"></td>
			<td class="td_l_l">收支日期</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Expense_eOutIn</td>
			<td class="td_l_l"> <input name="L_Expense_eOutIn" type="text" id="L_Expense_eOutIn" class="int" value="<%=L_Expense_eOutIn%>" size="60"></td>
			<td class="td_l_l">收支类型</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Expense_eType</td>
			<td class="td_l_l"> <input name="L_Expense_eType" type="text" id="L_Expense_eType" class="int" value="<%=L_Expense_eType%>" size="60"></td>
			<td class="td_l_l">费用类型</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Expense_eMoney</td>
			<td class="td_l_l"> <input name="L_Expense_eMoney" type="text" id="L_Expense_eMoney" class="int" value="<%=L_Expense_eMoney%>" size="60"></td>
			<td class="td_l_l">总金额</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Expense_eContent</td>
			<td class="td_l_l"> <input name="L_Expense_eContent" type="text" id="L_Expense_eContent" class="int" value="<%=L_Expense_eContent%>" size="60"></td>
			<td class="td_l_l">详情备注</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Expense_eUser</td>
			<td class="td_l_l"> <input name="L_Expense_eUser" type="text" id="L_Expense_eUser" class="int" value="<%=L_Expense_eUser%>" size="60"></td>
			<td class="td_l_l">业务员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Expense_eTime</td>
			<td class="td_l_l"> <input name="L_Expense_eTime" type="text" id="L_Expense_eTime" class="int" value="<%=L_Expense_eTime%>" size="60"></td>
			<td class="td_l_l">录入时间</td>
        </tr>
      </table>
    </td> 
  </tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdb10 "> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
			<tr> 
				<td class="td_r_l" style="border-top:0;" COLSPAN="6"><input name="Submit" type="submit" class="button45" id="Submit" value=" <%=L_Submit%> "></td>
			</tr>
		</table>
	</td>
  </tr>
</table>
</form>
<%
	If action="SaveExpense" then
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "'费用记录 Expense" & VbCrLf
		TempStr = TempStr & "L_Expense_eId="& Chr(34) & replace(Trim(Request.Form("L_Expense_eId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Expense_cId="& Chr(34) & replace(Trim(Request.Form("L_Expense_cId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Expense_eDate="& Chr(34) & replace(Trim(Request.Form("L_Expense_eDate")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Expense_eOutIn="& Chr(34) & replace(Trim(Request.Form("L_Expense_eOutIn")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Expense_eType="& Chr(34) & replace(Trim(Request.Form("L_Expense_eType")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Expense_eMoney="& Chr(34) & replace(Trim(Request.Form("L_Expense_eMoney")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Expense_eContent="& Chr(34) & replace(Trim(Request.Form("L_Expense_eContent")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Expense_eUser="& Chr(34) & replace(Trim(Request.Form("L_Expense_eUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Expense_eTime="& Chr(34) & replace(Trim(Request.Form("L_Expense_eTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "%" & chr(62) & VbCrLf
			ADODB_SaveToFile TempStr,"../Lang/zh-cn/Expense.asp"
		If GBL_CHK_TempStr = "" Then
			if ""&YNalert&"" = 1 then
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&alert2&"';</script>")
			else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"';</script>")
			end if
		Else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&GBL_CHK_TempStr&"';</script>")
		End If
	End if
elseIf otype="File" then
%>

<form action="?otype=File&action=SaveFile" method="post">
<table width="100%" border="0" cellpadding="0" cellspacing="0" id="TableMain">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdt10"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_File_fId</td>
			<td class="td_l_l"> <input name="L_File_fId" type="text" id="L_File_fId" class="int" value="<%=L_File_fId%>" size="60"></td>
			<td class="td_l_l">编号</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_File_cId</td>
			<td class="td_l_l"> <input name="L_File_cId" type="text" id="L_File_cId" class="int" value="<%=L_File_cId%>" size="60"></td>
			<td class="td_l_l">公司名称</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_File_fTitle</td>
			<td class="td_l_l"> <input name="L_File_fTitle" type="text" id="L_File_fTitle" class="int" value="<%=L_File_fTitle%>" size="60"></td>
			<td class="td_l_l">附件标题</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_File_fFile</td>
			<td class="td_l_l"> <input name="L_File_fFile" type="text" id="L_File_fFile" class="int" value="<%=L_File_fFile%>" size="60"></td>
			<td class="td_l_l">下载地址</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_File_fContent</td>
			<td class="td_l_l"> <input name="L_File_fContent" type="text" id="L_File_fContent" class="int" value="<%=L_File_fContent%>" size="60"></td>
			<td class="td_l_l">详情备注</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_File_fUser</td>
			<td class="td_l_l"> <input name="L_File_fUser" type="text" id="L_File_fUser" class="int" value="<%=L_File_fUser%>" size="60"></td>
			<td class="td_l_l">业务员</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_File_fTime</td>
			<td class="td_l_l"> <input name="L_File_fTime" type="text" id="L_File_fTime" class="int" value="<%=L_File_fTime%>" size="60"></td>
			<td class="td_l_l">上传时间</td>
        </tr>
      </table>
    </td> 
  </tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdb10 "> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
			<tr> 
				<td class="td_r_l" style="border-top:0;" COLSPAN="6"><input name="Submit" type="submit" class="button45" id="Submit" value=" <%=L_Submit%> "></td>
			</tr>
		</table>
	</td>
  </tr>
</table>
</form>
<%
	If action="SaveFile" then
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "'附件记录 File" & VbCrLf
		TempStr = TempStr & "L_File_fId="& Chr(34) & replace(Trim(Request.Form("L_File_fId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_File_cId="& Chr(34) & replace(Trim(Request.Form("L_File_cId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_File_fTitle="& Chr(34) & replace(Trim(Request.Form("L_File_fTitle")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_File_fFile="& Chr(34) & replace(Trim(Request.Form("L_File_fFile")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_File_fContent="& Chr(34) & replace(Trim(Request.Form("L_File_fContent")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_File_fUser="& Chr(34) & replace(Trim(Request.Form("L_File_fUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_File_fTime="& Chr(34) & replace(Trim(Request.Form("L_File_fTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "%" & chr(62) & VbCrLf
			ADODB_SaveToFile TempStr,"../Lang/zh-cn/File.asp"
		If GBL_CHK_TempStr = "" Then
			if ""&YNalert&"" = 1 then
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&alert2&"';</script>")
			else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"';</script>")
			end if
		Else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&GBL_CHK_TempStr&"';</script>")
		End If
	End if
elseIf otype="ManageLog" then
%>

<form action="?otype=ManageLog&action=SaveManageLog" method="post">
<table width="100%" border="0" cellpadding="0" cellspacing="0" id="TableMain">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdt10"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
	  <col width="200" /><col width="450" />
        <tr class="tr_t"> 
			<td class="td_l_l"><B>变量名</B></td>
			<td class="td_l_l"><B>显示文字</B></td>
			<td class="td_l_l"><B>原释义</B></td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Logfile_lId</td>
			<td class="td_l_l"> <input name="L_Logfile_lId" type="text" id="L_Logfile_lId" class="int" value="<%=L_Logfile_lId%>" size="60"></td>
			<td class="td_l_l">编号</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Logfile_lcId</td>
			<td class="td_l_l"> <input name="L_Logfile_lcId" type="text" id="L_Logfile_lcId" class="int" value="<%=L_Logfile_lcId%>" size="60"></td>
			<td class="td_l_l">公司名称</td>
        </tr>
        <tr class="tr">
			<td class="td_l_l">L_Logfile_lClass</td>
			<td class="td_l_l"> <input name="L_Logfile_lClass" type="text" id="L_Logfile_lClass" class="int" value="<%=L_Logfile_lClass%>" size="60"></td>
			<td class="td_l_l">数据表</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Logfile_lAction</td>
			<td class="td_l_l"> <input name="L_Logfile_lAction" type="text" id="L_Logfile_lAction" class="int" value="<%=L_Logfile_lAction%>" size="60"></td>
			<td class="td_l_l">行为</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Logfile_lReason</td>
			<td class="td_l_l"> <input name="L_Logfile_lReason" type="text" id="L_Logfile_lReason" class="int" value="<%=L_Logfile_lReason%>" size="60"></td>
			<td class="td_l_l">原因</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Logfile_lUser</td>
			<td class="td_l_l"> <input name="L_Logfile_lUser" type="text" id="L_Logfile_lUser" class="int" value="<%=L_Logfile_lUser%>" size="60"></td>
			<td class="td_l_l">帐号</td>
        </tr>
        <tr class="tr"> 
			<td class="td_l_l">L_Logfile_lTime</td>
			<td class="td_l_l"> <input name="L_Logfile_lTime" type="text" id="L_Logfile_lTime" class="int" value="<%=L_Logfile_lTime%>" size="60"></td>
			<td class="td_l_l">时间</td>
        </tr>
      </table>
    </td> 
  </tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" class="td_n pdl10 pdr10 pdb10 "> 
		<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
			<tr> 
				<td class="td_r_l" style="border-top:0;" COLSPAN="6"><input name="Submit" type="submit" class="button45" id="Submit" value=" <%=L_Submit%> "></td>
			</tr>
		</table>
	</td>
  </tr>
</table>
</form>
<%
	If action="SaveManageLog" then
		TempStr = ""
		TempStr = TempStr & chr(60) & "%" & VbCrLf
		TempStr = TempStr & "'历史记录 Logfile" & VbCrLf
		TempStr = TempStr & "L_Logfile_lId="& Chr(34) & replace(Trim(Request.Form("L_Logfile_lId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Logfile_lcId="& Chr(34) & replace(Trim(Request.Form("L_Logfile_lcId")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Logfile_lClass="& Chr(34) & replace(Trim(Request.Form("L_Logfile_lClass")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Logfile_lAction="& Chr(34) & replace(Trim(Request.Form("L_Logfile_lAction")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Logfile_lReason="& Chr(34) & replace(Trim(Request.Form("L_Logfile_lReason")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Logfile_lUser="& Chr(34) & replace(Trim(Request.Form("L_Logfile_lUser")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "L_Logfile_lTime="& Chr(34) & replace(Trim(Request.Form("L_Logfile_lTime")),CHR(34),"'") & Chr(34) &" " & VbCrLf
		TempStr = TempStr & "%" & chr(62) & VbCrLf
			ADODB_SaveToFile TempStr,"../Lang/zh-cn/Logfile.asp"
		If GBL_CHK_TempStr = "" Then
			if ""&YNalert&"" = 1 then
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&alert2&"';</script>")
			else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"';</script>")
			end if
		Else
			Response.Write("<script language=javascript>this.location.href='Lang.asp?otype="&otype&"&tipinfo="&GBL_CHK_TempStr&"';</script>")
		End If
	End if
end if

Sub ADODB_SaveToFile(ByVal strBody,ByVal File)
	On Error Resume Next
	Dim objStream,FSFlag,fs,WriteFile
	FSFlag = 1
	If DEF_FSOString <> "" Then
		Set fs = Server.CreateObject(DEF_FSOString)
		If Err Then
			FSFlag = 0
			Err.Clear
			Set fs = Nothing
		End If
	Else
		FSFlag = 0
	End If
	
	If FSFlag = 1 Then
		Set WriteFile = fs.CreateTextFile(Server.MapPath(File),True)
		WriteFile.Write strBody
		WriteFile.Close
		Set Fs = Nothing
	Else
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream ，无法完成操作！"
			Err.Clear
			Set objStream = Noting
			Exit Sub
		End If
		With objStream
			.Type = 2
			.Open
			.Charset = "GB2312"
			.Position = objStream.Size
			.WriteText = strBody
			.SaveToFile Server.MapPath(File),2
			.Close
		End With
		Set objStream = Nothing
	End If
End Sub

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>