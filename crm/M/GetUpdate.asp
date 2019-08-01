<!--#include file="inc.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%><%
action = Trim(Request("action"))
sType = Trim(Request("sType"))
cID = Trim(Request("cID"))
ID = Trim(Request("ID"))
tipinfo = Trim(Request("tipinfo"))
YNRange = Trim(Request("YNRange"))
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "/www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="/www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><%=title%></title>
<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=0" />
<meta name="keywords" content="" />
<meta name="description" content="" />
<link rel="stylesheet" type="text/css" href="style/reset.css" /> 
<link rel="stylesheet" type="text/css" href="style/root.css" /> 
<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript" src="js/toogle.js"></script>
</head>
<body>
<!-- start header -->
    <div id="header">
         <a href="index.asp"><img src="img/logo.png" width="120" height="40" alt="logo" class="logo" /></a>
		 <%if sType="View" then%>
         <a style="cursor:pointer" class="button list"><img src="img/list-button.png" width="15" height="16" alt="icon" /></a>
		 <%end if%>
         <a onClick=window.location.href="javascript:window.location.reload();" class="button create"><img src="img/reload-button.png" width="15" height="16" alt="icon" /></a>
        <div class="clear"></div>
	</div>
    <!-- end header -->
	
<script language="javascript">
function check()
{
	var obj = document.loginForm;
	if (obj.item1.value == '')
	{
		obj.item1.focus();
		return false;
	}
	if (obj.item2.value == '')
	{
		obj.item2.focus();
		return false;
	}
	return true;
}
</script>

<%

Select Case action
Case "Client"
    Call Client()
End Select

Sub Client()
cID = Trim(Request("cID"))
%>
    <!-- start page -->
    <div class="page">
	<script language="JavaScript">
	<!-- 客户档案必填项提示
	function CheckInput()
	{
	if (<%=Must_Client_cCompany%>=="1"){if(document.all.Company.value == ""){alert("<%=L_Client_cCompany & alert04%>");document.all.Company.focus();return false;}}
	if (<%=Must_Client_cArea%>=="1"){if(document.all.Area.value == ""){alert("<%=L_Client_cArea & alert04%>");document.all.Area.focus();return false;}}
	if (<%=Must_Client_cSquare%>=="1"){if(document.all.Square.value == ""){alert("<%=L_Client_cSquare & alert04%>");document.all.Square.focus();return false;}}
	if (<%=Must_Client_cAddress%>=="1"){if(document.all.Address.value == ""){alert("<%=L_Client_cAddress & alert04%>");document.all.Address.focus();return false;}}
	if (<%=Must_Client_cZip%>=="1"){if(document.all.Zip.value == ""){alert("<%=L_Client_cZip & alert04%>");document.all.Zip.focus();return false;}}
	if (<%=Must_Client_cLinkman%>=="1"){if(document.all.Linkman.value == ""){alert("<%=L_Client_cLinkman & alert04%>");document.all.Linkman.focus();return false;}}
	if (<%=Must_Client_cZhiwei%>=="1"){if(document.all.Zhiwei.value == ""){alert("<%=L_Client_cZhiwei & alert04%>");document.all.Zhiwei.focus();return false;}}
	if (<%=Must_Client_cMobile%>=="1"){if(document.all.Mobile.value == ""){alert("<%=L_Client_cMobile & alert04%>");document.all.Mobile.focus();return false;}}
	if (<%=Must_Client_cTel%>=="1"){if(document.all.Tel.value == ""){alert("<%=L_Client_cTel & alert04%>");document.all.Tel.focus();return false;}}
	if (<%=Must_Client_cFax%>=="1"){if(document.all.Fax.value == ""){alert("<%=L_Client_cFax & alert04%>");document.all.Fax.focus();return false;}}
	if (<%=Must_Client_cHomepage%>=="1"){if(document.all.Homepage.value == ""){alert("<%=L_Client_cHomepage & alert04%>");document.all.Homepage.focus();return false;}}
	if (<%=Must_Client_cEmail%>=="1"){if(document.all.Email.value == ""){alert("<%=L_Client_cEmail & alert04%>");document.all.Email.focus();return false;}}
	if (<%=Must_Client_cTrade%>=="1"){if(document.all.Trade.value == ""){alert("<%=L_Client_cTrade & alert04%>");document.all.Trade.focus();return false;}}
	if (<%=Must_Client_cStrade%>=="1"){if(document.all.Strade.value == ""){alert("<%=L_Client_cStrade & alert04%>");document.all.Strade.focus();return false;}}
	if (<%=Must_Client_cType%>=="1"){if(document.all.Type.value == ""){alert("<%=L_Client_cType & alert04%>");document.all.Type.focus();return false;}}
	if (<%=Must_Client_cStart%>=="1"){if(document.all.Start.value == ""){alert("<%=L_Client_cStart & alert04%>");document.all.Start.focus();return false;}}
	if (<%=Must_Client_cSource%>=="1"){if(document.all.Source.value == ""){alert("<%=L_Client_cSource & alert04%>");document.all.Source.focus();return false;}}
	if (<%=Must_Client_cInfo%>=="1"){if(document.all.Info.value == ""){alert("<%=L_Client_cInfo & alert04%>");document.all.Info.focus();return false;}}
	if (<%=Must_Client_cBeizhu%>=="1"){if(document.all.Beizhu.value == ""){alert("<%=L_Client_cBeizhu & alert04%>");document.all.Beizhu.focus();return false;}}
	}
		-->
	</script>
<%
if sType="Add" then '添加
%>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Check.min.js"></script>
            <div class="simplebox">
            	<h1 class="titleh">新增客户</h1>
                <div class="content">
                	
                <form name="Add" action="?action=Client&sType=SaveAdd" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cCompany = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cCompany%></label>
					  <input type="text" class="int" name="Company" id="Company" style=" width:60%;" maxlength="50" autocomplete="off" onChange="checkcompany(this.value);"> <span id="check1"> <span class="info_warn help01">唯一标识</span></span>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cArea = 1 or Must_Client_cSquare = 1 or Must_Client_cAddress = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cArea%> / <%=L_Client_cSquare%></label>
                      
									<select name="Area" onchange="getArea(this.options[this.selectedIndex].id);">
									<option value=""><%=L_Please_choose_01%></option>
									<% 
										Set rsb = Conn.Execute("select * from [AreaData] where aFId = '0' ")
										If Not rsb.Eof then
										Do While Not rsb.Eof
										aId= rsb("aId")
										aName= rsb("aName")
									%>
										<option value="<%=aName%>" id="<%=aId%>"><%=aName%></option>
									<%
										rsb.Movenext
										Loop
										End If
										rsb.Close
										Set rss = Nothing 
									%>
									</select> —
									<span id="Squarediv"  style="margin-left:5px;padding:0;">
										<select name="Squares">
											<option value=""><%=L_Please_choose_02%></option>
										</select>
									</span>　<input name="Square" type="hidden" id="Square" class="int">
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cArea = 1 or Must_Client_cSquare = 1 or Must_Client_cAddress = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cAddress%></label>
                      <input name="Address" type="text" class="int" id="Address"  style=" width:80%;" >　
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"> <%=L_Client_cZip%></label>
                      <input name="Zip" type="text" class="int" id="Zip" maxlength="6" onkeyup='this.value=this.value.replace(/\D/gi,"")' style=" width:30%;"> <%=L_Tip_Info_06%>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cLinkman%></label>
                      <input name="Linkman" type="text" class="int" id="Linkman" style=" width:30%;">
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cZhiwei = 1 then %><font color="#FF0000">*</font> <%end if%><%=L_Client_cZhiwei%></label>
                      <% = EasyCrm.getSelect("SelectData","Select_Zhiwei","Zhiwei","") %>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cMobile = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cMobile%></label>
                      <input name="Mobile" type="text" class="int" id="Mobile" onkeyup='this.value=this.value.replace(/\D/gi,"")' style=" width:30%;"> <%=L_Tip_Info_06%>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cTel = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cTel%></label>
                      <input name="Tel" type="text" class="int" id="Tel" style=" width:30%;">
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cFax = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cFax%></label>
                      <input name="Fax" type="text" class="int" id="Fax" style=" width:30%;">
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cHomepage = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cHomepage%></label>
                      <input name="Homepage" type="text" class="int" id="Homepage" style=" width:60%;"> <%=L_Tip_Info_04%>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cEmail = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cEmail%></label>
                      <input name="Email" type="text" class="int" id="Email" style=" width:60%;"> <%=L_Tip_Info_04%>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cTrade = 1 or Must_Client_cStrade = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cTrade%></label>
                      
									<select name="Trade" onchange="getTrade(this.options[this.selectedIndex].id);">
									<option value=""><%=L_Please_choose_01%></option>
									<% 
										Set rsb = Conn.Execute("select * from [ProductClass] where pClassFid = '0' ")
										If Not rsb.Eof then
										Do While Not rsb.Eof
										pClassid= rsb("pClassid")
										pClassname= rsb("pClassname")
									%>
										<option value="<%=pClassname%>" id="<%=pClassid%>"><%=pClassname%></option>
									<%
										rsb.Movenext
										Loop
										End If
										rsb.Close
										Set rsb = Nothing 
									%>
									</select> 
									<span id="Stradediv"  style="margin-left:10px;padding:0;">
										<select name="Strades">
											<option value=""><%=L_Please_choose_02%></option>
										</select>
									</span>
									<input name="Strade" type="hidden" id="Strade">
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cType%></label>
                      <% = EasyCrm.getSelect("SelectData","Select_Type","Type","") %>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cStart = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cStart%></label>
                      <% = EasyCrm.getSelect("SelectData","Select_Star","Start","") %>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cSource = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cSource%></label>
                      <% = EasyCrm.getSelect("SelectData","Select_Source","Source","") %>
                    </div>
					
					<%
						Set rss = Server.CreateObject("ADODB.Recordset")
						rss.Open "Select * From [CustomField] where cTable='Client' order by Id asc ",conn,3,1
						If rss.RecordCount > 0 Then
						Do While Not rss.BOF And Not rss.EOF
					%>
                    <div class="form-line">
						<label class="st-label"><%=rss("cTitle")%></label>
						<%if rss("cType") = "text" then%>
						<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="">
						<%elseif rss("cType") = "time" then%>
						<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="" />
						<%elseif rss("cType") = "select" then%>
						<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
							<option value=""><%=L_Select%></option>
							<%
							selectstr = split(""&rss("cContent")&"",",")
							for selectarr = 0 to ubound(selectstr)
							response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
							next
							%>
						</select>
						<%elseif rss("cType") = "checkbox" then%>
						<%
							checkboxstr = split(""&rss("cContent")&"",",")
							for checkboxarr = 0 to ubound(checkboxstr)
							response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
							next
						%>
						<%elseif rss("cType") = "radio" then%>
						<%
							radiostr = split(""&rss("cContent")&"",",")
							for radioarr = 0 to ubound(radiostr)
							response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
							next
						%>
						<%end if%>
                    </div>
					
					<%
						rss.MoveNext
						Loop
						end if
						rss.Close
						Set rss = Nothing
					%>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cInfo = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cInfo%></label>
                      <textarea name="Info" id="Info" class="int" style="width:80%;"></textarea>
                    </div>

                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cBeizhu = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cBeizhu%></label>
                      <textarea name="Beizhu" id="Beizhu" class="int" style="width:80%;"></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="User" type="hidden" value="<%=Session("CRM_name")%>">
					<input name="Group" type="hidden" value="<%=Session("CRM_group")%>">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input type="reset" name="button" id="button2" value="&nbsp;重 置&nbsp;" class="reset-button" />
                    </div>

                  </form>
                </div>
			</div>
<%
elseif sType="SaveAdd" then
	
	cCompany = Trim(Request("Company"))
	cArea = Trim(Request("Area"))
	cSquare = Trim(Request("Square"))
	cAddress = Trim(Request("Address"))
	cZip = Trim(Request("Zip"))
	cLinkman = Trim(Request("Linkman"))
	cZhiwei = Trim(Request("Zhiwei"))
	cMobile = Trim(Request("Mobile"))
	cTel = Trim(Request("Tel"))
	cFax = Trim(Request("Fax"))
	cHomepage = Trim(Request("Homepage"))
	cEmail = Trim(Request("Email"))  
	cTrade = Trim(Request("Trade"))
	cStrade = Trim(Request("Strade"))
	cType = Trim(Request("Type"))
	cStart = Trim(Request("Start"))
	cSource = Trim(Request("Source"))    
	cInfo = Trim(Request("Info"))
	cBeizhu = Trim(Request("Beizhu"))
	cUser = Trim(Request("User"))
	cGroup = Trim(Request("Group"))

	OnlyItem=""
	if ClientOnly = "100" then
		OnlyItem = OnlyItem & " and cCompany = '" & cCompany & "' "
	elseif ClientOnly = "110" then
		OnlyItem = OnlyItem & " and ( cCompany = '" & cCompany & "' or cLinkman = '" & cLinkman & "' ) "
	elseif ClientOnly = "111" then
		OnlyItem = OnlyItem & " and ( cCompany = '" & cCompany & "' or cLinkman = '" & cLinkman & "' or cMobile = '" & cMobile & "' )  "
	elseif ClientOnly = "101" then
		OnlyItem = OnlyItem & " and ( cCompany = '" & cCompany & "' or cMobile = '" & cMobile & "' )"
	elseif ClientOnly = "011" then
		OnlyItem = OnlyItem & " and ( cLinkman = '" & cLinkman & "' or cMobile = '" & cMobile & "' ) "
	elseif ClientOnly = "001" then
		OnlyItem = OnlyItem & " and cMobile = '" & cMobile & "' "
	end if
	Dim rs,cId
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From client Where 1=1 " & OnlyItem & " ",conn,1,1
	If rs.RecordCount > 0 Then
		Response.Write("<script>alert("""&alert01&""");history.back(1);</script>")
	Response.End()
	End If
	rs.Close
			
	rs.Open "Select Top 1 * From client",conn,3,2
	rs.AddNew

	rs("cCompany") = cCompany
	rs("cArea") = cArea
	rs("cSquare") = cSquare
	rs("cAddress") = cAddress
	rs("cZip") = cZip
	rs("cLinkman") = cLinkman
	rs("cZhiwei") = cZhiwei
	rs("cMobile") = cMobile
	rs("cTel") = cTel
	rs("cFax") = cFax
	rs("cHomepage") = cHomepage
	rs("cEmail") = cEmail
	rs("cTrade") = cTrade
	rs("cStrade") = cStrade
	rs("cType") = cType
	rs("cStart") = cStart
	rs("cSource") = cSource
	rs("cInfo") = cInfo
	rs("cBeizhu") = cBeizhu
	rs("cUser") = cUser
	rs("cGroup") = cGroup
	rs("cLastUpdated") = now()
	
	'写入默认值
	rs("cDate") = Date()
	rs("cYn") = 1
	rs("cShare") = 0

	rs.Update
	rs.Close
	Set rs = Nothing

	Dim rsid
	Set rsid = Server.CreateObject("ADODB.Recordset")
	if Accsql = 0 then
	rsid.Open "Select top 1 cid From client order by cid desc",conn,1,1
	elseif Accsql = 1 then
	rsid.Open "Select @@IDENTITY as cid From client",conn,1,1
	end if
	cid=rsid("cid")
	rsid.close
	
	'插入联系人表
	conn.execute ("insert into Linkmans(cid,lName,lZhiwei,lMobile,lUser,lTime) values('"&cid&"','"&cLinkman&"','"&cZhiwei&"','"&cMobile&"','"&cUser&"','"&now()&"')")	
	
	'插入自定义内容
	
	cContent = ""
	Set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open "Select * From [CustomField] where cTable='Client' order by Id asc ",conn,1,1
	If rsc.RecordCount > 0 Then
	Do While Not rsc.BOF And Not rsc.EOF
	
	cContent = cContent & rsc("cName") &":"& Trim(Request(rsc("cName"))) &"|"
	
	rsc.MoveNext
	Loop
	end if
	rsc.Close
	Set rsc = Nothing
	
	conn.execute ("insert into CustomFieldContent(cID,cContent) values('"&cid&"','"&cContent&"')")	
	
	'插入操作记录
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lReason,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_insert_action_01&"','Mobile','"&cUser&"','"&now()&"')")
	Response.Write("<script>location.href='GetUpdate.asp?action=Client&sType=View&cid="&cid&"' ;</script>")
	Response.end

elseif sType="InfoEdit" then
%>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/ajax.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Check.min.js"></script>
            <div class="simplebox">
            	<h1 class="titleh">修改客户</h1>
                <div class="content">
                	
                <form name="Add" action="?action=Client&sType=SaveEdit" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cCompany = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cCompany%></label>
					  <input type="text" class="int" name="Company" id="Company" value="<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cCompany")%>" style=" width:60%;" maxlength="50" autocomplete="off" <%if Session("CRM_level")<9 then%><%if EasyCrm.getNewItem("Client","cID",""&cID&"","cDate") <> date() then%>readonly="true"<%end if%><%end if%>> <span class="info_warn help01">不可修改</span>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cArea = 1 or Must_Client_cSquare = 1 or Must_Client_cAddress = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cArea%> / <%=L_Client_cSquare%></label>
                      
									<select name="Area" onchange="getArea(this.options[this.selectedIndex].id);">
									<option value=""><%=L_Please_choose_01%></option>
									<% 
										Set rsb = Conn.Execute("select * from [AreaData] where aFId = '0' ")
										If Not rsb.Eof then
										Do While Not rsb.Eof
										aId= rsb("aId")
										aName= rsb("aName")
									%>
										<option value="<%=aName%>" id="<%=aId%>" <%if aName = EasyCrm.getNewItem("Client","cID",""&cID&"","cArea") then %>selected<%end if%> ><%=aName%></option>
									<%
										rsb.Movenext
										Loop
										End If
										rsb.Close
										Set rsb = Nothing 
									%>
									</select> 
									<span id="Squarediv"  style="margin-left:10px;padding:0;">
										<select name="Squares" onchange="getSquare(options[selectedIndex])">
											<option value=""><%=L_Please_choose_02%></option>
											<% 
											IF EasyCrm.getNewItem("Client","cID",""&cID&"","cArea") <>"" then
											Set rss = Conn.Execute("select * from [AreaData] where aFId= '"&EasyCrm.getNewItem("AreaData","aName","'"&EasyCrm.getNewItem("Client","cID",""&cID&"","cArea")&"'","aId")&"' ")
											If Not rss.Eof then
											Do While Not rss.Eof
											aName= rss("aName")
											%>
											<option value="<%=aName%>" <%if aName = EasyCrm.getNewItem("Client","cID",""&cID&"","cSquare") then %>selected<%end if%> ><%=aName%></option>
											<%rss.Movenext
											Loop
											End If
											rss.Close
											Set rss = Nothing 
											End If
											%>
										</select>
									</span>　<input name="Square" type="hidden" id="Square" value="<% = EasyCrm.getNewItem("Client","cID",""&cID&"","cSquare") %>">
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cArea = 1 or Must_Client_cSquare = 1 or Must_Client_cAddress = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cAddress%></label>
					  <input name="Address" type="text" class="int" id="Address" style=" width:80%;" value="<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cAddress")%>" >
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"> <%=L_Client_cZip%></label>
                      <input name="Zip" type="text" class="int" id="Zip" maxlength="6" onkeyup='this.value=this.value.replace(/\D/gi,"")' style=" width:30%;" value="<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cZip")%>"> <%=L_Tip_Info_06%>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cLinkman%></label>
                      <input name="Linkman" type="text" class="int" id="Linkman" style=" width:30%;" value="<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cLinkman")%>">
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cZhiwei = 1 then %><font color="#FF0000">*</font> <%end if%><%=L_Client_cZhiwei%></label>
                      <% = EasyCrm.getSelect("SelectData","Select_Zhiwei","Zhiwei","'"&EasyCrm.getNewItem("Client","cID",""&cID&"","cZhiwei")&"'") %>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cMobile = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cMobile%></label>
                      <input name="Mobile" type="text" class="int" id="Mobile" onkeyup='this.value=this.value.replace(/\D/gi,"")' style=" width:30%;" value="<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cMobile")%>"> <%=L_Tip_Info_06%>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cTel = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cTel%></label>
                      <input name="Tel" type="text" class="int" id="Tel" style=" width:30%;" value="<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cTel")%>">
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cFax = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cFax%></label>
                      <input name="Fax" type="text" class="int" id="Fax" style=" width:30%;" value="<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cFax")%>">
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cHomepage = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cHomepage%></label>
                      <input name="Homepage" type="text" class="int" id="Homepage" style=" width:60%;" value="<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cHomepage")%>"> <%=L_Tip_Info_04%>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cEmail = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cEmail%></label>
                      <input name="Email" type="text" class="int" id="Email" style=" width:60%;" value="<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cEmail")%>"> <%=L_Tip_Info_04%>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cTrade = 1 or Must_Client_cStrade = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cTrade%></label>
                      
									<select name="Trade" onchange="getTrade(this.options[this.selectedIndex].id);">
									<option value=""><%=L_Please_choose_01%></option>
									<% 
										Set rsb = Conn.Execute("select * from [ProductClass] where pClassFid = '0' ")
										If Not rsb.Eof then
										Do While Not rsb.Eof
										pClassid= rsb("pClassid")
										pClassname= rsb("pClassname")
									%>
										<option value="<%=pClassname%>" id="<%=pClassid%>"><%=pClassname%></option>
									<%
										rsb.Movenext
										Loop
										End If
										rsb.Close
										Set rsb = Nothing 
									%>
									</select> 
									<span id="Stradediv"  style="margin-left:10px;padding:0;">
										<select name="Strades" onchange="getStrade(options[selectedIndex])">
											<option value=""><%=L_Please_choose_02%></option>
											<% 
											IF EasyCrm.getNewItem("Client","cID",""&cID&"","cTrade")<>"" then
											Set rsb = Conn.Execute("select * from [ProductClass] where pClassFid='"&EasyCrm.getNewItem("ProductClass","pClassname","'"&EasyCrm.getNewItem("Client","cID",""&cID&"","cTrade")&"'","pClassId")&"' ")
											If Not rsb.Eof then
											Do While Not rsb.Eof
											pClassname= rsb("pClassname")
											%>
											<option value="<%=pClassname%>"><%=pClassname%></option>
											<%rsb.Movenext
											Loop
											End If
											rsb.Close
											Set rsb = Nothing 
											end if
											%>
										</select>
									</span>
									<input name="Strade" type="hidden" id="Strade" value="<% = EasyCrm.getNewItem("Client","cID",""&cID&"","cStrade") %>">
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cType%></label>
                      <% = EasyCrm.getSelect("SelectData","Select_Type","Type",""&EasyCrm.getNewItem("Client","cID",""&cID&"","cType")&"") %>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cStart = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cStart%></label>
                      <% = EasyCrm.getSelect("SelectData","Select_Star","Start",""&EasyCrm.getNewItem("Client","cID",""&cID&"","cStart")&"") %>
                    </div>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cSource = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cSource%></label>
                      <% = EasyCrm.getSelect("SelectData","Select_Source","Source",""&EasyCrm.getNewItem("Client","cID",""&cID&"","cSource")&"") %>
                    </div>
					
					<%
								cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID",""&cID&"","cContent")
								cContentArr = split(cContentStr,"|")								
								Set rss = Server.CreateObject("ADODB.Recordset")
								rss.Open "Select * From [CustomField] where cTable='Client' order by Id asc ",conn,3,1
								If rss.RecordCount > 0 Then
								k=0
								Do While Not rss.BOF And Not rss.EOF
								if Ubound(cContentArr) > k then
								cContent = split(cContentArr(k),":")
					%>
                    <div class="form-line">
						<label class="st-label"><%=rss("cTitle")%></label>
									<%if inStr(cContentArr(k),cContent(0))>0 then%>
										<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="<%=cContent(1)%>">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="<%=cContent(1)%>" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											if selectstr(selectarr) = cContent(1) then
											response.Write "<option value="""&selectstr(selectarr)&""" selected>"&selectstr(selectarr)&"</option>"
											else
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											end if
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											if inStr(cContent(1),checkboxstr(checkboxarr))>0 then
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&""" checked> "&checkboxstr(checkboxarr)&"　"
											else
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											end if
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											if radiostr(radioarr) = cContent(1) then
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&""" checked> "&radiostr(radioarr)&"　"
											else
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											end if
											next
											%>
										<%end if%>
									<%end if%>
                    </div>
								<%
								else
								%>
                    <div class="form-line">
						<label class="st-label"><%=rss("cTitle")%></label>
									<%if rss("cType") = "text" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px" value="">
										<%elseif rss("cType") = "time" then%>
										<input name="<%=rss("cName")%>" type="text" id="<%=rss("cName")%>" class="Wdate" style="width:<%=rss("cWidth")%>px" onFocus="WdatePicker()" value="" />
										<%elseif rss("cType") = "select" then%>
											<select name="<%=rss("cName")%>" class="int" style="width:<%=rss("cWidth")%>px">
											<option value=""><%=L_Select%></option>
											<%
											selectstr = split(""&rss("cContent")&"",",")
											for selectarr = 0 to ubound(selectstr)
											response.Write "<option value="""&selectstr(selectarr)&""">"&selectstr(selectarr)&"</option>"
											next
											%>
											</select>
										<%elseif rss("cType") = "checkbox" then%>
											<%
											checkboxstr = split(""&rss("cContent")&"",",")
											for checkboxarr = 0 to ubound(checkboxstr)
											response.Write "<input name="""&rss("cName")&""" type=""checkbox"" value="""&checkboxstr(checkboxarr)&"""> "&checkboxstr(checkboxarr)&"　"
											next
											%>
										<%elseif rss("cType") = "radio" then%>
											<%
											radiostr = split(""&rss("cContent")&"",",")
											for radioarr = 0 to ubound(radiostr)
											response.Write "<input name="""&rss("cName")&""" type=""radio"" value="""&radiostr(radioarr)&"""> "&radiostr(radioarr)&"　"
											next
											%>
										<%end if%>
                    </div>
								<%
								end if
								k=k+1
								rss.MoveNext
								Loop
								end if
								rss.Close
								Set rss = Nothing
							%>
                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cInfo = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cInfo%></label>
                      <textarea name="Info" id="Info" class="int" style="width:80%;"><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cInfo")%></textarea>
                    </div>

                  		
                    <div class="form-line">
                   	  <label class="st-label"><%if Must_Client_cBeizhu = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Client_cBeizhu%></label>
                      <textarea name="Beizhu" id="Beizhu" class="int" style="width:80%;"><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cBeizhu")%></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="cId" type="hidden" id="cId" value="<% = cId %>">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input name="Back" type="button" id="Back" class="reset-button" value="返回" onClick="location.href='listall.asp?PN=<%=Session("CRM_pagenum")%>';">
                    </div>

                  </form>
                </div>
			</div>
<script language="JavaScript">
<!--
for(var i=0;i<document.all.Area.options.length;i++){
    if(document.all.Area.options[i].value == "<% = EasyCrm.getNewItem("Client","cID",""&cID&"","cArea") %>"){
    document.all.Area.options[i].selected = true;}}

for(var i=0;i<document.all.Squares.options.length;i++){
    if(document.all.Squares.options[i].value == "<% = EasyCrm.getNewItem("Client","cID",""&cID&"","cSquare") %>"){
    document.all.Squares.options[i].selected = true;}}

for(var i=0;i<document.all.Trade.options.length;i++){
    if(document.all.Trade.options[i].value == "<% = EasyCrm.getNewItem("Client","cID",""&cID&"","cTrade") %>"){
    document.all.Trade.options[i].selected = true;}}

for(var i=0;i<document.all.Strades.options.length;i++){
    if(document.all.Strades.options[i].value == "<% = EasyCrm.getNewItem("Client","cID",""&cID&"","cStrade") %>"){
    document.all.Strades.options[i].selected = true;}}
-->
</script>
<%
elseif sType="SaveEdit" then
	cId = CLng(ABS(Request("cId")))
	cCompany = Trim(Request("Company"))
	cArea = Trim(Request("Area"))
	
	if Trim(Request("Squares"))<>"" then 
		cSquare = Trim(Request("Squares"))
	else
		if Trim(Request("Square")) <> "" then 
		cSquare = Trim(Request("Square"))
		else
		cSquare = ""
		end if
	end if
	
	cAddress = Trim(Request("Address"))
	cZip = Trim(Request("Zip"))
	cLinkman = Trim(Request("Linkman"))
	cZhiwei = Trim(Request("Zhiwei"))
	cMobile = Trim(Request("Mobile"))
	cTel = Trim(Request("Tel"))
	cFax = Trim(Request("Fax"))
	cHomepage = Trim(Request("Homepage"))
	cEmail = Trim(Request("Email"))  
	cTrade = Trim(Request("Trade"))
	if Trim(Request("Strades"))<>"" then 
		cStrade = Trim(Request("Strades"))
	else
		if Trim(Request("Strade")) <> "" then 
		cstrade = Trim(Request("Strade"))
		else
		cstrade = ""
		end if
	end if
	cType = Trim(Request("Type"))
	cStart = Trim(Request("Start"))
	cSource = Trim(Request("Source"))    
	cInfo = Trim(Request("Info"))
	cBeizhu = Trim(Request("Beizhu"))

	OnlyItem=""
	if ClientOnly = "100" then
		OnlyItem = OnlyItem & " and cCompany = '" & cCompany & "' "
	elseif ClientOnly = "110" then
		OnlyItem = OnlyItem & " and ( cCompany = '" & cCompany & "' or cLinkman = '" & cLinkman & "' ) "
	elseif ClientOnly = "111" then
		OnlyItem = OnlyItem & " and ( cCompany = '" & cCompany & "' or cLinkman = '" & cLinkman & "' or cMobile = '" & cMobile & "' )  "
	elseif ClientOnly = "101" then
		OnlyItem = OnlyItem & " and ( cCompany = '" & cCompany & "' or cMobile = '" & cMobile & "' )"
	elseif ClientOnly = "011" then
		OnlyItem = OnlyItem & " and ( cLinkman = '" & cLinkman & "' or cMobile = '" & cMobile & "' ) "
	elseif ClientOnly = "001" then
		OnlyItem = OnlyItem & " and cMobile = '" & cMobile & "' "
	end if
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From client Where 1=1 " & OnlyItem & " And cId <> " & cId ,conn,1,1
	If rs.RecordCount > 0 Then
		Response.Write("<script>alert("""&alert01&""");history.back(1);</script>")
	Response.End()
	End If
	rs.Close
			
	rs.Open "Select Top 1 * From client Where cId = " & cId ,conn,3,2

	rs("cCompany") = cCompany
	rs("cArea") = cArea
	rs("cSquare") = cSquare
	rs("cAddress") = cAddress
	rs("cZip") = cZip
	rs("cLinkman") = cLinkman
	rs("cZhiwei") = cZhiwei
	rs("cMobile") = cMobile
	rs("cTel") = cTel
	rs("cFax") = cFax
	rs("cHomepage") = cHomepage
	rs("cEmail") = cEmail
	rs("cTrade") = cTrade
	rs("cStrade") = cStrade
	rs("cType") = cType
	rs("cStart") = cStart
	rs("cSource") = cSource
	rs("cInfo") = cInfo
	rs("cBeizhu") = cBeizhu
	rs("cLastUpdated") = now()

	rs.Update
	rs.Close
	Set rs = Nothing
	
	'同步更新联系人表第一条记录
	conn.execute ("UPDATE [Linkmans] SET lName='"&cLinkman&"',lZhiwei='"&cZhiwei&"',lMobile='"&cMobile&"' Where cId ="&cId&" and lid in ( select top 1 lid from [Linkmans] where cid="&cId&" ) ")
	
	'更新自定义内容
	
	cContent = ""
	Set rsc = Server.CreateObject("ADODB.Recordset")
	rsc.Open "Select * From [CustomField] where cTable='Client' order by Id asc ",conn,3,1
	If rsc.RecordCount > 0 Then
	Do While Not rsc.BOF And Not rsc.EOF
	'获取所有自定义字段
	cContent = cContent & rsc("cName") &":"& Trim(Request(rsc("cName"))) &"|" 
	rsc.MoveNext
	Loop
	end if
	rsc.Close
	Set rsc = Nothing
	if EasyCrm.getNewItem("CustomFieldContent","cID",""&cID&"","cContent")="0" then
	conn.execute ("insert into CustomFieldContent(cID,cContent) values('"&cid&"','"&cContent&"')")	
	else
	conn.execute ("UPDATE [CustomFieldContent] SET cContent='"&cContent&"' Where cId ="&cId&" ")
	end if
	
	'插入操作记录
	conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Client&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
	Response.Write("<script>location.href='GetUpdate.asp?action=Client&sType=View&cid="&cid&"' ;</script>")
	Response.end

elseif sType="View" then
	otype	=	Request.QueryString("otype")
%>
            <div class="simplebox">
            	<h1 class="titleh" onclick="location.href='?action=Client&sType=View&cid=<%=cid%>'" style="cursor:pointer">客户基本档案</h1>
<div class="listbox" style="margin:0px;border:1px solid #C1D6E6;">
<input type="button" class="reset-button" value="联系" onclick="location.href='?action=Client&sType=View&otype=Linkmans&cid=<%=cid%>'" style="cursor:pointer" />
<input type="button" class="reset-button" value="跟单" onclick="location.href='?action=Client&sType=View&otype=Records&cid=<%=cid%>'" style="cursor:pointer" />
<input type="button" class="reset-button" value="订单" onclick="location.href='?action=Client&sType=View&otype=Order&cid=<%=cid%>'" style="cursor:pointer" />
<input type="button" class="reset-button" value="合同" onclick="location.href='?action=Client&sType=View&otype=Hetong&cid=<%=cid%>'" style="cursor:pointer" />
<input type="button" class="reset-button" value="售后" onclick="location.href='?action=Client&sType=View&otype=Service&cid=<%=cid%>'" style="cursor:pointer" />
<input type="button" class="reset-button" value="费用" onclick="location.href='?action=Client&sType=View&otype=Expense&cid=<%=cid%>'" style="cursor:pointer" />
<input type="button" class="reset-button" value="共享" onclick="location.href='?action=Client&sType=View&otype=Share&cid=<%=cid%>'" style="cursor:pointer" />
<input type="button" class="reset-button" value="日志" onclick="location.href='?action=Client&sType=View&otype=History&cid=<%=cid%>'" style="cursor:pointer" />
</div>
			<%if otype="Client" or otype="" then%>
                <div class="content">
					<table class="tabledata"> 
						<col width="80">
						<tbody> 
                        <tr> 
							<td class="lr"><%=L_Client_cCompany%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cCompany")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cArea%> / <%=L_Client_cSquare%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cArea")%>&nbsp;&nbsp;<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cSquare")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cAddress%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cAddress")%>　<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cZip")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cLinkman%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cLinkman")%>&nbsp;&nbsp;<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cZhiwei")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cMobile%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cMobile")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cTel%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cTel")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cFax%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cFax")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cHomepage%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cHomepage")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cEmail%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cEmail")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cTrade%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cTrade")%>&nbsp;&nbsp;<%=EasyCrm.getNewItem("Client","cID",""&cID&"","cStrade")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cType%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cType")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cStart%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cStart")%> </td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cSource%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cSource")%></td> 
                        </tr> 
							<%
								cContentStr = EasyCrm.getNewItem("CustomFieldContent","cID",""&cID&"","cContent")
								cContentArr = split(cContentStr,"|")								
								Set rss = Server.CreateObject("ADODB.Recordset")
								rss.Open "Select * From [CustomField] where cTable='Client' order by Id asc ",conn,3,1
								If rss.RecordCount > 0 Then
								k=0
								Do While Not rss.BOF And Not rss.EOF
								if Ubound(cContentArr) > k then
								cContent = split(cContentArr(k),":")
								%>
                        <tr> 
									<td class="lr"><%=rss("cTitle")%></td>
									<td> 
									<%if inStr(cContentArr(k),cContent(0))>0 then%>
										<%=cContent(1)%>
									<%end if%>
									</td>
                        </tr> 
								<%
								else
								%>
                        <tr> 
									<td class="lr"><%=rss("cTitle")%></td>
									<td class="td_r_l"> 
									</td>
                        </tr> 
								<%
								end if
								k=k+1
								rss.MoveNext
								Loop
								end if
								rss.Close
								Set rss = Nothing
							%>
                        <tr> 
							<td class="lr"><%=L_Client_cInfo%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cInfo")%></td> 
                        </tr> 
                        <tr> 
							<td class="lr"><%=L_Client_cBeizhu%></td> 
							<td><%=EasyCrm.getNewItem("Client","cID",""&cID&"","cBeizhu")%></td> 
                        </tr> 
                    </table> 
                    <div class="form-line" style="margin-top:20px;">
					<% If mid(Session("CRM_qx"), 18, 1) = 1 Then %>
					<%if YNRange = "" then%>
					<input type="button" class="submit-button" value=" 编辑 " onclick="location.href='GetUpdate.asp?action=Client&sType=InfoEdit&cid=<%=cId%>'" style="cursor:pointer" />　
					<%end if%>
					<%end if%>
                    <input name="Back" type="button" id="Back" class="reset-button" value=" 返回 " onClick="location.href='<%=Session("CRM_thispage")%>?PN=<%=Session("CRM_pagenum")%>';">
                    </div>
                </div>
			<%elseif otype="Linkmans" then%>
                <div class="content">
					<input type="button" class="add-button" value=" 新增联系人 " onclick="location.href='?action=Client&sType=View&otype=LinkmansAdd&cid=<%=cid%>'" style="cursor:pointer" /> 
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [linkmans] where cId = "&cId&" Order By lId asc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody> 
                        <tr> 
							<th><%=L_Linkmans_lName%></th> 
							<td><%=rs("lName")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Linkmans_lMobile%></th> 
							<td><%=rs("lMobile")%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Linkmans_lTel%></th> 
							<td><%=rs("lTel")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Linkmans_lQQ%></th> 
							<td><%=rs("lQQ")%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Linkmans_lContent%></th> 
							<td><%=rs("lContent")%></td> 
                        </tr>  
                        <tr> 
							<td colspan=2>
								<input type="button" class="submit-button" value=" 编辑 " onclick="location.href='?action=Client&sType=View&otype=LinkmansEdit&cId=<%=rs("cId")%>&id=<%=rs("lid")%>'" style="cursor:pointer" /> 
								<input type="button" class="reset-button" value=" 删除 " onclick="location.href='?action=Client&sType=View&otype=LinkmansDel&cId=<%=rs("cId")%>&id=<%=rs("lid")%>'" style="cursor:pointer" /> 
							</td> 
                        </tr>  
                    </table>
					<%
						rs.MoveNext
						Loop
						else
					%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody> 
                        <tr> 
							<td>暂无数据</td> 
                        </tr> 
                    </table>
						
					<%
							end if
						rs.Close
						Set rs = Nothing
					%>
				</div>
			<%elseif otype="LinkmansAdd" then%>
			
	<script language="JavaScript">
	<!-- 联系人必填项提示
	function CheckInput()
	{
		if (<%=Must_Linkmans_lName%>=="1"){if(document.all.lName.value == ""){alert("<%=L_Linkmans_lName & alert04%>");document.all.lName.focus();return false;}}
		if (<%=Must_Linkmans_lSex%>=="1"){if(document.all.lSex.value == ""){alert("<%=L_Linkmans_lSex & alert04%>");document.all.lSex.focus();return false;}}
		if (<%=Must_Linkmans_lZhiwei%>=="1"){if(document.all.lZhiwei.value == ""){alert("<%=L_Linkmans_lZhiwei & alert04%>");document.all.lZhiwei.focus();return false;}}
		if (<%=Must_Linkmans_lBirthday%>=="1"){if(document.all.lBirthday.value == ""){alert("<%=L_Linkmans_lBirthday & alert04%>");document.all.lBirthday.focus();return false;}}
		if (<%=Must_Linkmans_lMobile%>=="1"){if(document.all.lMobile.value == ""){alert("<%=L_Linkmans_lMobile & alert04%>");document.all.lMobile.focus();return false;}}
		if (<%=Must_Linkmans_lTel%>=="1"){if(document.all.lTel.value == ""){alert("<%=L_Linkmans_lTel & alert04%>");document.all.lTel.focus();return false;}}
		if (<%=Must_Linkmans_lEmail%>=="1"){if(document.all.lEmail.value == ""){alert("<%=L_Linkmans_lEmail & alert04%>");document.all.lEmail.focus();return false;}}
		if (<%=Must_Linkmans_lQQ%>=="1"){if(document.all.lQQ.value == ""){alert("<%=L_Linkmans_lQQ & alert04%>");document.all.lQQ.focus();return false;}}
		if (<%=Must_Linkmans_lMSN%>=="1"){if(document.all.lMSN.value == ""){alert("<%=L_Linkmans_lMSN & alert04%>");document.all.lMSN.focus();return false;}}
		if (<%=Must_Linkmans_lALWW%>=="1"){if(document.all.lALWW.value == ""){alert("<%=L_Linkmans_lALWW & alert04%>");document.all.lALWW.focus();return false;}}
		if (<%=Must_Linkmans_lContent%>=="1"){if(document.all.lContent.value == ""){alert("<%=L_Linkmans_lContent & alert04%>");document.all.lContent.focus();return false;}}
	}
	-->
	</script>
                <div class="content">
            	<blockquote>新增联系人</blockquote>
				<form name="Save" action="?action=Client&sType=View&otype=LinkmansAddsave" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lName = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lName%></label>
						<input name="lName" type="text" class="int" id="lName" style=" width:30%;">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lSex = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lSex%></label>
						<% = EasyCrm.getSelect("SelectData","Select_Sex","lSex","") %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lZhiwei = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lZhiwei%></label>
						<% = EasyCrm.getSelect("SelectData","Select_Zhiwei","lZhiwei","") %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lBirthday = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lBirthday%></label>
						<input name="lBirthday" type="date" id="lBirthday" class="Wdate" size="22" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" />
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lMobile = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lMobile%></label>
						<input name="lMobile" type="text" class="int" id="lMobile" style=" width:30%;">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lTel = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lTel%></label>
						<input name="lTel" type="text" class="int" id="lTel" style=" width:30%;">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lEmail = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lEmail%></label>
						<input name="lEmail" type="text" class="int" id="lEmail" style=" width:30%;">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lQQ = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lQQ%></label>
						<input name="lQQ" type="text" class="int" id="lQQ" style=" width:30%;">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lMSN = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lMSN%></label>
						<input name="lMSN" type="text" class="int" id="lMSN" style=" width:30%;">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lALWW = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lALWW%></label>
						<input name="lALWW" type="text" class="int" id="lALWW" style=" width:30%;">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lContent%></label>
						<textarea name="lContent" id="lContent" class="int" style="width:80%;"></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="cID" type="hidden" value="<%=cID%>">
					<input name="lUser" type="hidden" value="<%=Session("CRM_name")%>">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input name="Back" type="button" id="Back" class="reset-button" value="返回" onClick="location.href='javascript:history.back();';">
                    </div>
					
                </form>
				</div>
			
			<%elseif otype="LinkmansAddsave" then
				cID = Request.Form("cID")
				lName = Request.Form("lName")
				lSex = Request.Form("lSex")
				lZhiwei = Request.Form("lZhiwei")
				lBirthday = Request.Form("lBirthday")
				lMobile = Request.Form("lMobile")
				lTel = Request.Form("lTel")
				lEmail = Request.Form("lEmail")
				lQQ = Request.Form("lQQ")
				lMSN = Request.Form("lMSN")
				lALWW = Request.Form("lALWW")
				lContent = Request.Form("lContent")
				lUser = Request.Form("lUser")
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open "Select * From [Linkmans] Where lName = '"&lName&"' and cID="&cID&" ",conn,1,1
				If rs.RecordCount > 0 Then
					Response.Write("<script>alert('该联系人已存在，请重新输入！');history.back(1);</script>")
				Response.End()
				End If
				rs.Close
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open "Select Top 1 * From [Linkmans]",conn,3,2
				rs.AddNew
				rs("cID") = cID
				rs("lName") = lName
				rs("lSex") = lSex
				rs("lZhiwei") = lZhiwei
				if lBirthday <>"" then
				rs("lBirthday") = lBirthday
				end if
				rs("lMobile") = lMobile
				rs("lTel") = lTel
				rs("lEmail") = lEmail
				rs("lQQ") = lQQ
				rs("lMSN") = lMSN
				rs("lALWW") = lALWW
				rs("lContent") = lContent
				rs("lUser") = lUser
				rs("lTime") = now()
				rs.Update
				rs.Close
				Set rs = Nothing
				
				conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
				'插入操作记录
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Linkmans&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
				Response.Write("<script>location.href='?action=Client&sType=View&otype=Linkmans&cid="&cid&"' ;</script>")
			
			%>
			<%elseif otype="LinkmansEdit" then%>
			
                <div class="content">
            	<blockquote>修改联系人</blockquote>
				<form name="Save" action="?action=Client&sType=View&otype=LinkmansEditsave" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lName = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lName%></label>
						<input name="lName" type="text" class="int" id="lName" style=" width:30%;" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lName")%>">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lSex = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lSex%></label>
						<% = EasyCrm.getSelect("SelectData","Select_Sex","lSex","'"&EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lSex")&"'") %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lZhiwei = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lZhiwei%></label>
						<% = EasyCrm.getSelect("SelectData","Select_Zhiwei","lZhiwei",EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lZhiwei")) %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lBirthday = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lBirthday%></label>
						<input name="lBirthday" type="date" id="lBirthday" class="Wdate" size="22" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lBirthday"),2)%>" />
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lMobile = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lMobile%></label>
						<input name="lMobile" type="text" class="int" id="lMobile" style=" width:30%;" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lMobile")%>">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lTel = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lTel%></label>
						<input name="lTel" type="text" class="int" id="lTel" style=" width:30%;" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lTel")%>">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lEmail = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lEmail%></label>
						<input name="lEmail" type="text" class="int" id="lEmail" style=" width:30%;" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lEmail")%>">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lQQ = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lQQ%></label>
						<input name="lQQ" type="text" class="int" id="lQQ" style=" width:30%;" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lQQ")%>">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lMSN = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lMSN%></label>
						<input name="lMSN" type="text" class="int" id="lMSN" style=" width:30%;" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lMSN")%>">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lALWW = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lALWW%></label>
						<input name="lALWW" type="text" class="int" id="lALWW" style=" width:30%;" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lALWW")%>">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Linkmans_lContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Linkmans_lContent%></label>
						<textarea name="lContent" id="lContent" class="int" style="width:80%;"><%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","lContent")%></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="lID" type="hidden" value="<%=ID%>">
					<input name="cID" type="hidden" value="<%=EasyCrm.getNewItem("Linkmans","lID",""&ID&"","cID")%>">
					<input name="YNUpdate" type="hidden" value="<%=YNUpdate%>">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input name="Back" type="button" id="Back" class="reset-button" value="返回" onClick="location.href='javascript:history.back();';">
                    </div>
					
                </form>
				</div>
			<%elseif otype="LinkmansEditsave" then
				cID = Request.Form("cID")
				lID = Request.Form("lID")
				lName = Request.Form("lName")
				lSex = Request.Form("lSex")
				lZhiwei = Request.Form("lZhiwei")
				lBirthday = Request.Form("lBirthday")
				lMobile = Request.Form("lMobile")
				lTel = Request.Form("lTel")
				lEmail = Request.Form("lEmail")
				lQQ = Request.Form("lQQ")
				lMSN = Request.Form("lMSN")
				lALWW = Request.Form("lALWW")
				lContent = Request.Form("lContent")
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open "Select * From [Linkmans] Where lName = '"&lName&"' and cID="&cID&" and lID<>"&lID&" ",conn,1,1
				If rs.RecordCount > 0 Then
					Response.Write("<script>alert('该联系人已存在，请重新输入！');history.back(1);</script>")
				Response.End()
				End If
				rs.Close
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open "Select Top 1 * From [Linkmans] where lID="&lID,conn,3,2
				rs("lName") = lName
				rs("lSex") = lSex
				rs("lZhiwei") = lZhiwei
				if lBirthday <>"" then
				rs("lBirthday") = lBirthday
				end if
				rs("lMobile") = lMobile
				rs("lTel") = lTel
				rs("lEmail") = lEmail
				rs("lQQ") = lQQ
				rs("lMSN") = lMSN
				rs("lALWW") = lALWW
				rs("lContent") = lContent
				rs.Update
				rs.Close
				Set rs = Nothing
				
				conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
				if YNUpdate="1" then
				conn.execute ("UPDATE [client] SET cLinkman='"&lName&"',cZhiwei='"&lZhiwei&"',cMobile='"&lMobile&"' Where cId ="&cId&" ")
				end if
				
				conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
				
				'插入操作记录
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Linkmans&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
				Response.Write("<script>location.href='?action=Client&sType=View&otype=Linkmans&cid="&cid&"' ;</script>")
				
			%>
			<%elseif otype="LinkmansDel" then
				If Id = "" Then Exit Sub
				cID = EasyCrm.getNewItem("Linkmans","lID",""&ID&"","cID")
				conn.execute("DELETE FROM [Linkmans] where lId = "&Id&" ")
				
				'插入操作记录
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cID&"','"&L_Linkmans&"','"&L_insert_action_03&"','"&Session("CRM_name")&"','"&now()&"')")	
				Response.Write("<script>location.href='?action=Client&sType=View&otype=Linkmans&cid="&cid&"' ;</script>")
			%>
			
			<%elseif otype="Records" then%>
                <div class="content">
					<input type="button" class="add-button" value=" 新增跟单记录 " onclick="location.href='?action=Client&sType=View&otype=RecordsAdd&cid=<%=cid%>'" style="cursor:pointer" /> 
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Records] where cId = "&cId&" Order By rId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody>  
                        <tr> 
							<th><%=L_Records_rTime%></th> 
							<td><%=rs("rTime")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Records_rLinkman%></th> 
							<td><%=rs("rLinkman")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Records_rType%></th> 
							<td><%=rs("rType")%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Records_rState%></th> 
							<td><%=rs("rState")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Records_rNextTime%></th> 
							<td><%=rs("rNextTime")%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Records_rContent%></th> 
							<td><%=rs("rContent")%></td> 
                        </tr>  
                        <tr> 
							<td colspan=2>
								<input type="button" class="submit-button" value=" 编辑 " onclick="location.href='?action=Client&sType=View&otype=RecordsEdit&cId=<%=rs("cId")%>&id=<%=rs("rid")%>'" style="cursor:pointer" /> 
								<input type="button" class="reset-button" value=" 删除 " onclick="location.href='?action=Client&sType=View&otype=RecordsDel&cId=<%=rs("cId")%>&id=<%=rs("rid")%>'" style="cursor:pointer" /> 
							</td> 
                        </tr>  
                    </table>
					<%
						rs.MoveNext
						Loop
						else
					%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody> 
                        <tr> 
							<td>暂无数据</td> 
                        </tr> 
                    </table>
						
					<%
							end if
						rs.Close
						Set rs = Nothing
					%>
				</div>
			<%elseif otype="RecordsAdd" then%>
	<script language="JavaScript">
	<!-- 必填项提示
	function CheckInput()
	{
		if (<%=Must_Records_rType%>=="1"){if(document.all.rType.value == ""){alert("<%=L_Records_rType & alert04%>");document.all.rType.focus();return false;}}
		if (<%=Must_Records_rState%>=="1"){if(document.all.rState.value == ""){alert("<%=L_Records_rState & alert04%>");document.all.rState.focus();return false;}}
		if (<%=Must_Records_rLinkman%>=="1"){if(document.all.rLinkman.value == ""){alert("<%=L_Records_rLinkman & alert04%>");document.all.rLinkman.focus();return false;}}
		if (<%=Must_Records_rNextTime%>=="1"){if(document.all.rNextTime.value == ""){alert("<%=L_Records_rNextTime & alert04%>");document.all.rNextTime.focus();return false;}}
		if (<%=Must_Records_rRemind%>=="1"){if(document.all.rRemind.value == ""){alert("<%=L_Records_rRemind & alert04%>");document.all.rRemind.focus();return false;}}
		if (<%=Must_Records_rContent%>=="1"){if(document.all.rContent.value == ""){alert("<%=L_Records_rContent & alert04%>");document.all.rContent.focus();return false;}}
	}
	-->
	</script>
			
                <div class="content">
            	<blockquote>新增跟单记录</blockquote>
				<form name="Save" action="?action=Client&sType=View&otype=RecordsAddsave" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
						<label class="st-label"><%if Must_Records_rType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rType%></label>
						<% = EasyCrm.getSelect("SelectData","Select_Records","rType","") %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Records_rState = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rState%></label>
						<% = EasyCrm.getSelect("SelectData","Select_Type","rState","") %>
						<span class="info_help help01" >&nbsp;同步客户类型</span>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Records_rLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rLinkman%></label>
						<% = EasyCrm.getNewSelect("linkmans","lName","rLinkman"," and cid="&cid&" ","") %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Records_rNextTime = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rNextTime%></label>
						<input name="rNextTime" type="date" id="rNextTime" class="Wdate" size="22" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:00:00'})" />
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Records_rContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rContent%></label>
						<textarea name="rContent" id="rContent" class="int" style="width:80%;"></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="cID" type="hidden" value="<%=cID%>">
					<input name="rUser" type="hidden" value="<%=Session("CRM_name")%>">
					<input name="cType" type="hidden" id="cType" value="<%=EasyCrm.getNewItem("Client","cId",""&cID&"","cType")%>">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input type="reset" name="button" id="button2" value="&nbsp;重 置&nbsp;" class="reset-button" />
                    </div>
					
                </form>
				</div>
			<%elseif otype="RecordsAddsave" then
				cID = Request("cID")
				rType = Request("rType")
				rState = Request("rState")
				rlinkman = Request("rlinkman")
				rNextTime = Request("rNextTime")
				rRemind = Request("rRemind")
				rContent = Request("rContent")
				rUser = Request("rUser")
				cType = Request("cType")
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open "Select Top 1 * From [Records]",conn,3,2
				rs.AddNew
				rs("cID") = cID
				rs("rType") = rType
				rs("rState") = rState
				rs("rlinkman") = rlinkman
				if rNextTime <>"" then
				rs("rNextTime") = rNextTime
				end if
				if rRemind <>"" then
				rs("rRemind") = rRemind
				else
				rs("rRemind") = 1
				end if
				rs("rContent") = rContent
				rs("rUser") = rUser
				rs("rTime") = now()
				rs.Update
				rs.Close
				Set rs = Nothing
		
				'同步更新客户类型和插入定时站内信
				
				if ""&rState&"" <> "" and ""&cType&"" <> ""&CRTypeEnd&"" then
				conn.execute ("UPDATE client SET cType='"&rState&"' Where cId ="&cId&" ")
				end if
		
				if rNextTime <> "" then
				RemindTime = Dateadd("h",-rRemind,rNextTime)
				conn.execute ("UPDATE client SET cRNextTime='"&rNextTime&"' Where cId ="&cId&" ")
				conn.execute ("insert into OA_mms_Receive(oReceiver,oSender,oTitle,oContent,oIsread,oAttime,oTime) values('"&Session("CRM_name")&"','系统通知','["&EasyCrm.getNewItem("Client","cid",""&cID&"","cCompany")&"] 于 ["&RemindTime&"] 需再次跟单!','<a href=""GetUpdate.asp?action=Client&sType=View&cid="&cID&""">点击查看</a>',0,'"&RemindTime&"','"&now()&"')")	
				end if
		
				conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
				'插入操作记录
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Records&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")
				Response.Write("<script>location.href='?action=Client&sType=View&otype=Records&cid="&cid&"' ;</script>")
			
			%>
			<%elseif otype="RecordsEdit" then%>
			
                <div class="content">
            	<blockquote>修改跟单记录</blockquote>
				<form name="Save" action="?action=Client&sType=View&otype=RecordsEditsave" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
						<label class="st-label"><%if Must_Records_rType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rType%></label>
						<% = EasyCrm.getSelect("SelectData","Select_Records","rType",""&EasyCrm.getNewItem("Records","rID",""&ID&"","rType")&"") %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Records_rState = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rState%></label>
						<% = EasyCrm.getSelect("SelectData","Select_Type","rState",""&EasyCrm.getNewItem("Records","rID",""&ID&"","rState")&"") %>
						<span class="info_help help01" >&nbsp;同步客户类型</span>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Records_rLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rLinkman%></label>
						<% = EasyCrm.getNewSelect("Linkmans","lName","rLinkman"," and cid="&EasyCrm.getNewItem("Records","rID",""&ID&"","cID")&" ",""&EasyCrm.getNewItem("Records","rID",""&ID&"","rLinkman")&"") %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Records_rNextTime = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rNextTime%></label>
						<input name="rNextTime" type="date" id="rNextTime" class="Wdate" size="22" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Records","rID",""&ID&"","rNextTime"),2)%>" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:00:00'})" />
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Records_rContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Records_rContent%></label>
						<textarea name="rContent" id="rContent" class="int" style="width:80%;"><%=EasyCrm.getNewItem("Records","rID",""&ID&"","rContent")%></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="rID" type="hidden" value="<%=ID%>">
					<input name="cType" type="hidden" id="cType" value="<%=EasyCrm.getNewItem("Client","cId",EasyCrm.getNewItem("Records","rID",""&ID&"","cID"),"cType")%>">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input type="reset" name="button" id="button2" value="&nbsp;重 置&nbsp;" class="reset-button" />
                    </div>
					
                </form>
				</div>
			<%elseif otype="RecordsEditsave" then
			
				rID = Request.Form("rID")
				rType = Request.Form("rType")
				rState = Request.Form("rState")
				rLinkman = Request.Form("rLinkman")
				rNextTime = Request.Form("rNextTime")
				rRemind = Request.Form("rRemind")
				rContent = Request.Form("rContent")
				cType = Request.Form("cType")
				cID = EasyCrm.getNewItem("Records","rID",""&rID&"","cID")
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open "Select Top 1 * From [Records] where rID="&rID,conn,3,2
				rs("rType") = rType
				rs("rState") = rState
				rs("rLinkman") = rLinkman
				if rNextTime <>"" then
				rs("rNextTime") = rNextTime
				end if
				rs("rRemind") = rRemind
				rs("rContent") = rContent
				rs.Update
				rs.Close
				Set rs = Nothing
		
				'同步更新客户类型和插入定时站内信
		
				if ""&rState&"" <> "" and ""&cType&"" <> ""&CRTypeEnd&"" then
				conn.execute ("UPDATE client SET cType='"&rState&"' Where cId ="&EasyCrm.getNewItem("Records","rID",""&rID&"","cID")&" ")
				end if
		
				conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
				'插入操作记录
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Records&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
				Response.Write("<script>location.href='?action=Client&sType=View&otype=Records&cid="&cid&"' ;</script>")
				
			%>
			<%elseif otype="RecordsDel" then
				If Id = "" Then Exit Sub
				Reason = Trim(Request("Reason"))
				cID = EasyCrm.getNewItem("Records","rID",""&ID&"","cID")
				conn.execute("DELETE FROM [Records] where rId = "&Id&" ")
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lReason,lUser,lTime) values('"&cID&"','"&L_Records&"','"&L_insert_action_03&"','"&Reason&"','"&Session("CRM_name")&"','"&now()&"')")
				Response.Write("<script>location.href='?action=Client&sType=View&otype=Records&cid="&cid&"' ;</script>")
			%>
			
			<%elseif otype="Order" then%>
                <div class="content">
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Order] where cId = "&cId&" Order By oId asc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody> 
                        <tr> 
							<th><%=L_Order_oCode%></th> 
							<td><%=rs("oCode")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Order_oLinkman%></th> 
							<td><%=rs("oLinkman")%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Order_oSDate%></th> 
							<td><%=EasyCrm.FormatDate(rs("oSDate"),2)%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Order_oEDate%></th> 
							<td><%=EasyCrm.FormatDate(rs("oEDate"),2)%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Order_oDeposit%></th> 
							<td><%=rs("oDeposit")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Order_oMoney%></th> 
							<td><%=rs("oMoney")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Order_oState%></th> 
							<td><%if rs("oState") = 0 then%>未处理<%elseif rs("oState") = 1 then%>处理中<%elseif rs("oState") = 2 then%>已完成<%elseif rs("oState") = 3 then%>已取消<%end if%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Order_oContent%></th> 
							<td><%=rs("oContent")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Order_oTime%></th> 
							<td><%=EasyCrm.FormatDate(rs("oTime"),2)%></td> 
                        </tr>  
                    </table>
					<%
						rs.MoveNext
						Loop
						else
					%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody> 
                        <tr> 
							<td>暂无数据</td> 
                        </tr> 
                    </table>
						
					<%
							end if
						rs.Close
						Set rs = Nothing
					%>
					<blockquote>仅支持订单查看，不可新增、编辑、删除</blockquote>
				</div>
			
			<%elseif otype="Hetong" then%>
                <div class="content">
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Hetong] where cId = "&cId&" Order By hId asc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody> 
                        <tr> 
							<th><%=L_Hetong_hNum%></th> 
							<td><%=rs("hNum")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Hetong_oId%></th> 
							<td><%=rs("oId")%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Hetong_hSdate%></th> 
							<td><%=EasyCrm.FormatDate(rs("hSdate"),2)%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Hetong_hEdate%></th> 
							<td><%=EasyCrm.FormatDate(rs("hEDate"),2)%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Hetong_hType%></th> 
							<td><%=rs("hType")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Hetong_hMoney%></th> 
							<td><%=rs("hMoney")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Hetong_hRevenue%></th> 
							<td><%=rs("hRevenue")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Hetong_hOwed%></th> 
							<td><%=rs("hOwed")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Hetong_hInvoice%></th> 
							<td><%=rs("hInvoice")%></td> 
                        </tr>
                        <tr> 
							<th><%=L_Hetong_hTax%></th> 
							<td><%=rs("hTax")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Hetong_hState%></th> 
							<td><%=rs("hState")%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Hetong_hContent%></th> 
							<td><%=rs("hContent")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Hetong_hAudit%></th> 
							<td><%=rs("hAudit")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Hetong_hAuditTime%></th> 
							<td><%=rs("hAuditTime")%></td> 
                        </tr>
                        <tr> 
							<th><%=L_Hetong_hAuditReasons%></th> 
							<td><%=rs("hAuditReasons")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Hetong_hTime%></th> 
							<td><%=EasyCrm.FormatDate(rs("hTime"),2)%></td> 
                        </tr>  
                    </table>
					<%
						rs.MoveNext
						Loop
						else
					%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody> 
                        <tr> 
							<td>暂无数据</td> 
                        </tr> 
                    </table>
						
					<%
							end if
						rs.Close
						Set rs = Nothing
					%>
					<blockquote>仅支持合同查看，不可新增、编辑、删除</blockquote>
				</div>
			<%elseif otype="Service" then%>
                <div class="content">
					<input type="button" class="add-button" value=" 新增售后记录 " onclick="location.href='?action=Client&sType=View&otype=ServiceAdd&cid=<%=cid%>'" style="cursor:pointer" /> 
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Service] where cId = "&cId&" Order By sId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody>  
                        <tr> 
							<th><%=L_Service_sSolve%></th> 
							<td><%if rs("sSolve") = 0 then%>未解决<%else%>已解决<%end if%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Service_sTitle%></th> 
							<td><%=rs("sTitle")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Service_sLinkman%></th> 
							<td><%=rs("sLinkman")%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Service_sType%></th> 
							<td><%=rs("sType")%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Service_sSDate%></th> 
							<td><%=EasyCrm.FormatDate(rs("sSDate"),2)%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Service_sContent%></th> 
							<td><%=rs("sContent")%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Service_sEDate%></th> 
							<td><%=EasyCrm.FormatDate(rs("sEDate"),2)%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Service_sInfo%></th> 
							<td><%=rs("sInfo")%></td> 
                        </tr>  
                        <tr> 
							<td colspan=2>
								<input type="button" class="submit-button" value=" 编辑 " onclick="location.href='?action=Client&sType=View&otype=ServiceEdit&cId=<%=rs("cId")%>&id=<%=rs("sid")%>'" style="cursor:pointer" /> 
								<input type="button" class="reset-button" value=" 删除 " onclick="location.href='?action=Client&sType=View&otype=ServiceDel&cId=<%=rs("cId")%>&id=<%=rs("sid")%>'" style="cursor:pointer" /> 
							</td> 
                        </tr>  
                    </table>
					<%
						rs.MoveNext
						Loop
						else
					%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody> 
                        <tr> 
							<td>暂无数据</td> 
                        </tr> 
                    </table>
						
					<%
							end if
						rs.Close
						Set rs = Nothing
					%>
				</div>
			<%elseif otype="ServiceAdd" then%>
	<script language="JavaScript">
	<!-- 必填项提示
	function CheckInput()
	{
		if (<%=Must_Service_ProId%>=="1"){if(document.all.ProId.value == ""){alert("<%=L_Service_ProId & alert04%>");document.all.ProId.focus();return false;}}
		if (<%=Must_Service_sTitle%>=="1"){if(document.all.sTitle.value == ""){alert("<%=L_Service_sTitle & alert04%>");document.all.sTitle.focus();return false;}}
		if (<%=Must_Service_sType%>=="1"){if(document.all.sType.value == ""){alert("<%=L_Service_sType & alert04%>");document.all.sType.focus();return false;}}
		if (<%=Must_Service_sLinkman%>=="1"){if(document.all.sLinkman.value == ""){alert("<%=L_Service_sLinkman & alert04%>");document.all.sLinkman.focus();return false;}}
		if (<%=Must_Service_sSDate%>=="1"){if(document.all.sSDate.value == ""){alert("<%=L_Service_sSDate & alert04%>");document.all.sSDate.focus();return false;}}
		if (<%=Must_Service_sContent%>=="1"){if(document.all.sContent.value == ""){alert("<%=L_Service_sContent & alert04%>");document.all.sContent.focus();return false;}}
	}
	-->
	</script>
	<script>
	function Setdisabled(evt)
	{
		var evt=evt || window.event;   
		var e =evt.srcElement || evt.target;
		 if(e.value=="1")
		 {
			document.all.sInfo.disabled = false; document.all.sInfo.readOnly = false;
			document.all.sEDate.disabled = false; document.all.sEDate.readOnly = false;
			document.all.sEDate.value = "<%=EasyCrm.FormatDate(date(),2)%>";
		 }
		 else
		 {
			document.all.sInfo.disabled = true; document.all.sInfo.readOnly = true;
			document.all.sEDate.disabled = true; document.all.sEDate.readOnly = true;
			document.all.sEDate.classname = "";document.all.sEDate.value = "";
		 }
	}
	</script>
			
                <div class="content">
            	<blockquote>新增售后记录</blockquote>
				<form name="Save" action="?action=Client&sType=View&otype=ServiceAddsave" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sTitle = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sTitle%></label>
						<input name="sTitle" type="text" class="int" id="sTitle" style=" width:60%;">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sType%></label>
						<% = EasyCrm.getSelect("SelectData","Select_Service","sType","") %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sLinkman%></label>
						<% = EasyCrm.getNewSelect("Linkmans","lName","sLinkman"," and cid="&cID&" ","") %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sSDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sSDate%></label>
						<input name="sSDate" type="date" id="sSDate" class="Wdate" size="22" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:00:00'})" value="<%=EasyCrm.FormatDate(date(),2)%>" />
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sContent%></label>
						<textarea name="sContent" id="sContent" class="int" style="width:80%;"></textarea>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sSolve = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sSolve%></label>
						<input name="sSolve" type="radio" id="sSolve" onclick="Setdisabled()" value="0" checked> <%=L_Service_sSolve_0%>
						<input name="sSolve" type="radio" id="sSolve" onclick="Setdisabled()" value="1"> <%=L_Service_sSolve_1%>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sEDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sEDate%></label>
						<input name="sEDate" type="text" id="sEDate" class="Wdate" size="22" value="" disabled readOnly />
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sInfo = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sInfo%></label>
						<textarea name="sInfo" id="sInfo" class="int" style="width:80%;" disabled readOnly></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="cID" type="hidden" value="<%=cID%>">
					<input name="sUser" type="hidden" value="<%=Session("CRM_name")%>">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input type="reset" name="button" id="button2" value="&nbsp;重 置&nbsp;" class="reset-button" />
                    </div>
					
                </form>
				</div>
			<%elseif otype="ServiceAddsave" then
				cID = Request.Form("cID")
				ProId = Request.Form("ProId")
				sTitle = Request.Form("sTitle")
				sType = Request.Form("sType")
				sLinkman = Request.Form("sLinkman")
				sSDate = Request.Form("sSDate")
				sEDate = Request.Form("sEDate")
				sContent = Request.Form("sContent")
				sSolve = Request.Form("sSolve")
				sInfo = Request.Form("sInfo")
				sUser = Request.Form("sUser")
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open "Select Top 1 * From [Service]",conn,3,2
				rs.AddNew
				rs("cID") = cID
				if ProId <> "" then
				rs("ProId") = ProId
				end if
				rs("sTitle") = sTitle
				rs("sType") = sType
				rs("sLinkman") = sLinkman
				if sSDate<>"" then
				rs("sSDate") = sSDate
				end if
				if sEDate<>"" then
				rs("sEDate") = sEDate
				end if
				rs("sContent") = sContent
				rs("sSolve") = sSolve
				rs("sInfo") = sInfo
				rs("sUser") = sUser
				rs("sTime") = now()
				rs.Update
				rs.Close
				Set rs = Nothing

				conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
				'插入操作记录
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Service&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
				
				Response.Write("<script>location.href='?action=Client&sType=View&otype=Service&cid="&cid&"' ;</script>")
			
			%>
			<%elseif otype="ServiceEdit" then%>
			
	<script>
	function Setdisabled(evt)
	{
		var evt=evt || window.event;   
		var e =evt.srcElement || evt.target;
		 if(e.value=="1")
		 {
			document.all.sInfo.disabled = false; document.all.sInfo.readOnly = false;
			document.all.sEDate.disabled = false; document.all.sEDate.readOnly = false;
			document.all.sEDate.value = "<%=EasyCrm.FormatDate(date(),2)%>";
		 }
		 else
		 {
			document.all.sInfo.disabled = true; document.all.sInfo.readOnly = true;
			document.all.sEDate.disabled = true; document.all.sEDate.readOnly = true;
			document.all.sEDate.classname = "";document.all.sEDate.value = "";
		 }
	}
	</script>
                <div class="content">
            	<blockquote>修改售后记录</blockquote>
				<form name="Save" action="?action=Client&sType=View&otype=ServiceEditsave" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sTitle = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sTitle%></label>
						<input name="sTitle" type="text" class="int" id="sTitle" style=" width:60%;" value="<%=EasyCrm.getNewItem("Service","sID",""&ID&"","sTitle")%>">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sType%></label>
						<% = EasyCrm.getSelect("SelectData","Select_Service","sType",EasyCrm.getNewItem("Service","sID",""&ID&"","sType")) %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sLinkman = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sLinkman%></label>
						<% = EasyCrm.getNewSelect("Linkmans","lName","sLinkman"," and cid="&EasyCrm.getNewItem("Service","sID",""&ID&"","cID")&" ",EasyCrm.getNewItem("Service","sID",""&ID&"","sLinkman")) %>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sSDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sSDate%></label>
						<input name="sSDate" type="date" id="sSDate" class="Wdate" size="22" onFocus="WdatePicker({dateFmt:'yyyy-MM-dd HH:00:00'})" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Service","sID",""&ID&"","sSDate"),2)%>" />
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sContent%></label>
						<textarea name="sContent" id="sContent" class="int" style="width:80%;"><%=EasyCrm.getNewItem("Service","sID",""&ID&"","sContent")%></textarea>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sSolve = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sSolve%></label>
						<input name="sSolve" type="radio" id="sSolve" onclick="Setdisabled()" value="0" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=0 then%>checked <%end if%>> <%=L_Service_sSolve_0%>
						<input name="sSolve" type="radio" id="sSolve" onclick="Setdisabled()" value="1" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=1 then%>checked <%end if%>> <%=L_Service_sSolve_1%>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sEDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sEDate%></label>
						<input name="sEDate" type="text" id="sEDate" class="Wdate" size="22" value="<%=EasyCrm.FormatDate(EasyCrm.getNewItem("Service","sID",""&ID&"","sEDate"),2)%>" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=0 then%>disabled readOnly<%end if%> />
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Service_sInfo = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Service_sInfo%></label>
						<textarea name="sInfo" id="sInfo" class="int" style="width:80%;" <%if EasyCrm.getNewItem("Service","sID",""&ID&"","sSolve")=0 then%>disabled readOnly<%end if%> ><%=EasyCrm.getNewItem("Service","sID",""&ID&"","sInfo")%></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="sID" type="hidden" value="<%=ID%>">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input type="reset" name="button" id="button2" value="&nbsp;重 置&nbsp;" class="reset-button" />
                    </div>
					
                </form>
				</div>
			<%elseif otype="ServiceEditsave" then
			
				sID = Request.Form("sID")
				ProId = Request.Form("ProId")
				sTitle = Request.Form("sTitle")
				sType = Request.Form("sType")
				sLinkman = Request.Form("sLinkman")
				sSDate = Request.Form("sSDate")
				sEDate = Request.Form("sEDate")
				sSolve = Request.Form("sSolve")
				sContent = Request.Form("sContent")
				sInfo = Request.Form("sInfo")
				cId = EasyCrm.getNewItem("Service","sID",""&sID&"","cId")

				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open "Select Top 1 * From [Service] where sID="&sID,conn,3,2
				if ProId <> "" then
				rs("ProId") = ProId
				end if
				rs("sTitle") = sTitle
				rs("sType") = sType
				rs("sLinkman") = sLinkman
				rs("sSDate") = sSDate
				if sEDate <> "" then
				rs("sEDate") = sEDate
				end if
				rs("sSolve") = sSolve
				rs("sContent") = sContent
				rs("sInfo") = sInfo
				rs.Update
				rs.Close
				Set rs = Nothing
		
				conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
				'插入操作记录
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Service&"','"&L_insert_action_02&"','"&Session("CRM_name")&"','"&now()&"')")	
				Response.Write("<script>location.href='?action=Client&sType=View&otype=Service&cid="&cid&"' ;</script>")
				
			%>
			<%elseif otype="ServiceDel" then
				If Id = "" Then Exit Sub
				Reason = Trim(Request("Reason"))
				cId = EasyCrm.getNewItem("Service","sID",""&ID&"","cId")
				conn.execute("DELETE FROM [Service] where sID = "&Id&" ")
				
				'插入操作记录
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lReason,lUser,lTime) values('"&cID&"','"&L_Service&"','"&L_insert_action_03&"','"&Reason&"','"&Session("CRM_name")&"','"&now()&"')")
				Response.Write("<script>location.href='?action=Client&sType=View&otype=Service&cid="&cid&"' ;</script>")
			%>
			
			<%elseif otype="Expense" then%>
                <div class="content">
					<input type="button" class="add-button" value=" 新增收入 " onclick="location.href='?action=Client&sType=View&otype=ExpenseAdd&eOutIn=1&cid=<%=cid%>'" style="cursor:pointer" /> 
					<input type="button" class="submit-button" value=" 新增支出 " onclick="location.href='?action=Client&sType=View&otype=ExpenseAdd&eOutIn=0&cid=<%=cid%>'" style="cursor:pointer" /> 
						<%
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.Open "Select * From [Expense] where cId = "&cId&" Order By eId desc ",conn,1,1
						If rs.RecordCount > 0 Then
						Do While Not rs.BOF And Not rs.EOF
						%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody>  
                        <tr> 
							<th><%=L_Expense_eDate%></th> 
							<td><%=EasyCrm.FormatDate(rs("eDate"),2)%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Expense_eOutIn%></th> 
							<td><%if rs("eOutIn") = 1 then %>收入<%else%>支出<%end if%></td> 
                        </tr> 
                        <tr> 
							<th><%=L_Expense_eType%></th> 
							<td><%=rs("eType")%></td> 
                        </tr>  
                        <tr> 
							<th><%=L_Expense_eMoney%></th> 
							<td><%=rs("eMoney")%> 元</td> 
                        </tr> 
                        <tr> 
							<th><%=L_Expense_eContent%></th> 
							<td><%=rs("eContent")%></td> 
                        </tr> 
                    </table>
					<%
						rs.MoveNext
						Loop
						else
					%>
					<table class="tabledata" style="margin-top:10px;"> 
						<col width="80">
						<tbody> 
                        <tr> 
							<td>暂无数据</td> 
                        </tr> 
                    </table>
						
					<%
							end if
						rs.Close
						Set rs = Nothing
					%>
				</div>
			<%elseif otype="ExpenseAdd" then
			eOutIn = Trim(Request("eOutIn"))
			%>
	<script language="JavaScript">
	<!-- 必填项提示
	function CheckInput()
	{
		if (<%=Must_Expense_eDate%>=="1"){if(document.all.eDate.value == ""){alert("<%=L_Expense_eDate & alert04%>");document.all.eDate.focus();return false;}}
		if (<%=Must_Expense_eType%>=="1"){if(document.all.eType.value == ""){alert("<%=L_Expense_eType & alert04%>");document.all.eType.focus();return false;}}
		if (<%=Must_Expense_eMoney%>=="1"){if(document.all.eMoney.value == ""){alert("<%=L_Expense_eMoney & alert04%>");document.all.eMoney.focus();return false;}}
		if (<%=Must_Expense_eContent%>=="1"){if(document.all.eContent.value == ""){alert("<%=L_Expense_eContent & alert04%>");document.all.eContent.focus();return false;}}
	}
	-->
	</script>
                <div class="content">
            	<blockquote>新增费用<%if eOutIn = 1 then %>收入<%else%>支出<%end if%>记录</blockquote>
				<form name="Save" action="?action=Client&sType=View&otype=ExpenseAddsave" method="post" onSubmit="return CheckInput();">
                    <div class="form-line">
						<label class="st-label"><%if Must_Expense_eDate = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eDate%></label>
						<input name="eDate" type="date" id="eDate" class="Wdate" size="22" value="<%=EasyCrm.FormatDate(date(),2)%>" />
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Expense_eType = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eType%></label>
						<%if eOutIn = 1 then%>
							<% = EasyCrm.getSelect("SelectData","Select_ExpenseIN","eType","") %>
						<%else%>
							<% = EasyCrm.getSelect("SelectData","Select_ExpenseOUT","eType","") %>
						<%end if%>
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Expense_eMoney = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eMoney%></label>
						<input name="eMoney" type="number" class="int" id="eMoney" style=" width:10%;">
                    </div>
					
                    <div class="form-line">
						<label class="st-label"><%if Must_Expense_eContent = 1 then %><font color="#FF0000">*</font> <%end if%> <%=L_Expense_eContent%></label>
						<textarea name="eContent" id="eContent" class="int" style="width:80%;"></textarea>
                    </div>
                    
                    <div class="form-line">
					<input name="cID" type="hidden" value="<%=cID%>">
					<input name="eOutIn" type="hidden" value="<%=eOutIn%>" />
					<input name="eUser" type="hidden" value="<%=Session("CRM_name")%>">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input type="reset" name="button" id="button2" value="&nbsp;重 置&nbsp;" class="reset-button" />
                    </div>
					
                </form>
				</div>
			<%elseif otype="ExpenseAddsave" then
				cID = Request.Form("cID")
				eDate = Request.Form("eDate")
				eOutIn = Request.Form("eOutIn")
				eType = Request.Form("eType")
				eMoney = Request.Form("eMoney")
				eContent = Request.Form("eContent")
				eUser = Request.Form("eUser")
				
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open "Select Top 1 * From [Expense]",conn,3,2
				rs.AddNew
				rs("cID") = cID
				if eDate<>"" then
				rs("eDate") = eDate
				end if
				rs("eOutIn") = eOutIn
				rs("eType") = eType
				rs("eMoney") = eMoney
				rs("eContent") = eContent
				rs("eUser") = eUser
				rs("eTime") = now()
				rs.Update
				rs.Close
				Set rs = Nothing
				
				conn.execute("update [Client] set cLastUpdated = '"&Now()&"' where cId = "&cId&" ")
				'插入操作记录
				conn.execute ("insert into Logfile(lCid,lClass,lAction,lUser,lTime) values('"&cid&"','"&L_Expense&"','"&L_insert_action_01&"','"&Session("CRM_name")&"','"&now()&"')")	
				
				Response.Write("<script>location.href='?action=Client&sType=View&otype=Expense&cid="&cid&"' ;</script>")
			
			%>
			
			<%elseif otype="Share" then%>
			
	<script>
	function Setdisabled(evt)
	{
		var evt=evt || window.event;   
		var e =evt.srcElement || evt.target;
		
		 if(e.value=="1")
		 {
			var a = document.all.cShareRange; 
			for (var i=0; i<a.length; i++)   
			{ 
				a[i].disabled=false; 
				a[i].readOnly=false; 
			} 
		 }
		 else
		 {
			var a = document.all.cShareRange; 
			for (var i=0; i<a.length; i++)   
			{ 
				a[i].disabled=true; 
				a[i].readOnly=true; 
			} 
		 }
	}
	</script>
                <div class="content">
            	<blockquote>选择共享对象</blockquote>
				<form name="Save" action="?action=Client&sType=View&otype=Sharesave" method="post" onSubmit="return CheckInput();">
                    <div class="form-line" style="color:red;">
						<label class="st-label">是否共享</label>
						<input type="radio" id="cShare" name="cShare" value= '0' <%if EasyCrm.getNewItem("Client","cID",""&cID&"","cShare")=0 then %>checked <%end if%> onclick="Setdisabled()"> 否　<input type="radio" id="cShare" name="cShare" value= '1' <%if EasyCrm.getNewItem("Client","cID",""&cID&"","cShare")=1 then %>checked <%end if%> onclick="Setdisabled()"> 是
                    </div>
					
							<%
								Set rsg = Server.CreateObject("ADODB.Recordset")
								rsg.Open "Select * From [system_group]",conn,1,1
								Do While Not rsg.BOF And Not rsg.EOF
							%>
                    <div class="form-line">
						<label class="st-label"><%=rsg("gName")%></label>
								<%
									Set rsm = Server.CreateObject("ADODB.Recordset")
									rsm.Open "Select * From [user] where uGroup="&rsg("gId")&" ",conn,1,1
									Do While Not rsm.BOF And Not rsm.EOF
								%>
									<input type="checkbox" id="cShareRange" name="cShareRange" value= '<%=rsm("uName")%>' <%if EasyCrm.getNewItem("Client","cID",""&cID&"","cShare")=0 then %>disabled readOnly <%end if%> <%if inStr(EasyCrm.getNewItem("Client","cID",""&cID&"","cShareRange"),rsm("uName"))>0 then%>checked<%end if%>> <%=rsm("uName")%>　
								<%
									rsm.MoveNext
									Loop
									rsm.Close
									Set rsm = Nothing
								%>
                    </div>
					 
							<%
								rsg.MoveNext
								Loop
								rsg.Close
								Set rsg = Nothing
							%>
                    
                    <div class="form-line">
					<input name="cID" type="hidden" value="<%=cID%>">
					<input name="sUser" type="hidden" value="<%=Session("CRM_name")%>">
                    <input type="submit" name="button" id="button" value="&nbsp;提 交&nbsp;" class="submit-button" />
                    <input type="reset" name="button" id="button2" value="&nbsp;重 置&nbsp;" class="reset-button" />
                    </div>
					
                </form>
				</div>
	
			<%elseif otype="Sharesave" then
			
			cShare = Request.Form("cShare")
			cShareRange = Request.Form("cShareRange")
			conn.execute("update [Client] set cShare='"&cShare&"',cShareRange='"&cShareRange&"' where cId = "&cID&" ")
			Response.Write("<script>location.href='?action=Client&sType=View&otype=Share&cid="&cid&"' ;</script>")
			
			%>
			<%elseif otype="History" then%>
					<table class="tabledata"> 
						<tbody> 
						<THEAD> 
                        <tr> 
							<td >编号</td> 
							<td>数据表</td> 
							<td>行为</td> 
							<td>操作人</td> 
							<td>时间</td> 
                        </tr> 
						</THEAD> 
						<%
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.Open "Select * From [Logfile] where lCid = "&cId&" Order By lId desc ",conn,1,1
							If rs.RecordCount > 0 Then
							Do While Not rs.BOF And Not rs.EOF
						%>
								<tr>
									<td>[<%=rs("lId")%>]</td>
									<td><%=rs("lClass")%></td>
									<td><%=rs("lAction")%></td>
									<td><%=rs("lUser")%></td>
									<td><%=rs("lTime")%></td>
								</tr>
						<%
							rs.MoveNext
							Loop
							end if
							rs.Close
							Set rs = Nothing
						%>
                    </table> 
			<%end if%>
			</div>

<%
end if
%>
            
			<%=Footer%>
            
    <div class="clear"></div>
    </div>
    <!-- end page -->
</body>
</html>
<script type="text/javascript" src="js/frame.js"></script>
<% End Sub %>