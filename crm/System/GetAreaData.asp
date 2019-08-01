<!--#include file="../Data/Conn.asp"--><!--#include file="../data/EasyCrm.asp"-->
<%Set EasyCrm  = New EasyCRM_CRM%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/Common.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/jquery.min.js"></script>
<script language="javascript" src="<%=SiteUrl&skinurl%>Js/tips.js"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/jquery.artDialog.js?skin=default"></script>
<script src="<%=SiteUrl&skinurl%>aridialog/iframeTools.js"></script>
</head>
<body>
<style>body{padding-bottom:55px;}</style>
<%
action = Trim(Request("action"))
sType = Trim(Request("sType"))
tipinfo = Trim(Request("tipinfo"))

Select Case action
Case "Setting"
    Call Setting()
Case "Products"
    Call Products()
Case "AreaData"
    Call AreaData()
Case "CustomField"
    Call CustomField()
Case "SelectData"
    Call SelectData()
Case "User"
    Call User()
Case "Group"
    Call Group()
Case "Level"
    Call Level()
Case "InfoList"
    Call InfoList()
End Select
 
Sub AreaData() '地区数据更新

if tipinfo<>"" then
	Response.Write("<script>art.dialog({title: '提示',time: 1,icon: 'warning',content: '"&tipinfo&"'});</script>")
end if

if sType="Import" then '导入全国地区
%>
		<form name="Save" action="GetAreaData.asp?action=AreaData&sType=ImportAdd" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="60" /><col /><col width="60" /><col /><col width="60" /><col /><col width="60" /><col /><col width="60" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="10"><B>选择要导入的省份（不可重复导入）</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">北京市</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport1" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">天津市</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport2" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">上海市</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport3" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">重庆市</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport4" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">河北</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport5" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">山西</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport6" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">内蒙古</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport7" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">辽宁</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport8" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">吉林</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport9" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'吉林' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">黑龙江</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport10" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">江苏</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport11" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">浙江</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport12" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">安徽</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport13" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">福建</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport14" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'福建' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">江西</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport15" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">山东</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport16" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">河南</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport17" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">湖北</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport18" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">湖南</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport19" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">广东</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport20" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">甘肃</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport21" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">四川</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport22" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">贵州</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport23" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'贵州' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">海南</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport24" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">云南</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport25" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">青海</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport26" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'青海' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">陕西</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport27" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">广西</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport28" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">西藏</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport29" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'西藏' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">宁夏</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport30" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'宁夏' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">新疆</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport31" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">台湾</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport32" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">澳门</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport33" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'澳门' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">香港</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport34" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'香港' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">海外</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport35" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
						</table>
					</td>
				</tr>
				<script language=javascript> 
				function selectall(id){ //用id区分  
				var tform=document.forms['Save'];  
				for(var i=0;i<tform.length;i++){  
				var e=tform.elements[i];  
				if(e.type=="checkbox" && e.id==id) e.checked=!e.checked;  } } 
				</script>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="button" name="checkall" class="button42" onclick="javascript:selectall('AreaImport')" value="全选/反选">　
							<input type="submit" name="Submit" class="button45" value="立即导入">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="ImportAdd" then '开始导入

	AreaImport = ""
	for i = 1 to 35
		if Request("AreaImport" & i) = "1" then
			AreaImport = AreaImport & "1"
		else
			AreaImport = AreaImport & "0"
		end if
	next
	
	if mid(AreaImport, 1, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'北京市')") 
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'东城')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'西城')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'崇文')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'宣武')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'朝阳')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'丰台')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'石景山')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'海淀')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'门头沟')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'房山')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'通州')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'顺义')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'昌平')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'大兴')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'平谷')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'怀柔')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'密云')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'北京市' and aFid='0'","aId")&",'延庆')")
		Response.Write("<script>art.dialog({title: '提示',time: 0.5,icon: 'warning',content: '北京市地区数据导入成功'});</script>")
	end if
	
	if mid(AreaImport, 2, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'天津市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'和平')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'东丽')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'河东')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'西青')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'河西')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'津南')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'南开')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'北辰')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'河北')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'武清')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'红挢')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'塘沽')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'汉沽')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'大港')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'宁河')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'静海')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'宝坻')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'天津市' and aFid='0'","aId")&",'蓟县')")
	end if
	
	if mid(AreaImport, 3, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'上海市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'崇明')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'黄浦')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'卢湾')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'徐汇')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'长宁')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'静安')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'普陀')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'闸北')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'虹口')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'杨浦')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'闵行')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'宝山')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'嘉定')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'浦东')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'金山')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'松江')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'青浦')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'南汇')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'上海市' and aFid='0'","aId")&",'奉贤')")
	end if
	
	if mid(AreaImport, 4, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'重庆市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'万州')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'涪陵')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'渝中')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'大渡口')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'江北')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'沙坪坝')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'九龙坡')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'南岸')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'北碚')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'万盛')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'双挢')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'渝北')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'巴南')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'黔江')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'长寿')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'綦江')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'潼南')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'铜梁')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'大足')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'荣昌')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'壁山')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'梁平')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'城口')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'丰都')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'垫江')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'武隆')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'忠县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'开县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'云阳')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'奉节')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'巫山')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'巫溪')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'石柱')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'秀山')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'酉阳')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'彭水')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'江津')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'合川')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'永川')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'重庆市' and aFid='0'","aId")&",'南川')")
	end if
	
	if mid(AreaImport, 5, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'河北')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'石家庄市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'唐山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'秦皇岛市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'邯郸市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'邢台市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'保定市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'张家口市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'承德市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'沧州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'廊坊市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河北' and aFid='0'","aId")&",'衡水市')")
	end if
	
	if mid(AreaImport, 6, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'山西')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'太原市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'大同市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'阳泉市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'长治市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'晋城市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'朔州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'晋中市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'运城市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'忻州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'临汾市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山西' and aFid='0'","aId")&",'吕梁市')")
	end if
	
	if mid(AreaImport, 7, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'内蒙古')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'呼和浩特市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'包头市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'乌海市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'赤峰市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'通辽市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'鄂尔多斯市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'呼伦贝尔市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'巴彦淖尔市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'乌兰察布市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'锡林郭勒盟')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'兴安盟')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'内蒙古' and aFid='0'","aId")&",'阿拉善盟')")
	end if
	
	if mid(AreaImport, 8, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'辽宁')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'沈阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'大连市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'鞍山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'抚顺市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'本溪市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'丹东市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'锦州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'营口市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'阜新市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'辽阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'盘锦市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'铁岭市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'朝阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'辽宁' and aFid='0'","aId")&",'葫芦岛市')")
	end if
	
	if mid(AreaImport, 9, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'吉林')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'吉林' and aFid='0'","aId")&",'长春市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'吉林' and aFid='0'","aId")&",'吉林市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'吉林' and aFid='0'","aId")&",'四平市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'吉林' and aFid='0'","aId")&",'辽源市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'吉林' and aFid='0'","aId")&",'通化市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'吉林' and aFid='0'","aId")&",'白山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'吉林' and aFid='0'","aId")&",'松原市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'吉林' and aFid='0'","aId")&",'白城市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'吉林' and aFid='0'","aId")&",'延边')")
	end if
	
	if mid(AreaImport, 10, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'黑龙江')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'哈尔滨市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'齐齐哈尔市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'鹤岗市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'双鸭山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'鸡西市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'大庆市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'伊春市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'牡丹江市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'佳木斯市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'七台河市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'黑河市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'绥化市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'黑龙江' and aFid='0'","aId")&",'大兴安岭')")
	end if
	
	if mid(AreaImport, 11, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'江苏')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'南京市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'无锡市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'徐州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'常州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'苏州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'南通市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'连云港市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'淮安市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'盐城市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'扬州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'镇江市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'泰州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江苏' and aFid='0'","aId")&",'宿迁市')")
	end if
	
	if mid(AreaImport, 12, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'浙江')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'杭州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'宁波市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'温州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'嘉兴市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'湖州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'绍兴市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'金华市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'衢州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'舟山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'台州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'浙江' and aFid='0'","aId")&",'丽水市')")
	end if
	
	if mid(AreaImport, 13, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'安徽')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'合肥市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'芜湖市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'蚌埠市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'淮南市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'马鞍山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'淮北市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'铜陵市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'安庆市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'黄山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'滁州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'阜阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'宿州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'巢湖市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'六安市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'亳州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'池州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'安徽' and aFid='0'","aId")&",'宣城市')")
	end if
	
	if mid(AreaImport, 14, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'福建')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'福建' and aFid='0'","aId")&",'福州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'福建' and aFid='0'","aId")&",'厦门市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'福建' and aFid='0'","aId")&",'莆田市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'福建' and aFid='0'","aId")&",'三明市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'福建' and aFid='0'","aId")&",'泉州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'福建' and aFid='0'","aId")&",'漳州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'福建' and aFid='0'","aId")&",'南平市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'福建' and aFid='0'","aId")&",'龙岩市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'福建' and aFid='0'","aId")&",'宁德市')")
	end if
	
	if mid(AreaImport, 15, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'江西')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'南昌市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'景德镇市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'萍乡市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'九江市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'新余市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'鹰潭市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'赣州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'吉安市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'宜春市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'抚州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'江西' and aFid='0'","aId")&",'上饶市')")
	end if
	
	if mid(AreaImport, 16, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'山东')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'济南市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'青岛市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'淄博市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'枣庄市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'东营市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'烟台市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'潍坊市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'济宁市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'泰安市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'威海市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'日照市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'莱芜市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'临沂市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'德州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'聊城市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'滨州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'山东' and aFid='0'","aId")&",'菏泽市')")
	end if
	
	if mid(AreaImport, 17, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'河南')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'郑州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'开封市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'洛阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'平顶山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'安阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'鹤壁市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'新乡市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'焦作市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'濮阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'许昌市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'漯河市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'三门峡市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'南阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'商丘市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'信阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'周口市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'驻马店市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'河南' and aFid='0'","aId")&",'济源市')")
	end if
	
	if mid(AreaImport, 18, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'湖北')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'武汉市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'黄石市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'十堰市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'荆州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'宜昌市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'襄樊市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'鄂州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'荆门市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'孝感市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'黄冈市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'咸宁市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'随州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'仙桃市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'天门市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'潜江市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'神农架')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖北' and aFid='0'","aId")&",'恩施')")
	end if
	
	if mid(AreaImport, 19, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'湖南')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'长沙市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'株洲市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'湘潭市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'衡阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'邵阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'岳阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'常德市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'张家界市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'益阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'郴州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'永州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'怀化市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'娄底市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'湖南' and aFid='0'","aId")&",'湘西')")
	end if
	
	if mid(AreaImport, 20, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'广东')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'广州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'深圳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'珠海市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'汕头市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'韶关市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'佛山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'江门市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'湛江市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'茂名市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'肇庆市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'惠州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'梅州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'汕尾市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'河源市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'阳江市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'清远市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'东莞市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'中山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'潮州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'揭阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广东' and aFid='0'","aId")&",'云浮市')")
	end if
	
	if mid(AreaImport, 21, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'甘肃')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'兰州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'金昌市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'白银市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'天水市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'嘉峪关市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'武威市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'张掖市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'平凉市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'酒泉市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'庆阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'定西市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'陇南市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'临夏')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'甘肃' and aFid='0'","aId")&",'甘南')")
	end if
	
	if mid(AreaImport, 22, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'四川')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'成都市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'自贡市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'攀枝花市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'泸州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'德阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'绵阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'广元市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'遂宁市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'内江市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'乐山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'南充市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'眉山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'宜宾市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'广安市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'达州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'雅安市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'巴中市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'资阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'阿坝')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'甘孜')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'四川' and aFid='0'","aId")&",'凉山')")
	end if
	
	if mid(AreaImport, 23, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'贵州')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'贵州' and aFid='0'","aId")&",'贵阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'贵州' and aFid='0'","aId")&",'六盘水市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'贵州' and aFid='0'","aId")&",'遵义市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'贵州' and aFid='0'","aId")&",'安顺市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'贵州' and aFid='0'","aId")&",'铜仁地区')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'贵州' and aFid='0'","aId")&",'毕节地区')")
	end if
	
	if mid(AreaImport, 24, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'海南')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'海口市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'三亚市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'五指山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'琼海市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'儋州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'文昌市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'万宁市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'东方市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'澄迈县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'定安县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'屯昌县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'临高县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'白沙')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'昌江')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'乐东')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'陵水')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'保亭')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海南' and aFid='0'","aId")&",'琼中')")
	end if
	
	if mid(AreaImport, 25, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'云南')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'昆明市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'曲靖市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'玉溪市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'保山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'昭通市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'丽江市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'思茅市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'临沧市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'文山')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'红河')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'西双版纳')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'楚雄')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'大理')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'德宏')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'怒江')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'云南' and aFid='0'","aId")&",'迪庆')")
	end if
	
	if mid(AreaImport, 26, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'青海')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'青海' and aFid='0'","aId")&",'西宁市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'青海' and aFid='0'","aId")&",'海东地区')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'青海' and aFid='0'","aId")&",'海北')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'青海' and aFid='0'","aId")&",'黄南')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'青海' and aFid='0'","aId")&",'海南')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'青海' and aFid='0'","aId")&",'果洛')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'青海' and aFid='0'","aId")&",'玉树')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'青海' and aFid='0'","aId")&",'海西')")
	end if
	
	if mid(AreaImport, 27, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'陕西')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId")&",'西安市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId")&",'铜川市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId")&",'宝鸡市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId")&",'咸阳市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId")&",'渭南市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId")&",'延安市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId")&",'汉中市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId")&",'榆林市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId")&",'安康市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'陕西' and aFid='0'","aId")&",'商洛市')")
	end if
	
	if mid(AreaImport, 28, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'广西')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'南宁市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'柳州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'桂林市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'梧州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'北海市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'防城港市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'钦州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'贵港市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'玉林市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'百色市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'贺州市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'河池市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'来宾市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'广西' and aFid='0'","aId")&",'崇左市')")
	end if
	
	if mid(AreaImport, 29, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'西藏')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'西藏' and aFid='0'","aId")&",'拉萨市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'西藏' and aFid='0'","aId")&",'那曲地区')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'西藏' and aFid='0'","aId")&",'昌都地区')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'西藏' and aFid='0'","aId")&",'山南地区')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'西藏' and aFid='0'","aId")&",'日喀则地区')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'西藏' and aFid='0'","aId")&",'阿里地区')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'西藏' and aFid='0'","aId")&",'林芝地区')")
	end if
	
	if mid(AreaImport, 30, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'宁夏')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'宁夏' and aFid='0'","aId")&",'银川市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'宁夏' and aFid='0'","aId")&",'石嘴山市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'宁夏' and aFid='0'","aId")&",'吴忠市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'宁夏' and aFid='0'","aId")&",'固原市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'宁夏' and aFid='0'","aId")&",'中卫市')")
	end if
	
	if mid(AreaImport, 31, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'新疆')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'乌鲁木齐市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'克拉玛依市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'石河子市　')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'阿拉尔市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'图木舒克市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'五家渠市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'吐鲁番市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'阿克苏市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'喀什市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'哈密市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'和田市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'阿图什市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'库尔勒市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'昌吉市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'阜康市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'米泉市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'博乐市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'伊宁市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'奎屯市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'塔城市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'乌苏市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'新疆' and aFid='0'","aId")&",'阿勒泰市')")
	end if
	
	if mid(AreaImport, 32, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'台湾')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'台北市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'高雄市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'基隆市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'台中市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'台南市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'新竹市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'嘉义市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'台北县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'宜兰县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'桃园县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'新竹县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'苗栗县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'台中县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'彰化县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'南投县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'云林县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'嘉义县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'台南县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'高雄县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'屏东县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'澎湖县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'台东县')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'台湾' and aFid='0'","aId")&",'花莲县')")
	end if
	
	if mid(AreaImport, 33, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'澳门')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'澳门' and aFid='0'","aId")&",'澳门市')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'澳门' and aFid='0'","aId")&",'海岛市')")
	end if
	
	if mid(AreaImport, 34, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'香港')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'香港' and aFid='0'","aId")&",'香港岛')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'香港' and aFid='0'","aId")&",'九龙')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'香港' and aFid='0'","aId")&",'新界')")
	end if
	
	if mid(AreaImport, 35, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'海外')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'美国')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'澳大利亚')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'巴西')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'英国')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'加拿大')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'埃及')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'法国')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'德国')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'印度')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'爱尔兰')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'以色列')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'意大利')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'日本')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'荷兰')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'新西兰')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'葡萄牙')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'俄国')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'西班牙')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'瑞典')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'海外' and aFid='0'","aId")&",'瑞士')")
	end if
	
	
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
%>
<%
elseif sType="BigClassAdd" then '添加大类
%>
		<form name="Save" action="GetAreaData.asp?action=AreaData&sType=SaveBigClassAdd" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">类别名称</td>
								<td class="td_l_l"><input name="aName" type="text" id="aName" class="int" size="40" /></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveBigClassAdd" then
		aName = Request.Form("aName")
		If aName = "" Then
			Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=BigClassAdd&tipinfo=地区大类名不能为空';</script>")
			Exit Sub
		End If
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [AreaData] Where aName = '"&aName&"' ",conn,3,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=BigClassAdd&tipinfo=已存在！';</script>")
		Response.End()
		End If
		rs.Close

    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From [AreaData] ",conn,3,2
		rs.AddNew
		rs("aFId") = 0
		rs("aName") = aName
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

elseif sType="BigClassEdit" then '修改大类
	aId = Request("aId")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [AreaData] Where aId = " & aId,conn,1,1
	aName = rs("aName")
	rs.Close
	Set rs = Nothing
%>
		<form name="Save" action="GetAreaData.asp?action=AreaData&sType=SaveBigClassEdit" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">类别名称</td>
								<td class="td_l_l"><input name="aName" type="text" id="aName" class="int" size="20" value="<%=aName%>" /></td>
								<input name="aId" type="hidden" id="aId" value="<% = aId %>">
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveBigClassEdit" then
		aId = Request.Form("aId")
		aName = Request.Form("aName")
		If aName = "" Then
			Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=BigClassEdit&aId="&aId&"&tipinfo=地区大类不能为空';</script>")
			Exit Sub
		End If
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [AreaData] Where aName = '"&aName&"' And aId <> " & aId,conn,3,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=BigClassEdit&aId="&aId&"&tipinfo=已存在！';</script>")
		Response.End()
		End If
		rs.Close

    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select * From [AreaData] where aId="&aId&" ",conn,3,2
		rs("aFId") = 0
		rs("aName") = aName
    	rs.Update
    	rs.Close
    	Set rs = Nothing
		
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

elseif sType="SmallClassAdd" then '添加小类
		aFid = Request("aFid")
%>
		<form name="Save" action="GetAreaData.asp?action=AreaData&sType=SaveSmallClassAdd" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title">上级分类</td>
								<td class="td_r_l">
									<select name="aFid" class="int">
										<option value="">请选择</option>
										<% 
											Set rsb = Conn.Execute("select * from [AreaData] where aFid = '0' ")
											If Not rsb.Eof then
											Do While Not rsb.Eof
											aId= rsb("aId")
											aName= rsb("aName")
										%>
										<option value="<%=aId%>" <%if ""&aId&"" = ""&aFid&"" then%>selected<%end if%>><%=aName%></option>
										<%
											rsb.Movenext
											Loop
											End If
											rsb.Close
											Set rsb = Nothing 
										%>
									</select> 
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">类别名称</td>
								<td class="td_l_l"><input name="aName" type="text" id="aName" class="int" size="40" /></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveSmallClassAdd" then
	aFId = Trim(Request.Form("aFId"))
	aName = Trim(Request.Form("aName"))
	If aFId = "" Then
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassAdd&aFId="&aFId&"&tipinfo=地区大类不能为空';</script>")
		Exit Sub
	End If
	If aName = "" Then
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassAdd&aFId="&aFId&"&tipinfo=地区小类不能为空';</script>")
		Exit Sub
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [AreaData] Where aFId='"&aFId&"' and aName = '" & aName & "'",conn,3,2
	If rs.RecordCount > 0 Then
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassAdd&aFId="&aFId&"&tipinfo=已存在！' ;</script>")
		rs.Close
		Set rs = Nothing
		Exit Sub
	Else
	    rs.AddNew
		rs("aFId") = aFId
		rs("aName") = aName
		rs.Update
		rs.Close
		Set rs = Nothing
	End If
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")

elseif sType="SmallClassEdit" then '编辑小类
		aId = Request("aId")
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [AreaData] Where aId = " & aId,conn,1,1
		aFId = rs("aFId")
		aName = rs("aName")
%>
		<form name="Save" action="GetAreaData.asp?action=AreaData&sType=SaveSmallClassEdit" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>详细内容</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title">上级分类</td>
								<td class="td_r_l">
									<select name="aFId" class="int">
										<option value="">请选择</option>
										<% 
											Set rsb = Conn.Execute("select * from [AreaData] where aFId = '0' ")
											If Not rsb.Eof then
											Do While Not rsb.Eof
										%>
										<option value="<%=rsb("aId")%>" <%if ""&aFId&"" = ""&rsb("aId")&"" then%>selected<%end if%>><%=rsb("aName")%></option>
										<%
											rsb.Movenext
											Loop
											End If
											rsb.Close
											Set rsb = Nothing 
										%>
									</select> 
								</td>
							</tr>
							<tr>
								<td class="td_l_r title">类别名称</td>
								<td class="td_l_l"><input name="aName" type="text" id="aName" class="int" size="40" value="<%=aName%>" /></td>
								<input name="aId" type="hidden" id="aId" value="<% = aId %>">
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div class="fixed_bg_B">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n Bottom_pd "> 
							<input type="submit" name="Submit" class="button45" value="保存">　
							<input name="Back" type="button" id="Back" class="button43" value="关闭" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveSmallClassEdit" then
	aId = Request.Form("aId")
	aFId = Trim(Request.Form("aFId"))
	aName = Trim(Request.Form("aName"))
	If aFId = "" Then
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassEdit&aId="&aId&"&tipinfo=地区大类不能为空' ;</script>")
		Exit Sub
	End If
	If aName = "" Then
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassEdit&aId="&aId&"&tipinfo=地区小类不能为空' ;</script>")
		Exit Sub
	End If
	
	Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [AreaData] Where aFId = '"&aFId&"' And aName = '"&aName&"' And aId <> "&aId,conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassEdit&aId="&aId&"&tipinfo=已存在！' ;</script>")
		Response.End()
		End If
		rs.Close
		
		rs.Open "Select * From AreaData Where aId = " & aId,conn,3,2
		rs("aFId") = aFId
		rs("aName") = aName
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
		
elseif sType="AreaDataClassDel" then '删除地区分类

	aId = Request("aId")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [AreaData] Where aFId = '"&aId&"'",conn,1,1 '判断当前分类下是否存在子分类
	If rs.RecordCount > 0 Then
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=ClassList&tipinfo=有子分类，禁止删除！';</script>")
	else
		Set rss = Server.CreateObject("ADODB.Recordset")
		rss.Open "Select * From [AreaData] Where aId = " & aId,conn,3,2
		If rss.RecordCount > 0 Then
			rss.Delete
			rss.Update
		End If
		rss.Close
		Set rss = Nothing
		Response.Redirect("GetAreaData.asp?action=AreaData&sType=ClassList")
	end if
	rs.Close
	Set rs = Nothing
end if
End Sub
%>

<div id="mjs:tip" class="tip" style="position:absolute;left:0;top:0;display:none;margin-left:10px;"></div>
<script src="../data/calendar/WdatePicker.js"></script>
</body><% Set EasyCrm = nothing %>