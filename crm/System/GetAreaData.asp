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
 
Sub AreaData() '�������ݸ���

if tipinfo<>"" then
	Response.Write("<script>art.dialog({title: '��ʾ',time: 1,icon: 'warning',content: '"&tipinfo&"'});</script>")
end if

if sType="Import" then '����ȫ������
%>
		<form name="Save" action="GetAreaData.asp?action=AreaData&sType=ImportAdd" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="60" /><col /><col width="60" /><col /><col width="60" /><col /><col width="60" /><col /><col width="60" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="10"><B>ѡ��Ҫ�����ʡ�ݣ������ظ����룩</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">������</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport1" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">�����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport2" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">�Ϻ���</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport3" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">������</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport4" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">�ӱ�</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport5" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">ɽ��</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport6" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">���ɹ�</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport7" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport8" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport9" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">������</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport10" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport11" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">�㽭</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport12" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport13" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport14" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport15" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">ɽ��</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport16" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport17" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport18" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport19" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">�㶫</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport20" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport21" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">�Ĵ�</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport22" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport23" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport24" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport25" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">�ຣ</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport26" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'�ຣ' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport27" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport28" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport29" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport30" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
							<tr>
								<td class="td_l_r title">�½�</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport31" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">̨��</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport32" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport33" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">���</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport34" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'���' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
								<td class="td_l_r title">����</td>
								<td class="td_l_l"><input type="checkbox" name="AreaImport35" id="AreaImport" value="1" <%if EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId") = 0 then%>checked <%else%> disabled readonly <%end if%> ></td>
							</tr>
						</table>
					</td>
				</tr>
				<script language=javascript> 
				function selectall(id){ //��id����  
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
							<input type="button" name="checkall" class="button42" onclick="javascript:selectall('AreaImport')" value="ȫѡ/��ѡ">��
							<input type="submit" name="Submit" class="button45" value="��������">��
							<input name="Back" type="button" id="Back" class="button43" value="�ر�" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="ImportAdd" then '��ʼ����

	AreaImport = ""
	for i = 1 to 35
		if Request("AreaImport" & i) = "1" then
			AreaImport = AreaImport & "1"
		else
			AreaImport = AreaImport & "0"
		end if
	next
	
	if mid(AreaImport, 1, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'������')") 
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��̨')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'ʯ��ɽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ͷ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ɽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'ͨ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'˳��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ƽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'ƽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		Response.Write("<script>art.dialog({title: '��ʾ',time: 0.5,icon: 'warning',content: '�����е������ݵ���ɹ�'});</script>")
	end if
	
	if mid(AreaImport, 2, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'��ƽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'�Ӷ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'�Ͽ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'�ӱ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�����' and aFid='0'","aId")&",'����')")
	end if
	
	if mid(AreaImport, 3, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'�Ϻ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'¬��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'բ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'��ɽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'�ζ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'�ֶ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'��ɽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'�ɽ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'�ϻ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ϻ���' and aFid='0'","aId")&",'����')")
	end if
	
	if mid(AreaImport, 4, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ɿ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'ɳƺ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�ϰ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ʢ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'˫��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�山')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'ǭ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�뽭')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'ͭ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�ٲ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ɽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ƽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�ǿ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�ᶼ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�潭')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��¡')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ɽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��Ϫ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'ʯ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ɽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ˮ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�ϴ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�ϴ�')")
	end if
	
	if mid(AreaImport, 5, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'�ӱ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'ʯ��ׯ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'�ػʵ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'��̨��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'�żҿ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'�е���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'�ȷ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ӱ�' and aFid='0'","aId")&",'��ˮ��')")
	end if
	
	if mid(AreaImport, 6, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'̫ԭ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'��ͬ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'��Ȫ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'˷����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'�˳���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'�ٷ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
	end if
	
	if mid(AreaImport, 7, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'���ɹ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'���ͺ�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'��ͷ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'�ں���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'ͨ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'������˹��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'���ױ�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'�����׶���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'�����첼��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'���ֹ�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'�˰���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���ɹ�' and aFid='0'","aId")&",'��������')")
	end if
	
	if mid(AreaImport, 8, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��˳��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��Ϫ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'Ӫ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�̽���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��«����')")
	end if
	
	if mid(AreaImport, 9, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ƽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��Դ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ͨ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ԭ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�׳���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�ӱ�')")
	end if
	
	if mid(AreaImport, 10, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'���������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�׸���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'˫Ѽɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'ĵ������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��ľ˹��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'��̨����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�ں���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'�绯��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'������' and aFid='0'","aId")&",'���˰���')")
	end if
	
	if mid(AreaImport, 11, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�Ͼ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ͨ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'���Ƹ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�γ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'̩����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��Ǩ��')")
	end if
	
	if mid(AreaImport, 12, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'�㽭')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'̨����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㽭' and aFid='0'","aId")&",'��ˮ��')")
	end if
	
	if mid(AreaImport, 13, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�Ϸ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�ߺ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ͭ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
	end if
	
	if mid(AreaImport, 14, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'Ȫ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ƽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
	end if
	
	if mid(AreaImport, 15, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�ϲ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'Ƽ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�Ž���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ӥ̶��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�˴���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
	end if
	
	if mid(AreaImport, 16, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'�ൺ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'�Ͳ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'��ׯ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'��Ӫ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'��̨��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'Ϋ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'̩����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'�ĳ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'ɽ��' and aFid='0'","aId")&",'������')")
	end if
	
	if mid(AreaImport, 17, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'֣����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ƽ��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�ױ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����Ͽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�ܿ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'פ�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��Դ��')")
	end if
	
	if mid(AreaImport, 18, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�人��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ʯ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ʮ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�˲���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�差��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'Т����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�Ƹ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'Ǳ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ũ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ʩ')")
	end if
	
	if mid(AreaImport, 19, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ɳ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��̶��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�żҽ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'¦����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
	end if
	
	if mid(AreaImport, 20, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'�㶫')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'�麣��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'��ͷ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'�ع���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'տ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'ï����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'÷����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'��β��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'��Դ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'��Զ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'��ݸ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�㶫' and aFid='0'","aId")&",'�Ƹ���')")
	end if
	
	if mid(AreaImport, 21, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ˮ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��Ҵ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ƽ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��Ȫ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'¤����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
	end if
	
	if mid(AreaImport, 22, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'�Ĵ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'�ɶ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'�Թ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'��֦����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'��Ԫ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'�ڽ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'�ϳ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'üɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'�˱���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'�㰲��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'�Ű���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�Ĵ�' and aFid='0'","aId")&",'��ɽ')")
	end if
	
	if mid(AreaImport, 23, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����ˮ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��˳��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ͭ�ʵ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�Ͻڵ���')")
	end if
	
	if mid(AreaImport, 24, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ָɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�Ĳ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�Ͳ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�ٸ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ɳ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�ֶ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ˮ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ͤ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
	end if
	
	if mid(AreaImport, 25, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��Ϫ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ͨ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'˼é��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�ٲ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ɽ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��˫����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�º�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ŭ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
	end if
	
	if mid(AreaImport, 26, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'�ຣ')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ຣ' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ຣ' and aFid='0'","aId")&",'��������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ຣ' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ຣ' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ຣ' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ຣ' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ຣ' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�ຣ' and aFid='0'","aId")&",'����')")
	end if
	
	if mid(AreaImport, 27, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ͭ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'μ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�Ӱ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
	end if
	
	if mid(AreaImport, 28, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'���Ǹ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ɫ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�ӳ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
	end if
	
	if mid(AreaImport, 29, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ɽ�ϵ���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�տ������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��֥����')")
	end if
	
	if mid(AreaImport, 30, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ʯ��ɽ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ԭ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
	end if
	
	if mid(AreaImport, 31, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'�½�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'��³ľ����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'����������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'ʯ�����С�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'��������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'ͼľ�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'�������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'��³����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'��������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'��ʲ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'��ͼʲ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'�������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'��Ȫ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'�½�' and aFid='0'","aId")&",'����̩��')")
	end if
	
	if mid(AreaImport, 32, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'̨��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'̨����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'��¡��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'̨����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'̨����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'̨����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'��԰��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'̨����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'�û���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'��Ͷ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'̨����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'̨����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'̨��' and aFid='0'","aId")&",'������')")
	end if
	
	if mid(AreaImport, 33, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
	end if
	
	if mid(AreaImport, 34, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���' and aFid='0'","aId")&",'��۵�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'���' and aFid='0'","aId")&",'�½�')")
	end if
	
	if mid(AreaImport, 35, 1) = "1" then
		conn.execute("insert into [AreaData] (aFId,aName) values(0,'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�Ĵ�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'Ӣ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'���ô�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�¹�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'ӡ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ɫ��')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'�ձ�')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'����')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'������')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'���')")
		conn.execute("insert into [AreaData] (aFId,aName) values("&EasyCrm.getNewItem("AreaData","aName","'����' and aFid='0'","aId")&",'��ʿ')")
	end if
	
	
		Response.Write("<script>art.dialog.close();art.dialog.open.origin.location.reload();</script>")
%>
<%
elseif sType="BigClassAdd" then '��Ӵ���
%>
		<form name="Save" action="GetAreaData.asp?action=AreaData&sType=SaveBigClassAdd" method="post">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>��ϸ����</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">�������</td>
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
							<input type="submit" name="Submit" class="button45" value="����">��
							<input name="Back" type="button" id="Back" class="button43" value="�ر�" onClick="art.dialog.close();">
					</td>
				</tr>
			</table>
			</div>
		</form>
<%
elseif sType="SaveBigClassAdd" then
		aName = Request.Form("aName")
		If aName = "" Then
			Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=BigClassAdd&tipinfo=��������������Ϊ��';</script>")
			Exit Sub
		End If
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [AreaData] Where aName = '"&aName&"' ",conn,3,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=BigClassAdd&tipinfo=�Ѵ��ڣ�';</script>")
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

elseif sType="BigClassEdit" then '�޸Ĵ���
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
								<td class="td_l_l" COLSPAN="2"><B>��ϸ����</B></td>
							</tr>
							<tr>
								<td class="td_l_r title">�������</td>
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
							<input type="submit" name="Submit" class="button45" value="����">��
							<input name="Back" type="button" id="Back" class="button43" value="�ر�" onClick="art.dialog.close();">
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
			Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=BigClassEdit&aId="&aId&"&tipinfo=�������಻��Ϊ��';</script>")
			Exit Sub
		End If
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [AreaData] Where aName = '"&aName&"' And aId <> " & aId,conn,3,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=BigClassEdit&aId="&aId&"&tipinfo=�Ѵ��ڣ�';</script>")
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

elseif sType="SmallClassAdd" then '���С��
		aFid = Request("aFid")
%>
		<form name="Save" action="GetAreaData.asp?action=AreaData&sType=SaveSmallClassAdd" method="post" onSubmit="return CheckInput();">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="td_n pdl10 pdr10 pdt10"> 
						<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
							<col width="100" />
							<tr class="tr_t"> 
								<td class="td_l_l" COLSPAN="2"><B>��ϸ����</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title">�ϼ�����</td>
								<td class="td_r_l">
									<select name="aFid" class="int">
										<option value="">��ѡ��</option>
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
								<td class="td_l_r title">�������</td>
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
							<input type="submit" name="Submit" class="button45" value="����">��
							<input name="Back" type="button" id="Back" class="button43" value="�ر�" onClick="art.dialog.close();">
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
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassAdd&aFId="&aFId&"&tipinfo=�������಻��Ϊ��';</script>")
		Exit Sub
	End If
	If aName = "" Then
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassAdd&aFId="&aFId&"&tipinfo=����С�಻��Ϊ��';</script>")
		Exit Sub
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [AreaData] Where aFId='"&aFId&"' and aName = '" & aName & "'",conn,3,2
	If rs.RecordCount > 0 Then
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassAdd&aFId="&aFId&"&tipinfo=�Ѵ��ڣ�' ;</script>")
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

elseif sType="SmallClassEdit" then '�༭С��
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
								<td class="td_l_l" COLSPAN="2"><B>��ϸ����</B></td>
							</tr>
							<tr> 
								<td class="td_l_r title">�ϼ�����</td>
								<td class="td_r_l">
									<select name="aFId" class="int">
										<option value="">��ѡ��</option>
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
								<td class="td_l_r title">�������</td>
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
							<input type="submit" name="Submit" class="button45" value="����">��
							<input name="Back" type="button" id="Back" class="button43" value="�ر�" onClick="art.dialog.close();">
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
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassEdit&aId="&aId&"&tipinfo=�������಻��Ϊ��' ;</script>")
		Exit Sub
	End If
	If aName = "" Then
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassEdit&aId="&aId&"&tipinfo=����С�಻��Ϊ��' ;</script>")
		Exit Sub
	End If
	
	Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [AreaData] Where aFId = '"&aFId&"' And aName = '"&aName&"' And aId <> "&aId,conn,1,1
		If rs.RecordCount > 0 Then
			Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=SmallClassEdit&aId="&aId&"&tipinfo=�Ѵ��ڣ�' ;</script>")
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
		
elseif sType="AreaDataClassDel" then 'ɾ����������

	aId = Request("aId")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From [AreaData] Where aFId = '"&aId&"'",conn,1,1 '�жϵ�ǰ�������Ƿ�����ӷ���
	If rs.RecordCount > 0 Then
        Response.Write("<script>location.href='GetAreaData.asp?action=AreaData&sType=ClassList&tipinfo=���ӷ��࣬��ֹɾ����';</script>")
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