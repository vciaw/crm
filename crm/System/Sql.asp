<!--#include file="../data/conn.asp" -->
<%
	'��ȡgetֵ
	action 	= 	Request.QueryString("action")
	otype	=	Request.QueryString("otype")
	if otype="" then otype="Main"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=GBK" />
<link href="<%=SiteUrl&skinurl%>Style/common.css" rel="stylesheet" type="text/css">
<title></title>
<script>
function checkInput()
{
	if(document.getElementById('refinput').value == ""){
		alert("����Ϊ��");
		document.getElementById('refinput').focus();
		return false;
	}
	if(document.getElementById('refselect').value == ""){
		alert("����Ϊ��");
		document.getElementById('refselect').focus();
		return false;
	}
}

function checkclick(msg){
	if(confirm(msg)){
		event.returnValue=true;
		}
	else{
	event.returnValue=false;
	}
}

function copyToClipBoard(){ 
var clipBoardContent=document.title;
clipBoardContent+='\r\n' + document.location.href;
window.clipboardData.setData("Text",clipBoardContent);
alert("���Ƴɹ�����ճ�����Ϸ�·�������\r\n\r\n�������£�\r\n" +clipBoardContent);
}
</script>
</head>

<body style="padding-top:35px;">

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">��ǰλ�ã�ϵͳ���� > ���ݿ����</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="ˢ��" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="����" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="ǰ��" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
            <div class="MenuboxS">
              <ul>
                <li <%if otype="Main" then%>class="hover"<%end if%>><span><a href="?otype=Main&action=Main">���ù���</a></span></li>
                <li <%if otype="DBmanage" then%>class="hover"<%end if%>><span><a href="?otype=DBmanage&action=DBmanage">���ݿ����</a></span></li>
                <li <%if otype="DBsql" then%>class="hover"<%end if%>><span><a href="?otype=DBsql&action=DBsql">SQL���</a></span></li>
                <li <%if otype="DBbak" then%>class="hover"<%end if%>><span><a href="?otype=DBbak&action=DBbak">�������ݿ�</a></span></li>
				<%if Accsql=0 then%>
                <li <%if otype="DBCompress" then%>class="hover"<%end if%>><span><a href="?otype=DBCompress&action=DBCompress">ѹ�����ݿ�</a></span></li>
				<%end if%>
                <!--<li <%if otype="DBrestored" then%>class="hover"<%end if%>><span><a href="?otype=DBrestored&action=DBrestored">�ָ����ݿ�</a></span></li>-->
              </ul>
            </div>
		</td>
	</tr>
<%
dim i,rs,sql

	
    '�û���	��Ʊ�	�򿪱�	�½���	�½��ֶ�	ɾ����	�����ֶ��޸�
	'cz=1	cz=2	cz=3	cz=4 	cz=5 		cz=6	cz=7	
	'ɾ���ֶ�	����	SQL���	ִ��SQL	�������ݿ�	ִ�б������ݿ�	��ԭ���ݿ�	ִ�л�ԭ���ݿ�
	'cz=8		cz=9	cz=10	cz=11	cz=12		cz=13			cz=14		cz=15
	
Select Case action
Case "EditMssql" 			'���ݿ�������Ϣ
    Call EditMssqllink()
Case "DBmanage" 			'�û���
    Call DBmanage()
Case "DBdesign" 			'��Ʊ�
    Call DBdesign()
Case "DBopen" 				'�򿪱� 
    Call DBopen()
Case "DBaddnew" 			'�½���
    Call DBaddnew()
Case "DBaddnewfield" 		'�½��ֶ�
    Call DBaddnewfield()
Case "DBdel"				'ɾ����
    Call DBdel()
Case "DBsavefield"			'�����ֶ��޸� 
    Call DBsavefield()
Case "DBdelfield"			'ɾ���ֶ�
    Call DBdelfield()
Case "DBsave"				'����
    Call DBsave()
Case "DBsql"				'SQL���
    Call DBsql()
Case "DBsqlsub"				'ִ��SQL���
    Call DBsqlsub()
Case "DBbak"				'�������ݿ�
    Call DBbak()
Case "DBbacksave"			'ִ�б������ݿ�
if Accsql=1 then
    Call DBbacksave()
else
    Call DBbacksaveacc()
end if
Case "DBCompress"			'ѹ�����ݿ�
    Call DBCompress()
Case "DBCompresssave"		'ִ��ѹ�����ݿ�
    Call DBCompresssave()
Case "DBrestored"			'�ָ����ݿ�
    Call DBrestored()
Case "DBrestoredsave"		'ִ�л�ԭ���ݿ�
    Call DBrestoredsave()
Case Else
    Call main()
End Select

if trim(request.Cookies("linkok"))="yes" then
    if not IsObject(conn) then
        LinkData
    end if
%>
	<% Sub main()%>
	<tr>
		<td valign="top" class="td_n pd10">
		<form name="login" action="sql.asp?action=EditMssql" method="post">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="130" />
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="2"><B>���ݿ����� </B></td>
				</tr>
				<Tr>
					<TD class="td_l_r title" width="100">���ݿ�����</TD>
					<TD class="td_l_l">
						<input name="Accsql" type="radio" class="noborder" value="1" <%if Accsql=1 then%>checked<%end if%>> Mssql2005��
						<input name="Accsql" type="radio" class="noborder" value="0" <%if Accsql=0 then%>checked<%end if%>> Access
					</TD>
				</TR>
				<Tr>
					<TD class="td_l_r title" width="100">���ݿ�����</TD>
					<TD class="td_l_l"><input name="item1" type="text" id="item1" value="<%if Data_Source<>"" then%><%=Data_Source%><%else%>(local)<%end if%>" class="setup_int" style=" width:150px;"> ������LENOVO\SQLEXPRESS��</TD>
				</TR>
				<Tr>
					<TD class="td_l_r title">���ݿ�����</TD>
					<TD class="td_l_l "><input name="item2" type="text" id="item2" value="<%if Data_Catalog<>"" then%><%=Data_Catalog%><%end if%>" class="setup_int" style=" width:150px;"> ������Easycrm��</TD>
				</TR>
				<Tr>
					<TD class="td_l_r title">���ݿ��û�</TD>
					<TD class="td_l_l "><input name="item3" type="text" id="item3" value="<%if Data_User<>"" then%><%=Data_User%><%end if%>" class="setup_int" style=" width:150px;"> ������sa��</TD>
				</TR>
				<Tr>
					<TD class="td_l_r title">���ݿ�����</TD>
					<TD class="td_l_l "><input name="item4" type="text" id="item4" value="<%if Data_Password<>"" then%><%=Data_Password%><%end if%>" class="setup_int" style=" width:150px;"></TD>
				</TR>
				<Tr>
					<TD class="td_l_r title">ACC���ݿ�</TD>
					<TD class="td_l_l "><input name="item5" type="text" id="item5" value="<%if Data_MDBPath<>"" then%><%=Data_MDBPath%><%end if%>" class="setup_int" style=" width:150px;"></TD>
				</TR>
				<tr class="tr_f"> 
					<td class="td_l_l fontnobold" COLSPAN="2"><input type="Submit" class="button45" value="�޸�" />��<font color="#color:#CC0000">��</font> ע�⣺�˲�����һ�����ա�</td>
				</tr> 
			</table>
		</form>
		</td> 
	</tr>
	<%
	end Sub
	
	Sub EditMssqllink()
	Data_Source = replace(Trim(Request.Form("item1")),CHR(34),"'")
	Data_Catalog = replace(Trim(Request.Form("item2")),CHR(34),"'")
	Data_User = replace(Trim(Request.Form("item3")),CHR(34),"'")
	Data_Password = replace(Trim(Request.Form("item4")),CHR(34),"'")	
	Data_MDBPath = replace(Trim(Request.Form("item5")),CHR(34),"'")	
	Accsql = replace(Trim(Request.Form("Accsql")),CHR(34),"'")	
	Dim TempStr
	TempStr = ""
	TempStr = TempStr & chr(60) & "%" & VbCrLf
	TempStr = TempStr & "Dim Accsql,Data_MDBPath,Data_Source,Data_User,Data_Password,Data_Catalog,SystemNumber" & VbCrLf
	
	TempStr = TempStr & "'���ݿ�����" & VbCrLf
	TempStr = TempStr & "Accsql="&Accsql&" '���ݿ�����" & VbCrLf
	TempStr = TempStr & "Data_Source="& Chr(34) & Data_Source & Chr(34) &" 'MSSQL����Դ��������\ʵ���� �� IP��ַ\ʵ������" & VbCrLf
	TempStr = TempStr & "Data_Catalog="& Chr(34) & Data_Catalog & Chr(34) &" '���ݿ�����" & VbCrLf
	TempStr = TempStr & "Data_User="& Chr(34) & Data_User & Chr(34) &" '���ݿ��û�" & VbCrLf
	TempStr = TempStr & "Data_Password="& Chr(34) & Data_Password & Chr(34) &" '���ݿ�����" & VbCrLf & VbCrLf
	TempStr = TempStr & "Data_MDBPath="& Chr(34) & Data_MDBPath & Chr(34) &" 'Access���ݿ�·��" & VbCrLf & VbCrLf
	TempStr = TempStr & "SystemNumber="& Chr(34) & SystemNumber & Chr(34) &" '��Ȩ��" & VbCrLf & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf
	ADODB_SaveToFile TempStr,"../data/Mssql.asp"
	Response.Write("<script>alert(""�޸ĳɹ���"");</script>")
	Response.Write "<script>location.href='?otype=Main&action=Main';</script>"
	end Sub
	
	Sub dbmanage()
        set rsSchema=conn.openSchema(20) 
        rsSchema.movefirst 
    %>
	<tr>
		<td colspan=2 class="Search_All td_n">
		 <form action="sql.asp?action=DBaddnew" method="post" onSubmit="return checkInput();">���� <input type="text" id="refinput" name="crtablename" class="int" size="25" > <input type="submit" class="button245" value=" �����±� " id="submit">��<span class="info_help help01">�½��ı�Ĭ�ϴ���һ��ID�ֶ�,�����������ͣ�������������</span></form>
		</td>
	</tr>
	<tr>
		<td valign="top" colspan=2 style="padding:0 10px 10px 10px;" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_c">����</td>
					<td class="td_l_c" width="30%" colspan="4">����</td>
					<td class="td_l_c" width="30%" colspan="4">���¼��SQL��䴦��</td>
				</tr>
		<%
        Do Until rsSchema.EOF
            if rsSchema("TABLE_TYPE")="TABLE" then
		%>
				<tr class="tr"> 
					<form action="sql.asp?action=DBsave&tablename2=<%=rsSchema("TABLE_NAME")%>" method="post">
					<td class="td_r_l"><input type="text" name="tablename" value="<%=rsSchema("TABLE_NAME")%>" class="int" size="25"></td>
					<td class="td_r_c"><input type="submit" class="button227" value="����"></td>
					</form>
					<td class="td_r_c"><a href="?action=DBdesign&otype=DBmanage&tablename=<%=rsSchema("TABLE_NAME")%>">��Ʊ�</a></td>
					<td class="td_r_c"><a href="?action=DBopen&otype=DBmanage&tablename=<%=rsSchema("TABLE_NAME")%>">�򿪱�</a></td>
					<td class="td_r_c"><a onclick="checkclick('��ȷ��Ҫɾ���ñ������������������?')" href="?action=DBdel&tablename=<%=rsSchema("TABLE_NAME")%>">ɾ����</a></td>
					<td class="td_r_c"><a href="?action=DBsql&otype=DBsql&czsql=1&tablename=<%=rsSchema("TABLE_NAME")%>">��ѯ</a></td>
					<td class="td_r_c"><a href="?action=DBsql&otype=DBsql&czsql=2&tablename=<%=rsSchema("TABLE_NAME")%>">����</a></td>
					<td class="td_r_c"><a href="?action=DBsql&otype=DBsql&czsql=3&tablename=<%=rsSchema("TABLE_NAME")%>">����</a></td>
					<td class="td_r_c"><a href="?action=DBsql&otype=DBsql&czsql=4&tablename=<%=rsSchema("TABLE_NAME")%>">ɾ��</a></td>
				</tr>
		<%
                end if
            rsSchema.movenext
        Loop
        rsSchema.close
        set rsSchema=Nothing
		
		%>
			</table>
		</td> 
	</tr>
	
	<%
	end Sub
	
    Sub DBdesign()
        dim fieldCount
        set rs=conn.execute("select * from ["&trim(request.QueryString("tablename"))&"]")
        fieldCount = rs.Fields.Count
	%>
	<tr>
		<td colspan=2 class="Search_All td_n">
		 <form action="sql.asp?action=DBaddnewfield&tablename=<%=trim(request.QueryString("tablename"))%>" method="post" id="form1" name="form1" onSubmit="return checkInput();">�ֶ��� <input type="text" id="refinput" name="crfield" size="25" class="int" > 
			<select name="fieldtype" class="int" id="refselect">
				<option value="">�ֶ�����</option>
				<option value="int">int</option>
				<option value="bigint">bigint</option>
				<option value="smallint">smallint</option>
				<option value="varchar">varchar</option>
				<option value="ntext">ntext</option>
				<option value="float">float</option>
				<option value="bit">bit</option>
				<option value="nvarchar">nvarchar</option>
				<option value="datetime">datetime</option>
				<option value="image">image</option>
				<option value="text">text</option>
				<option value="nchar">nchar</option>
				<option value="money">money</option>
				<option value="smalldatetime">smalldatetime</option>
				<option value="numeric">numeric</option>
				<option value="varbinary">varbinary</option>
				<option value="tinyint">tinyint</option>
				<option value="timestamp">timestamp</option>
				<option value="sql_variant">sql_variant</option>
				<option value="real">real</option>
			</select> <input type="submit" class="button245" value=" �½��ֶ� " id="1" name="1"> (��Ʊ�)��<B style="color:#f00;"><%=trim(request.QueryString("tablename"))%></B>����<%=fieldCount%>���ֶ�</form>
		</td>
	</tr>
	<tr>
		<td valign="top" colspan=2 style="padding:0 10px 10px 10px;" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_c">�ֶ�����</td>
					<td class="td_l_c">�ֶ�����</td>
					<td class="td_l_c">�ֶγ���</td>
					<td class="td_l_c" colspan="2">����</td>
				</tr>
				<% For i=0 to fieldCount - 1 %>
				<tr class="tr"> 
					<form action="sql.asp?action=DBsavefield&tablename=<%=trim(request.QueryString("tablename"))%>" method="post" onSubmit="return checkInput();">
					<td class="td_l_c"><input type="text" name="fieldsname" id="refinput" value="<%=rs.Fields(i).Name%>" size="10"><input type="hidden" name="fieldsname2" value="<%=rs.Fields(i).Name%>"></td>
					<td class="td_l_c">
						<select name="fieldtype" id="refselect">
						<%
							select case rs.Fields(i).type
							case 3
								Response.Write "<option value=""int"">int</option>"
							case 5
								Response.Write "<option value=""float"">float</option>"
							case 11
								Response.Write "<option value=""bit"">bit</option>"
							case 20
								Response.Write "<option value=""bigint"">bigint</option>"
							case 130
								Response.Write "<option value=""nchar"">nchar</option>"
							case 200
								Response.Write "<option value=""varchar"">varchar</option>"
							case 202
								Response.Write "<option value=""nvarchar"">nvarchar</option>"
							case 203
								Response.Write "<option value=""ntext"">ntext</option>"
							case 205
								Response.Write "<option value=""image"">image</option>"
							case 135
								Response.Write "<option value=""datetime"">datetime</option>"
							case else
								Response.Write "<option value="""">"&rs.Fields(i).type&"</option>"
							end select
						%>
							<option value="int">int</option>
							<option value="bigint">bigint</option>
							<option value="smallint">smallint</option>
							<option value="varchar">varchar</option>
							<option value="ntext">ntext</option>
							<option value="float">float</option>
							<option value="bit">bit</option>
							<option value="nvarchar">nvarchar</option>
							<option value="datetime">datetime</option>
							<option value="image">image</option>
							<option value="text">text</option>
							<option value="nchar">nchar</option>
							<option value="money">money</option>
							<option value="smalldatetime">smalldatetime</option>
							<option value="numeric">numeric</option>
							<option value="varbinary">varbinary</option>
							<option value="tinyint">tinyint</option>
							<option value="timestamp">timestamp</option>
							<option value="sql_variant">sql_variant</option>
							<option value="real">real</option>
							</select>
					</td>
					<td class="td_l_c"><input name="fieldssize" type="text" value="<%=rs.Fields(i).DefinedSize%>" size="10"></td>
					<td class="td_l_c" width="10%"><input type="submit" class="button227" value="����"></td>
					<td class="td_l_c" width="10%"><a onclick="checkclick('��ȷ��Ҫɾ�����ֶΣ������������������?')" href="sql.asp?action=DBdelfield&tablename=<%=trim(request.QueryString("tablename"))%>&fieldsname=<%=rs.Fields(i).Name%>">ɾ��</a></td>
					</form>
				</tr>
				<%
				Next
				rs.close
				set rs=nothing
				%>
			</table>
		</td> 
	</tr>
	<%
	end Sub
	
    Sub DBopen()
	%>
	<tr>
		<td colspan=2 class="Search_All td_n">
			<a href="?action=DBsql&czsql=1&tablename=<%=trim(request.QueryString("tablename"))%>">���¼��ѯ</a> | <a href="?action=DBsql&czsql=2&tablename=<%=trim(request.QueryString("tablename"))%>">����</a> | <a href="?action=DBsql&czsql=3&tablename=<%=trim(request.QueryString("tablename"))%>">����</a> | <a href="?action=DBsql&czsql=4&tablename=<%=trim(request.QueryString("tablename"))%>">ɾ��</a> | ��<span class="info_help help01">ֻ��ʾǰ10����¼</span>
		</td>
	</tr>
	<tr>
		<td valign="top" colspan=2 style="padding:0 10px 10px 10px;" class="td_n">
		<%
        set rs=conn.execute("select top 10 * from ["&trim(request.QueryString("tablename"))&"]")
        fieldCount = rs.Fields.Count
		%>
			<table border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
				<% For i=0 to fieldCount - 1 %>
					<td class="td_l_c"><%=rs.Fields(i).Name%></td>
				<% Next %>
				</tr>
				<% while not rs.eof %>
				<tr class="tr"> 
					<% For i=0 to fieldCount - 1 %>
					<td class="td_l_c"><TEXTAREA class="int" cols="15" style="height:23px;margin:5px;"><%if ISEMPTY(rs(i)) then Response.Write () else Response.Write rs(i) end if %></TEXTAREA></td>
					<% Next %>
				</tr>
				<%
				rs.movenext
				wend
				rs.close
				set rs=nothing
				%>
			</table>
		</td> 
	</tr>
	
	<%
	End Sub
	
    Sub DBaddnew()
        dim crtablename
        crtablename=trim(request.Form("crtablename"))
        crtable("CREATE TABLE ["&crtablename&"] (ID int IDENTITY (1,1) not null PRIMARY key)")
	End Sub
	
    Sub DBaddnewfield()
        dim crfield
        tablename=trim(request.QueryString("tablename"))
        crfield=trim(request.Form("crfield"))
        fieldtype=trim(request.Form("fieldtype"))
        select case fieldtype
        case ""
            Response.Write "��ѡ���ֶ�����"
        case "varchar"
            crtable("ALTER TABLE ["&tablename&"] ADD ["&crfield&"] varchar(255)")
        case "nvarchar"
            crtable("ALTER TABLE ["&tablename&"] ADD ["&crfield&"] nvarchar(50)")
        case else
            crtable("ALTER TABLE ["&tablename&"] ADD ["&crfield&"] "&fieldtype&"")
        end select
	End Sub
	
    Sub DBdel()
        tablename=trim(request.QueryString("tablename"))
        crtable("DROP TABLE ["&tablename&"]")
	End Sub
	
    Sub DBsavefield()
        dim fieldsname,fieldsname2,fieldssize,fieldar
        tablename=trim(request.QueryString("tablename"))
        fieldsname=trim(request.Form("fieldsname"))
        fieldsname2=trim(request.Form("fieldsname2")) 'ԭ����
        fieldtype=trim(request.Form("fieldtype"))
        crtable("sp_rename '"&tablename&"."&fieldsname2&"','"&fieldsname&"','column';") '�ֶ����޸�
        
        fieldssize=trim(request.Form("fieldssize"))
        fieldar=""
        select case fieldtype
        case "varchar","nvarchar"
            fieldar="("&fieldssize&")"
        end select
        if fieldssize=0 then fieldar="" end if
        crtable("ALTER TABLE ["&tablename&"] ALTER COLUMN ["&fieldsname&"] "&fieldtype&""&fieldar&"") '�ֶ����ʹ���
	End Sub
	
    Sub DBdelfield()
        tablename=trim(request.QueryString("tablename"))
        fieldsname=trim(request.QueryString("fieldsname"))
        crtable("Alter table ["&tablename&"] drop column ["&fieldsname&"]")
	End Sub
	
    Sub DBsave()
        dim tablename2
        tablename=trim(request.Form("tablename"))
        tablename2=trim(request.QueryString("tablename2"))
        crtable("EXEC sp_rename ["&tablename2&"],["&tablename&"]")
	End Sub
	
    Sub DBsql()
	%>
	<tr>
		<td valign="top" class="td_n pd10">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="130" />
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>��䰸���� </B></td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_r title">�������</td>
					<td class="td_l_l">insert into ����(�ֶ�1,�ֶ�2)values('����1','����2')</td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_r title">�������</td>
					<td class="td_l_l">update ���� set �ֶ�1='����1',�ֶ�2='����2' where �ֶ�3='����3'</td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_r title">ɾ�����</td>
					<td class="td_l_l">delete from ���� where �ֶ�='����'</td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_r title">��ѯ���</td>
					<td class="td_l_l">select top ��ʾ�ļ�¼��Ŀ �ֶ�1,�ֶ�2 from ���� where �ֶ�1='����1'</td>
				</tr> 
				<%        
				tablename=trim(request.QueryString("tablename"))
				if tablename<>"" then
					dim czsql
					czsql=""
					select case request.QueryString("czsql")
					case 1
						czsql="SELECT TOP 10 * FROM ["&tablename&"]"
					case 2
						czsql="INSERT INTO ["&tablename&"] ( ) VALUES ( )"
					case 3
						czsql="UPDATE ["&tablename&"] SET"
					case 4
						czsql="DELETE FROM ["&tablename&"]"
					end select
				end if
				%>
				<form action="sql.asp?action=DBsqlsub&otype=DBsql" method="post" onSubmit="return checkInput();">
				<tr>
					<td valign="top" class="td_l_l" colspan=2> 
					<textarea name="sqlstr" id="refinput" class="int" style="width:98%;height:150px;margin:10px 0;"><%=czsql%></textarea> 
					</td> 
				</tr>
				<tr>
					<td class="td_l_l" colspan=2> 
						<input name="Submit" type="submit" class="button45" value="ִ��SQL"> 
						<input name="Rest" type="reset" class="button43" value="��д"> 
					</td> 
				</tr>
				</form>
				<tr class="tr_f"> 
					<td class="td_l_l fontnobold" COLSPAN="6"><font color="#color:#CC0000">��</font>ע�⣺��select��ѯ��¼��ʱ�����top���,�����¼����,�ͻ������ʱ,�򲻿�������,����top�Ϳ�����ֹ��ʾ��������ѯ�Ľ����</td>
				</tr>
			</table>
		</td> 
	</tr>
	<%
	End Sub
	
    Sub DBsqlsub()
        if instr(1,trim(request.Form("sqlstr")),"select",1)>0 then
            On Error Resume Next
            set rs=conn.Execute(trim(request.Form("sqlstr")))
	        If Err Then
				Response.Write ("<script>alert("" "&Err.Description&" "");history.back(-1);</script>")
            else
				Response.Write("<script>alert('ִ�У�"&trim(request.Form("sqlstr"))&" �ɹ���');</script>")
                fieldCount = rs.Fields.Count
				%>
	<tr>
		<td valign="top" colspan=2 class="td_n pd10">
			<table border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
				<% For i=0 to fieldCount - 1 %>
					<td class="td_l_c"><%=rs.Fields(i).Name%></td>
				<% Next %>
				</tr>
				<% while not rs.eof %>
				<tr class="tr"> 
					<% For i=0 to fieldCount - 1 %>
					<td class="td_l_c"><TEXTAREA id=textarea1 name=textarea1 class="int" cols="15" style="height:23px;margin:5px;"><%if ISEMPTY(rs(i)) then Response.Write () else Response.Write rs(i) end if %></TEXTAREA></td>
					<% Next %>
				</tr>
				<%
				rs.movenext
				wend
				rs.close
				set rs=nothing
				%>
			</table>
		</td> 
	</tr>
				<%
	        end if
        else
            crtable(trim(request.Form("sqlstr")))
        end if
	End Sub
	
    Sub DBbak()
    %>
	<tr>
		<td valign="top" class="td_n pd10">
			<form action="sql.asp?action=DBbacksave" method="post" onSubmit="return checkInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<col width="130" />
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="6"><B>�������ݿ⣺ </B></td>
				</tr>
				<%if Accsql=1 then%>
				<tr> 
					<td class="td_l_r title">�ļ�����·��</td>
					<td class="td_l_l"><input type="text" size=70 name="dbpath" id="refinput" class="int" value="<%=server.MapPath("../data/bak")%>\<%=year(now())&month(now())&day(now())&hour(now())&minute(now())&second(now())%>.bak">��<span class="info_help help01">����·���ǣ�../data/bak/</span></td>
				</tr>
				<tr> 
					<td class="td_l_r title">�ҵ��ı���</td>
					<td class="td_l_l"><%=FileListBak("../data/bak","bak")%></td>
				</tr>
				<%else%>
				<tr> 
					<td class="td_l_r title">���ݿ���ʵ·��</td>
					<td class="td_l_l"><input name=DBpath type=text id="DBpath" value="<%=Data_MDBPath%>" size="40" /></td>
				</tr>
				<tr> 
					<td class="td_l_r title">�����ļ���</td>
					<td class="td_l_l"><input name=bkfolder type=text value="../data/bak/" size="40" Readonly="true"/></td>
				</tr>
				<tr> 
					<td class="td_l_r title">�������ݿ�����</td>
					<td class="td_l_l"><input name=bkDBname type=text value="<%=year(now())&month(now())&day(now())&hour(now())&minute(now())&second(now())%>.mdb" size="40" /></td>
				</tr>
				<tr> 
					<td class="td_l_r title">�ҵ��ı���</td>
					<td class="td_l_l"><%=FileListBak("../data/bak","bak")%></td>
				</tr>
				<%end if%>
				<tr> 
					<td class="td_l_l" colspan=2><input type="submit" class="button45" value="��ʼ����" id=submit1 name=submit1></td>
				</tr>
				<tr class="tr_f"> 
					<td class="td_l_l fontnobold" COLSPAN="6"><font color="#color:#CC0000">��</font> ע�⣺��������ݿⱸ���п��ܳ�ʱ�������ڷ������ٵ�ʱ�������</td>
				</tr> 
			</table>
			</form>
		</td> 
	</tr>
    <%
	End Sub
	
    Sub DBbacksave()
        dbpath = trim(Request.Form("dbpath"))
        If dbpath <> "" Then
            dim fso,Files
            Set fso = CreateObject("Scripting.FileSystemObject")
            If fso.FileExists(dbPath) Then
				Response.Write ("<script>alert(""�������ı������ݿ��Ѿ����ڣ���ɾ������������ļ������б���"");history.back(-1);</script>")
            Else
                Response.Flush
                dim srv,bak
                server.ScriptTimeout = 3600
                Set srv=Server.CreateObject("SQLDMO.SQLServer")
                srv.LoginTimeout = 3600
                srv.Connect Data_Source,Data_User,Data_Password
                Set bak = Server.CreateObject("SQLDMO.Backup")
                bak.Database=Data_Catalog
                bak.Devices=Files
                bak.Files=dbpath
                bak.SQLBackup srv
                if err.number>0 then
				Response.Write ("<script>alert("""&err.number&"��"&err.description&" "");history.back(-1);</script>")
                else
				Response.Write ("<script>alert(""���ݿ��Ѿ����ݳɹ�"");location.href='?otype=DBbak&action=DBbak';</script>")
                end if
            End If
        End If
	End Sub
	
    Sub DBbacksaveacc()
		dim Dbpath,bkfolder,bkdbname,fso
		Dbpath=request.form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.form("bkfolder")
		bkdbname=request.form("bkdbname")
		Set Fso=Server.CreateObject("Scripting.FileSystemObject")
			if fso.fileexists(dbpath) then
				If CheckDir(bkfolder) = True Then
				fso.copyfile dbpath,bkfolder& "\"& bkdbname
				else
				MakeNewsDir bkfolder
				fso.copyfile dbpath,bkfolder& "\"& bkdbname
				end if
			End if
		Response.Write ("<script>alert(""���ݿ��Ѿ����ݳɹ�"");location.href='?otype=DBbak&action=DBbak';</script>")
	End Sub
	
	'------------------���ĳһĿ¼�Ƿ����-------------------
	Function CheckDir(FolderPath)
		dim fso1
		folderpath=Server.MapPath(".")&"\"&folderpath
		Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
		If fso1.FolderExists(FolderPath) then
		   '����
		   CheckDir = True
		Else
		   '������
		   CheckDir = False
		End if
		Set fso1 = nothing
	End Function
	'-------------����ָ����������Ŀ¼-----------------------
	Function MakeNewsDir(foldername)
		dim f,fso1
		Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
			Set f = fso1.CreateFolder(foldername)
			MakeNewsDir = True
		Set fso1 = nothing
	End Function
	
    Sub DBCompress()
    %>
	<tr>
		<td valign="top" class="td_n pd10">
			<form action="sql.asp?action=DBCompresssave" method="post" onSubmit="return checkInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
			<col width="130" />
				<tr class="tr_t"> 
					<td class="td_l_l" COLSPAN="2"><B>ѹ�����ݿ⣺ </B></td>
				</tr>
				<tr> 
					<td class="td_l_r title">���ݿ�·��</td>
					<td class="td_l_l"><input name=dbpath type=text id="dbpath" value="<%=Data_MDBPath%>" size="40" /></td>
				</tr>
				<tr> 
					<td class="td_l_l" colspan=2><input type="submit" class="button45" value="��ʼѹ��" id=submit1 name=submit1></td>
				</tr>
				<tr class="tr_f"> 
					<td class="td_l_l fontnobold" COLSPAN="6"><font color="#color:#CC0000">��</font> ע�⣺��ǰ������һ�����գ�ѹ��ǰ��ȷ���Ѿ����ݡ�</td>
				</tr> 
			</table>
			</form>
		</td> 
	</tr>
    <%
	End Sub
	
	Sub DBCompresssave()
		dim dbpath
		dbpath = request("dbpath")
		If dbpath <> "" Then
		dbpath = server.mappath(dbpath)
		end if
		
		Dim fso, Engine, strDBPath,JET_3X
		strDBPath = left(dbPath,instrrev(DBPath,"\"))
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		
		If fso.FileExists(dbPath) Then
			fso.CopyFile dbpath,strDBPath & "temp.mdb"
			Set Engine = CreateObject("JRO.JetEngine")
		
				Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
				"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
		
			fso.CopyFile strDBPath & "temp1.mdb",dbpath
			fso.DeleteFile(strDBPath & "temp.mdb")
			fso.DeleteFile(strDBPath & "temp1.mdb")
			Set fso = nothing
			Set Engine = nothing
		End If
		Response.Write ("<script>alert(""���ݿ�ѹ���ɹ�"");location.href='?otype=DBCompress&action=DBCompress';</script>")
	End Sub
	
    Sub DBrestored()
    %>
	<tr>
		<td valign="top" class="td_n pd10">
			<form action="sql.asp?action=DBrestoredsave" method="post" onSubmit="return checkInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_1"> 
					<td class="td_2" COLSPAN="6"><B>�ָ����ݿ⣺ </B></td>
				</tr>
				<tr class="tr"> 
					<td class="td_2 td_title">�ļ�����·��</td>
					<td class="td_2 td_content"><input type="text" size=70 name="dbpath" id="refinput" class="int" value="">��<span class="info_help help01">�ָ�·�������ݿ��������·��</span></td>
				</tr>
				<tr class="tr"> 
					<td class="td_2 td_title">�ҵ��ı���</td>
					<td class="td_2 td_content"><%=FileList("../data/bak","bak")%>��<span class="info_help help01">ѡ��󱸷��ļ���ַ�Զ�д��ָ�·��</span></td>
				</tr>
				<tr class="tr"> 
					<td class="td_2 td_title"></td>
					<td class="td_2 td_content"><input type="submit" class="button45" value="��ʼ�ָ�" id=submit1 name=submit1></td>
				</tr>
				<tr class="tr_f"> 
					<td class="td_l_l fontnobold" COLSPAN="6"><font color="#color:#CC0000">��</font> ע�⣺���ݿ�����ʹ��ʱ�޷�������</td>
				</tr> 
			</table>
			</form>
		</td> 
	</tr>
    <%
	End Sub
	
    Sub DBrestoredsave()
        closedata
        dbpath = trim(Request.Form("dbpath"))
        If dbpath <> "" Then
            Set fso = CreateObject("Scripting.FileSystemObject")
            If not fso.FileExists(dbPath) Then
				Response.Write ("<script>alert(""û���ҵ��ñ����ļ�,(ע��·��Ϊ���ݿ��������·��)"");history.back(-1);</script>")
            Else
				Response.Write ("<script>alert(""���ڴ�"&dbPath&"��ԭ,���Ժ�"");</script>")
                Response.Flush
                server.ScriptTimeout = 3600
                Set srv=Server.CreateObject("SQLDMO.SQLServer")
                srv.LoginTimeout = 3600
                srv.Connect trim(request.Cookies("ipdress")),trim(request.Cookies("username")), trim(request.Cookies("password"))
                Set bak = Server.CreateObject("SQLDMO.Restore")
                bak.Action=0
                bak.Database=trim(request.Cookies("dataname"))
                bak.Devices=Files
                bak.Files=dbpath
                bak.ReplaceDatabase=True
                bak.SQLRestore srv
                if err.number>0 then
				Response.Write ("<script>alert("""&err.number&"��"&err.description&" "");history.back(-1);</script>")
                else
				Response.Write ("<script>alert(""���ݿ�ָ��ɹ�"");location.href='?otype=DBbak&action=DBbak';</script>")
                end if
            End If
        End If
        Response.End
	End Sub
    closedata
end if

'----------------------------------------------------
Function RequestNumSafe(qudata)
    If isNumeric(qudata) then
        if qudata="" then
            RequestNumSafe=0
        else
            RequestNumSafe=qudata
        end if
    else
        RequestNumSafe=0
    end if
End Function

Function RequestCStringSafe(cstring)
    If Instr(1,cstring,"%")>0 or Instr(1,cstring,"=")>0 or Instr(1,cstring,"&")>0 or Instr(1,cstring,"#")>0 or Instr(1,cstring,">")>0 or Instr(1,cstring,"<")>0 or Instr(1,cstring,"'")>0 or Instr(1,cstring,";")>0 or Instr(1,cstring,"��")>0 or Instr(1,cstring,"`")>0 or Instr(1,cstring,"*")>0 or Instr(1,cstring,",")>0 then
        RequestCStringSafe=""
    else
        RequestCStringSafe=cstring
    end if
End Function

Sub LinkData()
    Dim ConnStr
    ConnStr = "Provider = Sqloledb; User ID = " & Data_User & "; Password = " & Data_Password & "; Initial Catalog = " & Data_Catalog & "; Data Source = " & Data_Source & ";"
    On Error Resume Next
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open ConnStr
    If Err Then
	    err.Clear
	    Set Conn = Nothing
	    Response.Write "���ݿ����ӳ������������ִ���"
	    Response.End
	else
	    Response.Cookies("linkok")="yes"
    End If
End Sub

Sub CloseData()
    if IsObject(conn) then
        conn.Close
        set conn=nothing
    end if
End Sub

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
			GBL_CHK_TempStr = "����������֧��ADODB.Stream���޷���ɲ�������ʹ��FTP�ȹ��ܣ���<font color=Red >data/config.asp</font>�ļ������滻�ɿ�������"
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

Function crtable(SqlCommand)
	On Error Resume Next
	Conn.Execute(SqlCommand)
	If Err Then
		'Response.write ""&Err.Description&"<br>"
        Response.Write ("<script>alert("" "&Err.Description&" "");history.back(-1);</script>")
    else
        'Response.Write "ִ��:&nbsp;"&SqlCommand&"&nbsp;&nbsp;�ɹ�<br>"
		if action="DBaddnewfield" or action="DBsavefield" or action="DBdelfield" then
		Response.Write "<script>alert('ִ�У�"&SqlCommand&" �ɹ���');location.href='sql.asp?action=DBdesign&otype=DBmanage&tablename="&trim(request.QueryString("tablename"))&"';</script>"
		elseif action="DBaddnew" then
        Response.Write("<script>alert('ִ�У�"&SqlCommand&" �ɹ���');location.href='?otype=DBmanage&action=DBmanage';</script>")
		elseif action="DBsqlsub" then
        Response.Write("<script>alert('ִ�У�"&SqlCommand&" �ɹ���');history.back(-1);</script>")
		else
        Response.Write("<script>alert('ִ�У�"&SqlCommand&" �ɹ���');location.href='?otype=DBmanage&action=DBmanage';</script>")
		end if
	end if
	Response.Flush
End Function

Function FileList(FolderUrl,FileExName)
Set fso=Server.CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set folder=fso.GetFolder(Server.MapPath(Trim(FolderUrl)))
Set file=folder.Files
FileList=""
FileList=FileList&"<select onchange=""this.form.refinput.value=this.value;"" name=""FileList"" >"
    	FileList=FileList&"<option value="""">��ѡ��</option>"
For Each FileName in file
If Trim(FileExName)<>"" Then
	If InStr(Trim(FileExName),Trim(Mid(FileName.Name,InStr(FileName.Name,".")+1,len(FileName.Name))))>0 Then
    	FileList=FileList&"<option value="""&server.MapPath("../data/bak")&"\"&FileName.Name&""">"&FileName.Name&"</option>"
	End If
Else
     FileList=FileList&"<a href='#'>"&FileName.Name&"</a><br>"
End If
Next
FileList=FileList&"</select>"
Set file=Nothing
Set folder=Nothing
Set fso=Nothing
End Function

Function FileListBak(FolderUrl,FileExName)
Set fso=Server.CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set folder=fso.GetFolder(Server.MapPath(Trim(FolderUrl)))
Set file=folder.Files
FileListBak=""
For Each FileName in file
    FileListBak=FileListBak&"<a href='../data/bak\"&FileName.Name&"'>"&FileName.Name&"</a> <input type='button' class='button227' value='ɾ��'  onClick=window.location.href='?otype=DBbak&action=DBbak&Subaction=DelDataBakFlie&DataBakFlie=../data/bak/"&FileName.Name&"' /> <br>"
Next
Set file=Nothing
Set folder=Nothing
Set fso=Nothing
End Function

Subaction 	= 	Request.QueryString("Subaction")

if Subaction="DelDataBakFlie" then
	DataBakFlie=request("DataBakFlie")
	if IsObjInstalled("Scripting.FileSystemObject") then
	s= server.MapPath(DataBakFlie)
	Set fso = CreateObject("Scripting.FileSystemObject")     
	If fso.FileExists(s) Then     
	   fso.Deletefile(s)     
	End If     
	Set fso = Nothing
	end if
	Response.Write ("<script>alert(""���ݿ�ɾ���ɹ�"");location.href='?otype=DBbak&action=DBbak';</script>")
end if
%>
</table>

</body>
</html>
