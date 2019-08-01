<!--#include file="../data/conn.asp" -->
<%
	'获取get值
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
		alert("不能为空");
		document.getElementById('refinput').focus();
		return false;
	}
	if(document.getElementById('refselect').value == ""){
		alert("不能为空");
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
alert("复制成功，请粘贴到上方路径输入框！\r\n\r\n内容如下：\r\n" +clipBoardContent);
}
</script>
</head>

<body style="padding-top:35px;">

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="top_bg">
	<tr>
		<td class="top_left td_t_n td_r_n">当前位置：系统管理 > 数据库管理</td>
		<td class="top_right td_t_n td_r_n">
			<input type="button" class="button_top_reload" value=" " title="刷新" onClick=window.location.href="javascript:window.location.reload();" />
			<input type="button" class="button_top_back" value=" " title="后退" onClick=window.location.href="javascript:history.back();" />
			<input type="button" class="button_top_go" value=" " title="前进" onClick=window.location.href="javascript:history.go(1);" />
        </td>
	</tr>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign="top" class="td_n pdr10 pdt10">   
            <div class="MenuboxS">
              <ul>
                <li <%if otype="Main" then%>class="hover"<%end if%>><span><a href="?otype=Main&action=Main">配置管理</a></span></li>
                <li <%if otype="DBmanage" then%>class="hover"<%end if%>><span><a href="?otype=DBmanage&action=DBmanage">数据库操作</a></span></li>
                <li <%if otype="DBsql" then%>class="hover"<%end if%>><span><a href="?otype=DBsql&action=DBsql">SQL语句</a></span></li>
                <li <%if otype="DBbak" then%>class="hover"<%end if%>><span><a href="?otype=DBbak&action=DBbak">备份数据库</a></span></li>
				<%if Accsql=0 then%>
                <li <%if otype="DBCompress" then%>class="hover"<%end if%>><span><a href="?otype=DBCompress&action=DBCompress">压缩数据库</a></span></li>
				<%end if%>
                <!--<li <%if otype="DBrestored" then%>class="hover"<%end if%>><span><a href="?otype=DBrestored&action=DBrestored">恢复数据库</a></span></li>-->
              </ul>
            </div>
		</td>
	</tr>
<%
dim i,rs,sql

	
    '用户表	设计表	打开表	新建表	新建字段	删除表	保存字段修改
	'cz=1	cz=2	cz=3	cz=4 	cz=5 		cz=6	cz=7	
	'删除字段	保存	SQL语句	执行SQL	备份数据库	执行备份数据库	还原数据库	执行还原数据库
	'cz=8		cz=9	cz=10	cz=11	cz=12		cz=13			cz=14		cz=15
	
Select Case action
Case "EditMssql" 			'数据库连接信息
    Call EditMssqllink()
Case "DBmanage" 			'用户表
    Call DBmanage()
Case "DBdesign" 			'设计表
    Call DBdesign()
Case "DBopen" 				'打开表 
    Call DBopen()
Case "DBaddnew" 			'新建表
    Call DBaddnew()
Case "DBaddnewfield" 		'新建字段
    Call DBaddnewfield()
Case "DBdel"				'删除表
    Call DBdel()
Case "DBsavefield"			'保存字段修改 
    Call DBsavefield()
Case "DBdelfield"			'删除字段
    Call DBdelfield()
Case "DBsave"				'保存
    Call DBsave()
Case "DBsql"				'SQL语句
    Call DBsql()
Case "DBsqlsub"				'执行SQL语句
    Call DBsqlsub()
Case "DBbak"				'备份数据库
    Call DBbak()
Case "DBbacksave"			'执行备份数据库
if Accsql=1 then
    Call DBbacksave()
else
    Call DBbacksaveacc()
end if
Case "DBCompress"			'压缩数据库
    Call DBCompress()
Case "DBCompresssave"		'执行压缩数据库
    Call DBCompresssave()
Case "DBrestored"			'恢复数据库
    Call DBrestored()
Case "DBrestoredsave"		'执行还原数据库
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
					<td class="td_l_l" COLSPAN="2"><B>数据库连接 </B></td>
				</tr>
				<Tr>
					<TD class="td_l_r title" width="100">数据库类型</TD>
					<TD class="td_l_l">
						<input name="Accsql" type="radio" class="noborder" value="1" <%if Accsql=1 then%>checked<%end if%>> Mssql2005　
						<input name="Accsql" type="radio" class="noborder" value="0" <%if Accsql=0 then%>checked<%end if%>> Access
					</TD>
				</TR>
				<Tr>
					<TD class="td_l_r title" width="100">数据库主机</TD>
					<TD class="td_l_l"><input name="item1" type="text" id="item1" value="<%if Data_Source<>"" then%><%=Data_Source%><%else%>(local)<%end if%>" class="setup_int" style=" width:150px;"> （例：LENOVO\SQLEXPRESS）</TD>
				</TR>
				<Tr>
					<TD class="td_l_r title">数据库名称</TD>
					<TD class="td_l_l "><input name="item2" type="text" id="item2" value="<%if Data_Catalog<>"" then%><%=Data_Catalog%><%end if%>" class="setup_int" style=" width:150px;"> （例：Easycrm）</TD>
				</TR>
				<Tr>
					<TD class="td_l_r title">数据库用户</TD>
					<TD class="td_l_l "><input name="item3" type="text" id="item3" value="<%if Data_User<>"" then%><%=Data_User%><%end if%>" class="setup_int" style=" width:150px;"> （例：sa）</TD>
				</TR>
				<Tr>
					<TD class="td_l_r title">数据库密码</TD>
					<TD class="td_l_l "><input name="item4" type="text" id="item4" value="<%if Data_Password<>"" then%><%=Data_Password%><%end if%>" class="setup_int" style=" width:150px;"></TD>
				</TR>
				<Tr>
					<TD class="td_l_r title">ACC数据库</TD>
					<TD class="td_l_l "><input name="item5" type="text" id="item5" value="<%if Data_MDBPath<>"" then%><%=Data_MDBPath%><%end if%>" class="setup_int" style=" width:150px;"></TD>
				</TR>
				<tr class="tr_f"> 
					<td class="td_l_l fontnobold" COLSPAN="2"><input type="Submit" class="button45" value="修改" />　<font color="#color:#CC0000">★</font> 注意：此操作有一定风险。</td>
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
	
	TempStr = TempStr & "'数据库连接" & VbCrLf
	TempStr = TempStr & "Accsql="&Accsql&" '数据库类型" & VbCrLf
	TempStr = TempStr & "Data_Source="& Chr(34) & Data_Source & Chr(34) &" 'MSSQL数据源（本机名\实例名 或 IP地址\实例名）" & VbCrLf
	TempStr = TempStr & "Data_Catalog="& Chr(34) & Data_Catalog & Chr(34) &" '数据库名称" & VbCrLf
	TempStr = TempStr & "Data_User="& Chr(34) & Data_User & Chr(34) &" '数据库用户" & VbCrLf
	TempStr = TempStr & "Data_Password="& Chr(34) & Data_Password & Chr(34) &" '数据库密码" & VbCrLf & VbCrLf
	TempStr = TempStr & "Data_MDBPath="& Chr(34) & Data_MDBPath & Chr(34) &" 'Access数据库路径" & VbCrLf & VbCrLf
	TempStr = TempStr & "SystemNumber="& Chr(34) & SystemNumber & Chr(34) &" '授权码" & VbCrLf & VbCrLf

	TempStr = TempStr & "%" & chr(62) & VbCrLf
	ADODB_SaveToFile TempStr,"../data/Mssql.asp"
	Response.Write("<script>alert(""修改成功！"");</script>")
	Response.Write "<script>location.href='?otype=Main&action=Main';</script>"
	end Sub
	
	Sub dbmanage()
        set rsSchema=conn.openSchema(20) 
        rsSchema.movefirst 
    %>
	<tr>
		<td colspan=2 class="Search_All td_n">
		 <form action="sql.asp?action=DBaddnew" method="post" onSubmit="return checkInput();">表名 <input type="text" id="refinput" name="crtablename" class="int" size="25" > <input type="submit" class="button245" value=" 建立新表 " id="submit">　<span class="info_help help01">新建的表默认创建一个ID字段,属性是数字型，递增，主键。</span></form>
		</td>
	</tr>
	<tr>
		<td valign="top" colspan=2 style="padding:0 10px 10px 10px;" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_c">表名</td>
					<td class="td_l_c" width="30%" colspan="4">操作</td>
					<td class="td_l_c" width="30%" colspan="4">表记录的SQL语句处理</td>
				</tr>
		<%
        Do Until rsSchema.EOF
            if rsSchema("TABLE_TYPE")="TABLE" then
		%>
				<tr class="tr"> 
					<form action="sql.asp?action=DBsave&tablename2=<%=rsSchema("TABLE_NAME")%>" method="post">
					<td class="td_r_l"><input type="text" name="tablename" value="<%=rsSchema("TABLE_NAME")%>" class="int" size="25"></td>
					<td class="td_r_c"><input type="submit" class="button227" value="保存"></td>
					</form>
					<td class="td_r_c"><a href="?action=DBdesign&otype=DBmanage&tablename=<%=rsSchema("TABLE_NAME")%>">设计表</a></td>
					<td class="td_r_c"><a href="?action=DBopen&otype=DBmanage&tablename=<%=rsSchema("TABLE_NAME")%>">打开表</a></td>
					<td class="td_r_c"><a onclick="checkclick('您确定要删除该表，包括里面的所有资料?')" href="?action=DBdel&tablename=<%=rsSchema("TABLE_NAME")%>">删除表</a></td>
					<td class="td_r_c"><a href="?action=DBsql&otype=DBsql&czsql=1&tablename=<%=rsSchema("TABLE_NAME")%>">查询</a></td>
					<td class="td_r_c"><a href="?action=DBsql&otype=DBsql&czsql=2&tablename=<%=rsSchema("TABLE_NAME")%>">插入</a></td>
					<td class="td_r_c"><a href="?action=DBsql&otype=DBsql&czsql=3&tablename=<%=rsSchema("TABLE_NAME")%>">更新</a></td>
					<td class="td_r_c"><a href="?action=DBsql&otype=DBsql&czsql=4&tablename=<%=rsSchema("TABLE_NAME")%>">删除</a></td>
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
		 <form action="sql.asp?action=DBaddnewfield&tablename=<%=trim(request.QueryString("tablename"))%>" method="post" id="form1" name="form1" onSubmit="return checkInput();">字段名 <input type="text" id="refinput" name="crfield" size="25" class="int" > 
			<select name="fieldtype" class="int" id="refselect">
				<option value="">字段类型</option>
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
			</select> <input type="submit" class="button245" value=" 新建字段 " id="1" name="1"> (设计表)　<B style="color:#f00;"><%=trim(request.QueryString("tablename"))%></B>　共<%=fieldCount%>个字段</form>
		</td>
	</tr>
	<tr>
		<td valign="top" colspan=2 style="padding:0 10px 10px 10px;" class="td_n">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_t"> 
					<td class="td_l_c">字段名称</td>
					<td class="td_l_c">字段类型</td>
					<td class="td_l_c">字段长度</td>
					<td class="td_l_c" colspan="2">操作</td>
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
					<td class="td_l_c" width="10%"><input type="submit" class="button227" value="保存"></td>
					<td class="td_l_c" width="10%"><a onclick="checkclick('您确定要删除该字段，包括里面的所有资料?')" href="sql.asp?action=DBdelfield&tablename=<%=trim(request.QueryString("tablename"))%>&fieldsname=<%=rs.Fields(i).Name%>">删除</a></td>
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
			<a href="?action=DBsql&czsql=1&tablename=<%=trim(request.QueryString("tablename"))%>">表记录查询</a> | <a href="?action=DBsql&czsql=2&tablename=<%=trim(request.QueryString("tablename"))%>">插入</a> | <a href="?action=DBsql&czsql=3&tablename=<%=trim(request.QueryString("tablename"))%>">更新</a> | <a href="?action=DBsql&czsql=4&tablename=<%=trim(request.QueryString("tablename"))%>">删除</a> | 　<span class="info_help help01">只显示前10条记录</span>
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
            Response.Write "请选择字段类型"
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
        fieldsname2=trim(request.Form("fieldsname2")) '原表名
        fieldtype=trim(request.Form("fieldtype"))
        crtable("sp_rename '"&tablename&"."&fieldsname2&"','"&fieldsname&"','column';") '字段名修改
        
        fieldssize=trim(request.Form("fieldssize"))
        fieldar=""
        select case fieldtype
        case "varchar","nvarchar"
            fieldar="("&fieldssize&")"
        end select
        if fieldssize=0 then fieldar="" end if
        crtable("ALTER TABLE ["&tablename&"] ALTER COLUMN ["&fieldsname&"] "&fieldtype&""&fieldar&"") '字段类型处理
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
					<td class="td_l_l" COLSPAN="6"><B>语句案例： </B></td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_r title">插入语句</td>
					<td class="td_l_l">insert into 表名(字段1,字段2)values('内容1','内容2')</td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_r title">更新语句</td>
					<td class="td_l_l">update 表名 set 字段1='内容1',字段2='内容2' where 字段3='内容3'</td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_r title">删除语句</td>
					<td class="td_l_l">delete from 表名 where 字段='内容'</td>
				</tr>
				<tr class="tr"> 
					<td class="td_l_r title">查询语句</td>
					<td class="td_l_l">select top 显示的记录数目 字段1,字段2 from 表名 where 字段1='内容1'</td>
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
						<input name="Submit" type="submit" class="button45" value="执行SQL"> 
						<input name="Rest" type="reset" class="button43" value="重写"> 
					</td> 
				</tr>
				</form>
				<tr class="tr_f"> 
					<td class="td_l_l fontnobold" COLSPAN="6"><font color="#color:#CC0000">★</font>注意：用select查询记录的时候加上top语句,如果记录过多,就会出现延时,打不开等现象,加上top就可以限止显示多少条查询的结果。</td>
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
				Response.Write("<script>alert('执行："&trim(request.Form("sqlstr"))&" 成功！');</script>")
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
					<td class="td_l_l" COLSPAN="6"><B>备份数据库： </B></td>
				</tr>
				<%if Accsql=1 then%>
				<tr> 
					<td class="td_l_r title">文件名及路径</td>
					<td class="td_l_l"><input type="text" size=70 name="dbpath" id="refinput" class="int" value="<%=server.MapPath("../data/bak")%>\<%=year(now())&month(now())&day(now())&hour(now())&minute(now())&second(now())%>.bak">　<span class="info_help help01">备份路径是：../data/bak/</span></td>
				</tr>
				<tr> 
					<td class="td_l_r title">找到的备份</td>
					<td class="td_l_l"><%=FileListBak("../data/bak","bak")%></td>
				</tr>
				<%else%>
				<tr> 
					<td class="td_l_r title">数据库真实路径</td>
					<td class="td_l_l"><input name=DBpath type=text id="DBpath" value="<%=Data_MDBPath%>" size="40" /></td>
				</tr>
				<tr> 
					<td class="td_l_r title">备份文件夹</td>
					<td class="td_l_l"><input name=bkfolder type=text value="../data/bak/" size="40" Readonly="true"/></td>
				</tr>
				<tr> 
					<td class="td_l_r title">备份数据库名称</td>
					<td class="td_l_l"><input name=bkDBname type=text value="<%=year(now())&month(now())&day(now())&hour(now())&minute(now())&second(now())%>.mdb" size="40" /></td>
				</tr>
				<tr> 
					<td class="td_l_r title">找到的备份</td>
					<td class="td_l_l"><%=FileListBak("../data/bak","bak")%></td>
				</tr>
				<%end if%>
				<tr> 
					<td class="td_l_l" colspan=2><input type="submit" class="button45" value="开始备份" id=submit1 name=submit1></td>
				</tr>
				<tr class="tr_f"> 
					<td class="td_l_l fontnobold" COLSPAN="6"><font color="#color:#CC0000">★</font> 注意：过大的数据库备份有可能超时，建议在访问量少的时候操作。</td>
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
				Response.Write ("<script>alert(""该命名的备份数据库已经存在，请删除或更换备份文件名进行备份"");history.back(-1);</script>")
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
				Response.Write ("<script>alert("""&err.number&"："&err.description&" "");history.back(-1);</script>")
                else
				Response.Write ("<script>alert(""数据库已经备份成功"");location.href='?otype=DBbak&action=DBbak';</script>")
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
		Response.Write ("<script>alert(""数据库已经备份成功"");location.href='?otype=DBbak&action=DBbak';</script>")
	End Sub
	
	'------------------检查某一目录是否存在-------------------
	Function CheckDir(FolderPath)
		dim fso1
		folderpath=Server.MapPath(".")&"\"&folderpath
		Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
		If fso1.FolderExists(FolderPath) then
		   '存在
		   CheckDir = True
		Else
		   '不存在
		   CheckDir = False
		End if
		Set fso1 = nothing
	End Function
	'-------------根据指定名称生成目录-----------------------
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
					<td class="td_l_l" COLSPAN="2"><B>压缩数据库： </B></td>
				</tr>
				<tr> 
					<td class="td_l_r title">数据库路径</td>
					<td class="td_l_l"><input name=dbpath type=text id="dbpath" value="<%=Data_MDBPath%>" size="40" /></td>
				</tr>
				<tr> 
					<td class="td_l_l" colspan=2><input type="submit" class="button45" value="开始压缩" id=submit1 name=submit1></td>
				</tr>
				<tr class="tr_f"> 
					<td class="td_l_l fontnobold" COLSPAN="6"><font color="#color:#CC0000">★</font> 注意：当前操作有一定风险，压缩前请确认已经备份。</td>
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
		Response.Write ("<script>alert(""数据库压缩成功"");location.href='?otype=DBCompress&action=DBCompress';</script>")
	End Sub
	
    Sub DBrestored()
    %>
	<tr>
		<td valign="top" class="td_n pd10">
			<form action="sql.asp?action=DBrestoredsave" method="post" onSubmit="return checkInput();">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="table_1">
				<tr class="tr_1"> 
					<td class="td_2" COLSPAN="6"><B>恢复数据库： </B></td>
				</tr>
				<tr class="tr"> 
					<td class="td_2 td_title">文件名及路径</td>
					<td class="td_2 td_content"><input type="text" size=70 name="dbpath" id="refinput" class="int" value="">　<span class="info_help help01">恢复路径是数据库服务器的路径</span></td>
				</tr>
				<tr class="tr"> 
					<td class="td_2 td_title">找到的备份</td>
					<td class="td_2 td_content"><%=FileList("../data/bak","bak")%>　<span class="info_help help01">选择后备份文件地址自动写入恢复路径</span></td>
				</tr>
				<tr class="tr"> 
					<td class="td_2 td_title"></td>
					<td class="td_2 td_content"><input type="submit" class="button45" value="开始恢复" id=submit1 name=submit1></td>
				</tr>
				<tr class="tr_f"> 
					<td class="td_l_l fontnobold" COLSPAN="6"><font color="#color:#CC0000">★</font> 注意：数据库正在使用时无法操作。</td>
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
				Response.Write ("<script>alert(""没有找到该备份文件,(注意路径为数据库服务器的路径)"");history.back(-1);</script>")
            Else
				Response.Write ("<script>alert(""正在从"&dbPath&"还原,请稍候"");</script>")
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
				Response.Write ("<script>alert("""&err.number&"："&err.description&" "");history.back(-1);</script>")
                else
				Response.Write ("<script>alert(""数据库恢复成功"");location.href='?otype=DBbak&action=DBbak';</script>")
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
    If Instr(1,cstring,"%")>0 or Instr(1,cstring,"=")>0 or Instr(1,cstring,"&")>0 or Instr(1,cstring,"#")>0 or Instr(1,cstring,">")>0 or Instr(1,cstring,"<")>0 or Instr(1,cstring,"'")>0 or Instr(1,cstring,";")>0 or Instr(1,cstring,"　")>0 or Instr(1,cstring,"`")>0 or Instr(1,cstring,"*")>0 or Instr(1,cstring,",")>0 then
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
	    Response.Write "数据库连接出错，请检查连接字串。"
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
			GBL_CHK_TempStr = "您的主机不支持ADODB.Stream，无法完成操作，请使用FTP等功能，将<font color=Red >data/config.asp</font>文件内容替换成框中内容"
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
        'Response.Write "执行:&nbsp;"&SqlCommand&"&nbsp;&nbsp;成功<br>"
		if action="DBaddnewfield" or action="DBsavefield" or action="DBdelfield" then
		Response.Write "<script>alert('执行："&SqlCommand&" 成功！');location.href='sql.asp?action=DBdesign&otype=DBmanage&tablename="&trim(request.QueryString("tablename"))&"';</script>"
		elseif action="DBaddnew" then
        Response.Write("<script>alert('执行："&SqlCommand&" 成功！');location.href='?otype=DBmanage&action=DBmanage';</script>")
		elseif action="DBsqlsub" then
        Response.Write("<script>alert('执行："&SqlCommand&" 成功！');history.back(-1);</script>")
		else
        Response.Write("<script>alert('执行："&SqlCommand&" 成功！');location.href='?otype=DBmanage&action=DBmanage';</script>")
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
    	FileList=FileList&"<option value="""">请选择</option>"
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
    FileListBak=FileListBak&"<a href='../data/bak\"&FileName.Name&"'>"&FileName.Name&"</a> <input type='button' class='button227' value='删除'  onClick=window.location.href='?otype=DBbak&action=DBbak&Subaction=DelDataBakFlie&DataBakFlie=../data/bak/"&FileName.Name&"' /> <br>"
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
	Response.Write ("<script>alert(""数据库删除成功"");location.href='?otype=DBbak&action=DBbak';</script>")
end if
%>
</table>

</body>
</html>
