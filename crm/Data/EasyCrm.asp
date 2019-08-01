<%
Class EasyCRM_CRM
	Private Sub Class_Initialize()
	End Sub 
 
	'组件支持
	Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
	End Function
 
	Function Mydb(MySqlstr,MyDBType)
		Select Case MyDBType
		Case 1 : Set Mydb = Conn.Execute(MySqlstr) : Dataquery = Dataquery + 1
		End Select
	End Function
 
	function getHTTPPage(url)
		dim Http
		set Http=server.createobject("MS"&"XML2.XML"&"HTTP")
		Http.open "GET",url,false
		Http.send()
		if Http.readystate<>4 then 
			exit function
		end if
		getHTTPPage=bytesToBSTR(Http.responseBody,"GB2312")
		set http=nothing
		if err.number<>0 then err.Clear 
	end function
 
	'转换代码，不然全是乱码
	Function BytesToBstr(body,Cset)
		dim objstream
		set objstream = Server.CreateObject("ado"&"db.str"&"eam")
		objstream.Type = 1
		objstream.Mode =3
		objstream.Open
		objstream.Write body
		objstream.Position = 0
		objstream.Type = 2
		objstream.Charset = Cset
		BytesToBstr = objstream.ReadText
		objstream.Close
		set objstream = nothing
	End Function
 
	'导出/导入数据过滤
	function clearWord(str)
		dim regEx
		set regEx=New RegExp
		regEx.IgnoreCase=True
		regEx.Global=True
		regEx.Pattern="<[^>]*>"
		str = regEx.replace(str,"" )
		regEx.Pattern="{[^}]*}"
		str = regEx.replace(str,"" )
		regEx.Pattern="/[^/]*/"
		str = regEx.replace(str,"" )
		str = Replace(str,"&nbsp;"," ")
		str = Replace(str,","," ")
		str = Replace(str,";"," ")
		str = Replace(str,"<"," ")
		str = Replace(str,">"," ")
		str = Replace(str,"#","＃")
		str = Replace(str,"$","￥")
		str = Replace(str,"%","％")
		str = Replace(str,"^","……")
		str = Replace(str,"&","＆")
		str = Replace(str,"(","（")
		str = Replace(str,")","）")
		str = Replace(str,"?","？")
		str = Replace(str,"[","【")
		str = Replace(str,"]","】")
		str = Replace(str," ","")
		str = Replace(str,"　","")
		str = Replace(str,chr(10)," ")
		str = Replace(str,chr(13)," ")
		clearWord= str 
		set regEx=nothing
	end function
 
	'过滤用户名	
	Function ReName(str)
		str = Replace(str,"'","")
		str = Replace(str,".","")
		str = Replace(str,",","")
		str = Replace(str,";","")
		str = Replace(str,"<","")
		str = Replace(str,">","")
		str = Replace(str,"!","")
		str = Replace(str,"@","")
		str = Replace(str,"#","")
		str = Replace(str,"$","")
		str = Replace(str,"%","")
		str = Replace(str,"^","")
		str = Replace(str,"&","")
		str = Replace(str,"*","")
		str = Replace(str,"(","")
		str = Replace(str,")","")
		str = Replace(str,"_","")
		str = Replace(str,"+","")
		str = Replace(str,"|","")
		str = Replace(str,"?","")
		str = Replace(str,"[","")
		str = Replace(str,"]","")
		str = Replace(str,"asp","")
		str = Replace(str,"asa","")
		str = Replace(str,"php","")
		str = Replace(str,"aspx","")
		str = Replace(str,"cer","")
		str = Replace(str,"cdx","")
		str = Replace(str,"htr","")
		str = Replace(str,chr(0),"")
		str = Replace(str,chr(10),"")
		str = Replace(str,chr(13),"")
		str = Replace(str," ","")
		ReName = str
	End Function
 
 
	'过滤代码
	Function htmlEncode2(str)
		 If IsEmpty(str) Or str = "" Then
		    htmlEncode2 = "&nbsp;"
		Else
		    str = Replace(str,">","&gt;")
			str = Replace(str,"<","&lt;")
			str = Replace(str,"'","&quot;")
			str = Replace(str,Chr(13),"<br>")
			str = Replace(str,VBCrlf,"<br>")
			str = Replace(str," ","&nbsp;")
			str = Replace(str,",",".")
			htmlEncode2 = str
		End If
	End Function
 
	Function htmlEncode3(str)
		If IsEmpty(str) Or str <> "" Then
		str = Replace(str,"&quot;","'")
		str = Replace(str,"<br>",Chr(13))
		str = Replace(str,"&nbsp;"," ")
		str = Replace(str,"&gt;",">")
		str = Replace(str,"&lt;","<")
		str = Replace(str,",",".")
		End If
		htmlEncode3 = str
	End Function
 
	'过滤搜索漏洞//or=or
	Function Searchcode(str)
		str = Replace(str,"'","")
		str = Replace(str,"or","")
		Searchcode = str
	End Function 
 
	'权限范围
	Function getUserList(intLevel,intGroup,inManagerange)
		Dim rs,strUserList
		Set rs = Server.CreateObject("ADODB.Recordset")
		if intLevel<>"" and intGroup<>"" then
			arrManagerange = Replace(Replace(inManagerange," ",""),",", "','")
			rs.Open "Select * From [user] where uName In ( Select uName From [user] Where uLevel < "&Session("CRM_level")&" And uGroup = "&Session("CRM_group")&" or uName in ( '"&arrManagerange&"' ) )",conn,1,1
		else
			Response.write"<script>location.href=""../main/login.asp"";</script>"
		end if
		strUserList = "'"&Session("CRM_name")&"'" '添加自己
		Do While Not rs.BOF And Not rs.EOF
			if rs("uName") <> Session("CRM_name") then  '跳过自己
				strUserList = strUserList & ",'" & rs("uName") & "'"
			end if
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		getUserList = strUserList
	End Function
 
	'下拉菜单
	Function getList(i,sTable,iId,sValue,sName,str)
		If i < 1 Or i > 2 Then
			getList = ""
			Exit Function
		End If
		Dim rs,strList
		strList = "<select name='" & sName & "' class='int'>"
		strList = strList & "<option value="""">"&l_Select&"</option>"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [" & sTable & "]",conn,1,1
		Do While Not rs.BOF And Not rs.EOF
			If i = 1 Then
				if rs(sValue) = ""&str&"" then
					strList = strList & "<option value="""&rs(sValue)&""" selected>"&rs(sValue)&"</option>"  '读取默认值
				else
					strList = strList & "<option value="""&rs(sValue)&""">"&rs(sValue)&"</option>" 
				end if
			Else
				if rs(sValue) = ""&str&"" then
					 strList = strList & "<option value="""&rs(iId)&""" selected>"&rs(sValue)&"</option>"  '读取默认值
				else
					strList = strList & "<option value="""&rs(iId)&""">"&rs(sValue)&"</option>" 
				end if
			End If
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		strList = strList & "</select>"
		getList = strList
	End Function
 
	'用户列表下拉菜单
	' i=1 权限内用户  i=2 所有用户（管理员）
	' sName 表单名
	Function UserList(i,sName,str)
		If i < 1 Or i > 2 Then
			getList = ""
			Exit Function
		End If
		Dim rs,strList
		strList = "<select name='" & sName & "' class='int'>"
		strList = strList & "<option value="""">"&l_Select&"</option>"
		Set rs = Server.CreateObject("ADODB.Recordset")
		If i = 1 Then
			arrManagerange = Replace(Replace(Session("CRM_MR")," ",""),",", "','")
			rs.Open "Select uName From [user] Where ( uLevel <= "&Session("CRM_level")&" And uGroup = "&Session("CRM_group")&" ) or uName in ( '"&arrManagerange&"' ) ",conn,1,1
		Else
			rs.Open "Select * From [user]",conn,1,1
		End If
		Do While Not rs.BOF And Not rs.EOF
			if rs("uName") = ""&str&"" then
				strList = strList & "<option value="""&rs("uName")&""" selected>"&rs("uName")&"</option>"  '读取默认值
			else
				strList = strList & "<option value="""&rs("uName")&""">"&rs("uName")&"</option>" 
			end if
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		strList = strList & "</select>"
		UserList = strList
	End Function
 
	'读取下拉框
	Function getSelect(sTable,sValue,sName,sInfo)
		Dim strSelect
		Dim rs
		strSelect = "<select name="""&sName&""" id="""&sValue&""">"
		strSelect = strSelect & "<option value="""">"&l_Select&"</option>"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [" & sTable & "] where "&sValue&"<>'' and "&sValue&"<>'Null' ",conn,1,1
		Do While Not rs.BOF And Not rs.EOF
			if rs(sValue) = ""&sInfo&"" then
				strSelect = strSelect & "<option value="""&rs(sValue)&""" selected>"&rs(sValue)&"</option>" 
			else
				strSelect = strSelect & "<option value="""&rs(sValue)&""">"&rs(sValue)&"</option>" 
			end if
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		strSelect = strSelect & "</select>"
		getSelect = strSelect
	End Function
 
	Function getNewSelect(sTable,sValue,sName,sql,sInfo)
		Dim strSelect
		Dim rs
		strSelect = "<select name="""&sName&""">"
		strSelect = strSelect & "<option value="""">"&l_Select&"</option>"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [" & sTable & "] where "&sValue&"<>'' and "&sValue&"<>'Null'"&sql,conn,1,1
		Do While Not rs.BOF And Not rs.EOF
			if rs(sValue) = ""&sInfo&"" then
				strSelect = strSelect & "<option value="""&rs(sValue)&""" selected>"&rs(sValue)&"</option>" 
			else
				strSelect = strSelect & "<option value="""&rs(sValue)&""">"&rs(sValue)&"</option>" 
			end if
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		strSelect = strSelect & "</select>"
		getNewSelect = strSelect
	End Function
 
	'读取按钮radio
	Function getRadio(sTable,sValue,sName,sInfo)
		Dim strRadio
		Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From [" & sTable & "] where "&sValue&"<>'' and "&sValue&"<>'Null' ",conn,1,1
		Do While Not rs.BOF And Not rs.EOF
			if rs(sValue) = ""&sInfo&"" then
				strRadio = strRadio & "<input name="""&sName&""" type=""radio"" value="""&rs(sValue)&""" checked>"&rs(sValue)&"　" 
			else
				strRadio = strRadio & "<input name="""&sName&""" type=""radio"" value="""&rs(sValue)&""">"&rs(sValue)&"　" 
			end if
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		getRadio = strRadio
	End Function
 
	'获取字段内容
	Function getNewItem(table,id,idstr,item)
		Dim rsNew
		Set rsNew = Server.CreateObject("ADODB.Recordset")
		rsNew.Open "Select * From ["&table&"] Where "&id&" = "&idstr&" ",conn,1,1
		If rsNew.RecordCount < 1 Then
			getNewItem = 0
		ElseIf rsNew.RecordCount > 1 Then
			getNewItem = rsNew.RecordCount
		Else
			getNewItem = rsNew(""&item&"")
		End If
		rsNew.Close
		Set rsNew = Nothing
	End Function
 
	'统计数量
	Function getCountItem(table,id,idstr,sql)
		Dim rsNew
		Set rsNew = Server.CreateObject("ADODB.Recordset")
		rsNew.Open "Select count("&id&") As "&idstr&" From ["&table&"] Where 1=1 "&sql&" " ,conn,1,1
		getCountItem = rsNew(idstr)
		rsNew.Close
		Set rsNew = Nothing
	End Function
 
	Function getSUMItem(table,id,idstr,sql)
		Dim rsNew
		Set rsNew = Server.CreateObject("ADODB.Recordset")
		rsNew.Open "Select sum("&id&") As "&idstr&" From ["&table&"] Where 1=1 "&sql&" " ,conn,1,1
		getSUMItem = rsNew(idstr)
		rsNew.Close
		Set rsNew = Nothing
	End Function 
 
	'统计-输出值-表-时间字段-时间开始-时间结束-用户字段-当前用户
	Function getCount(Item0,Item1,Item2,Item3,Item4,Item5,Item6,Item7,Item8,Item9,Item10,Item11)
		Dim rs,itemValue,sql
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = ""
		if Item4<>"" then
			sql = sql & " and "&Item3&" >= #" & Item4&"#"
		end if
		if Item5<>"" then
			sql = sql & " and "&Item3&" <= #" & Item5&"#"
		end if
		if Item6<>"" and Item7<>"" then
			sql = sql & " and "&Item6&" = " & Item7&""
		end if
		if Item8<>"" and Item9<>"" then
			sql = sql & " and "&Item8&" = " & Item9&""
		end if
		if Item10<>"" and Item11<>"" then
			sql = sql & " and "&Item10&" = " & Item11&""
		end if
		rs.Open "Select count("&Item0&") As "&Item1&" From "&Item2&" Where 1=1 "&sql&" " ,conn,1,1
	    itemValue = rs(Item1)
		rs.Close
		Set rs = Nothing
		getCount = itemValue
	End Function
 
	'模糊查询
	' otypestr = 查询字段
	' keystr = 查询关键字
	function seachKey(otypestr,keystr) 
		dim tmpstr,MyArray,I
		MyArray = Split(keystr) '默认以空格分组
		For I = Lbound(MyArray) to Ubound(MyArray)
			if I=0 then
				tmpstr=tmpstr & " and "&otypestr&" like '%"&MyArray(I)&"%'"
			else
				tmpstr=tmpstr & " and "&otypestr&" like '%"&MyArray(I)&"%'"
			end if
		Next
		seachKey=tmpstr
	end function
 
	function urldecode(encodestr) 
		newstr="" 
		havechar=false 
		lastchar="" 
		for i=1 to len(encodestr) 
			char_c=mid(encodestr,i,1) 
			if char_c="+" then 
				newstr=newstr & " " 
			elseif char_c="%" then 
				next_1_c=mid(encodestr,i+1,2) 
				next_1_num=cint("&H" & next_1_c) 
				if havechar then 
					havechar=false 
					newstr=newstr & chr(cint("&H" & lastchar & next_1_c)) 
				else 
					if abs(next_1_num)<=127 then 
						newstr=newstr & chr(next_1_num) 
					else 
						havechar=true 
						lastchar=next_1_c 
					end if 
				end if 
				i=i+2 
			else 
				newstr=newstr & char_c 
			end if 
		next 
		urldecode=newstr 
	end Function
	
	Function showsize(filename) 
		FPath=server.mappath(filename) 
		set fso=server.CreateObject("scripting.FileSystemObject")
		If fso.fileExists(FPath) Then 
			Set f = fso.GetFile(FPath) 
			filetype=f.type 
			filesize=f.size 
			adddate=f.DateCreated 
			showsize = filesize
		end if 
	End Function 
 
	'时间格式化
	Function FormatDate(DateAndTime, para)
		On Error Resume Next
		Dim y, m, d, h, mi, s, strDateTime
		FormatDate = ""
		If IsNull(DateAndTime)=false and IsEmpty(DateAndTime)=false Then
			If IsNumeric(para) and IsDate(DateAndTime) Then
				y = CStr(Year(DateAndTime))
				m = CStr(Month(DateAndTime))
				If Len(m) = 1 Then m = "0" & m
				d = CStr(Day(DateAndTime))
				If Len(d) = 1 Then d = "0" & d
				h = CStr(Hour(DateAndTime))
				If Len(h) = 1 Then h = "0" & h
				mi = CStr(Minute(DateAndTime))
				If Len(mi) = 1 Then mi = "0" & mi
				s = CStr(Second(DateAndTime))
				If Len(s) = 1 Then s = "0" & s
				Select Case para
				Case "1"
					strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
				Case "2"
					strDateTime = y & "-" & m & "-" & d
				Case "3"
					strDateTime = y & "/" & m & "/" & d
				Case "4"
					strDateTime = y & "年" & m & "月" & d & "日"
				Case "5"
					strDateTime = y & "年" & m & "月"
				Case "6"
					strDateTime = m & "月" & d & "日"
				Case "7"
					strDateTime = y & "-" & m
				Case "8"
					strDateTime = m & "-" & d
				Case "9"
					strDateTime = y & "/" & m
				Case "10"
					strDateTime = m & "/" & d
				Case "11"
					strDateTime = d & "日"
				Case "12"
					strDateTime = h & ":" & mi
				Case "14"
					strDateTime = y & m & d & h & mi & s
				Case Else
					strDateTime = DateAndTime
				End Select
				FormatDate = strDateTime
			End if
		End if
	End Function
 
	'翻页类
	Function pagelist(baseURL, PN, TotalPages,TotalRecords)
		if baseURL = "" or isNull(baseURL) then exit function
		if inStr(baseURL, "?") and right(baseURL, 1) <> "?" and right(baseURL, 1) <> "&" then
			baseURL = baseURL & "&"
		else
			baseURL = baseURL & "?"
		end if
		Dim strList
	    strList = "<div class=""pager"">"
		if PN < 2 then
			strList = strList & "<a class='backpage grey'>首页</a>"
		else
			strList = strList & "<a href='"&baseURL&"PN=1' class='backpage'>首页</a>"
		end if
		if PN>3 then s1=PN-3 else s1=1
		if PN < TotalPages-3 then s2=PN+3 else s2 = TotalPages
		for j=s1 to s2
			if j=PN then
				strList = strList & "<a href='javascript:;' class='active'><b>"&j&"</b></a>"
			else
				strList = strList & "<a href='"&baseURL&"PN="&j&"'><b>"&j&"</b></a>"
			end if
		next
		If PN = TotalPages then
			strList = strList & "<a class='nextpage greys'>尾页</a>"
		Else
			strList = strList & "<a href='"&baseURL&"PN="&TotalPages&"' class='nextpage'>尾页</a>"
		End If
		strList = strList & "　<label style='vertical-align: middle;'>"
		strList = strList & "记录：<B style=""color:#f30;""> "&TotalRecords&" </B>条</label>　"
		strList = strList & "<script language='javascript'>function getValue(obj){if (document.getElementById(""geturl"").value!="""") {location.href="""&baseURL&"PN=""+escape(document.getElementById(""geturl"").value)+""""}}</script>"
		strList = strList & "<input class=""int"" style=""width:30px;vertical-align:middle;"" type=""text"" id=""geturl"" value=""" &PN& """ /> / <B style=""color:#f30;"">"&TotalPages&"</B> 页 <input class=""button_page_go"" type=""button"" value="" "" style=""vertical-align:middle;"" onclick=""getValue(geturl.value)"" /> "
		strList = strList & "</div>"
		pagelist = strList
	End Function
 
End Class
%>