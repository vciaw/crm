<%
'=========================================================
 '类名: AnUpLoad(艾恩无组件上传类)
 '作者: Anlige
 '版本: 艾恩ASP无组件上传类优化版(V9.11.1)
 '开发日期: 2008-4-12
 '修改日期: 20010-5-24
 '主页: http://dev.mo.cn
 'Email: i@ruboy.com
 'QQ: 1034555083
'=========================================================
Class AnUpLoad
	Private Form, Fils
	Private vCharSet, vMaxSize, vSingleSize, vErr, vVersion, vTotalSize
	Private vExe, pID, vOP, vErrExe,vboundary, vLostTime, vMode, vFileCount
	Private vIsNum ,vNum
	'==============================
	'设置和读取属性开始
	'==============================
	Public Property Let IsNum(ByVal value)
		vIsNum = value
	End Property
	
	Public Property Let IsNumIng(ByVal value)
		vNum = value
	End Property
	
	Public Property Let Mode(ByVal value)
		vMode = value
	End Property
	
	Public Property Let MaxSize(ByVal value)
		vMaxSize = value
	End Property
	
	Public Property Let SingleSize(ByVal value)
		vSingleSize = value
	End Property
	
	Public Property Let Exe(ByVal value)
		vExe = LCase(value)
	End Property
	
	Public Property Let CharSet(ByVal value)
		vCharSet = value
	End Property
	
	Public Property Get ErrorID()
		ErrorID = vErr
	End Property
	
	Public Property Get FileCount()
		FileCount = Fils.count
	End Property
	
	Public Property Get Description()
		Description = GetErr(vErr)
	End Property
	
	Public Property Get Version()
		Version = vVersion
	End Property
	
	Public Property Get TotalSize()
		TotalSize = vTotalSize
	End Property
	
	Public Property Get ProcessID()
		ProcessID = pID
	End Property
	
	Public Property Let openProcesser(ByVal value)
		vOP = value
	End Property
	
	Public Property Get LostTime()
		LostTime = vLostTime
	End Property
	'==============================
	'设置和读取属性结束，初始化类
	'==============================
	
	Private Sub Class_Initialize()
		set Form = server.createobject("Scripting.Dictionary")
		set Fils = server.createobject("Scripting.Dictionary")
		vVersion = "艾恩ASP无组件上传类优化版(V9.11.1)"
		vMaxSize = -1
		vSingleSize = -1
		vErr = -1
		vExe = ""
		vTotalSize = 0
		vCharSet = "gb2312"
		vOP=false
		pID="AnUpload"
		setApp "",0,0,""
		vMode = 0
		vIsNum = 1
		vNum = 18
	End Sub
	
	Private Sub Class_Terminate()
		Dim f
		Form.RemoveAll()
		For each f in Fils 
			Fils(f).value=empty
			Set Fils(f) = Nothing
		Next
		Fils.RemoveAll()
		Set Form = Nothing
		Set Fils = Nothing
	End Sub
	
	'==============================
	'函数名:GetData
	'作用:处理客户端提交来的所有数据
	'==============================
	Public Sub GetData()
		Dim time1
		time1 = timer()
		if vOP then pID=request.querystring("processid")
		Dim value, str, bcrlf, fpos, sSplit, slen, istart,ef
		Dim TotalBytes,tempdata,BytesRead,ChunkReadSize,PartSize,DataPart,formend, formhead, startpos, endpos, formname, FileName, fileExe, valueend, NewName,localname,type_1,contentType
		TotalBytes = Request.TotalBytes
		ef = false
		If checkEntryType = false Then ef = true : vErr = 2
		'下面3句注释掉了，因为在IIS5.0中，如果上传大小大于限制大小的文件，会出错，一直没找到解决方法。如果是在IIS5以上的版本使用，可以取消下面3句的注释
		'If Not ef Then
			'If vMaxSize > 0 And TotalBytes > vMaxSize Then ef = true : vErr = 1
		'End If
		If ef Then Exit Sub
		If vMode = 0 Then
			vTotalSize = 0
			Dim StreamT 
			Set StreamT = server.CreateObject("Ado"&"db.str"&"eam")
			StreamT.Type = 1
			StreamT.Mode = 3
			StreamT.Open
			BytesRead = 0
			ChunkReadSize = 1024 * 16
			Do While BytesRead < TotalBytes
				PartSize = ChunkReadSize
				If PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
				DataPart = Request.BinaryRead(PartSize)
				StreamT.Write DataPart
				BytesRead = BytesRead + PartSize
				setApp "uploading",TotalBytes,BytesRead,""
			Loop
			setApp "uploaded",TotalBytes,BytesRead,""
			StreamT.Position = 0
			tempdata = StreamT.Read
			StreamT.Close()
			Set StreamT = Nothing
		Else
			tempdata = Request.BinaryRead(TotalBytes)
		End If
		bcrlf = ChrB(13) & ChrB(10)
		fpos = InStrB(1, tempdata, bcrlf)
        sSplit = MidB(tempdata, 1, fpos - 1)
		slen = LenB(sSplit)
		istart = slen + 2
		Do While lenb(tempdata) > 2 + slen
			formend = InStrB(istart, tempdata, bcrlf & bcrlf)
			formhead = MidB(tempdata, istart, formend - istart)
			str = Bytes2Str(formhead)
			startpos = InStr(str, "name=""") + 6
			endpos = InStr(startpos, str, """")
			formname = LCase(Mid(str, startpos, endpos - startpos))
			valueend = InStrB(formend + 3, tempdata, sSplit)
			If InStr(str, "filename=""") > 0 Then
				startpos = InStr(str, "filename=""") + 10
				endpos = InStr(startpos, str, """")
				type_1=instr(endpos,lcase(str),"content-type")
				contentType=trim(mid(str,type_1+13))
				FileName = Mid(str, startpos, endpos - startpos)
				If Trim(FileName) <> "" Then
					LocalName = FileName
					FileName = Replace(FileName, "/", "\")
					FileName = Mid(FileName, InStrRev(FileName, "\") + 1)
					FileName = Replace(FileName,chr(0),"")
					If instr(FileName,".")>0 Then
						fileExe = Split(FileName, ".")(UBound(Split(FileName, ".")))
					else
						fileExe = ""
					End If
					If vExe <> "" Then '判断扩展名
						If checkExe(fileExe) = True Then
							vErr = 3
							vErrExe = fileExe
							tempdata = empty
							Exit Sub
						End If
					End If
					NewName = Getname()
					NewName = NewName & "." & fileExe
					vTotalSize = vTotalSize + valueend - formend - 6
					If vSingleSize > 0 And (valueend - formend - 6) > vSingleSize Then '判断上传单个文件大小
						vErr = 5
						tempdata = empty
						Exit Sub
					End If
					If vMaxSize > 0 And vTotalSize > vMaxSize Then '判断上传数据总大小
						vErr = 1
						tempdata = empty
						Exit Sub
					End If
					If Fils.Exists(formname) Then
						vErr = 4
						tempdata = empty
						Exit Sub
					Else
						Dim fileCls:set fileCls=getNewFileObj()
						fileCls.ContentType=contentType
						fileCls.Size = (valueend - formend - 5)
						fileCls.FormName = formname
						fileCls.NewName = NewName
						fileCls.FileName = FileName
						fileCls.LocalName = FileName
						fileCls.extend=split(NewName,".")(ubound(split(NewName,".")))
						fileCls.value =midb(tempdata,formend + 4,valueend - formend - 5)
						Fils.Add formname, fileCls
						Set fileCls = Nothing
					End If
				End If
			Else
				value = MidB(tempdata, formend + 4, valueend - formend - 6)
				If Form.Exists(formname) Then
					Form(formname) = Form(formname) & "," & Bytes2Str(value)
				Else
					Form.Add formname, Bytes2Str(value)
				End If
			End If
			istart = 2 + slen
			tempdata = midb(tempdata,valueend+2)
		Loop
		vErr = 0
		tempdata = empty
		vLostTime = FormatNumber((timer-time1)*1000,2)
	End Sub
	
	Public sub setApp(stp,total,current,desc)
		Application.lock()
		Application(pID)="{ID:""" & pID & """,step:""" & stp & """,total:" & total & ",now:" & current & ",description:""" & desc & """,dt:""" & now() & """}"
		Application.unlock()
	end sub
	'==============================
	'判断扩展名
	'==============================
	Private Function checkExe(ByVal ex)
		Dim notIn: notIn = True
		If vExe="*" then
			notIn=false 
		elseIf InStr(1, vExe, "|") > 0 Then
			Dim tempExe: tempExe = Split(vExe, "|")
			Dim I: I = 0
			For I = 0 To UBound(tempExe)
				If LCase(ex) = tempExe(I) Then
					notIn = False
					Exit For
				End If
			Next
		Else
			If vExe = LCase(ex) Then
				notIn = False
			End If
		End If
		checkExe = notIn
	End Function
	
	'==============================
	'把数字转换为文件大小显示方式
	'==============================
	Public Function GetSize(ByVal iSize)
		Dim sRet,KB,MB,S
		KB = 1024 : MB = KB * KB
		If Not IsNumeric(iSize) Then
			GetSize = "未知"
			Exit Function
		End If
		If iSize < KB Then
			sRet = iSize & " Bytes"
		Else
			S = iSize / KB
			If S < 10 Then
				sRet = FormatNumber(iSize / KB, 2, -1) & " KB"
			ElseIf S < 100 Then
				sRet = FormatNumber(iSize / KB, 1, -1) & " KB"
			ElseIf S < 1000 Then
				sRet = FormatNumber(iSize / KB, 0, -1) & " KB"
			ElseIf S < 10000 Then
				sRet = FormatNumber(iSize / MB, 2, -1) & " MB"
			ElseIf S < 100000 Then
				sRet = FormatNumber(iSize / MB, 1, -1) & " MB"
			ElseIf S < 1000000 Then
				sRet = FormatNumber(iSize / MB, 0, -1) & " MB"
			ElseIf S < 10000000 Then
				sRet = FormatNumber(iSize / MB / KB, 2, -1) & " GB"
			Else
				sRet = FormatNumber(iSize / MB / KB, 1, -1) & " GB"
			End If
		End If
		GetSize = sRet
	End Function
	
	'==============================
	'二进制数据转换为字符
	'==============================
	Private Function Bytes2Str(ByVal byt)
		If LenB(byt) = 0 Then
			Bytes2Str = ""
			Exit Function
		End If
		Dim mystream, bstr
		Set mystream =server.createobject("ADO"&"DB.Str"&"eam")
		mystream.Type = 2
		mystream.Mode = 3
		mystream.Open
		mystream.WriteText byt
		mystream.Position = 0
		mystream.CharSet = vCharSet
		mystream.Position = 2
		bstr = mystream.ReadText()
		mystream.Close
		Set mystream = Nothing
		Bytes2Str = bstr
	End Function
	'==============================
	'弹出提示信息框
	'==============================
	Private Function goStr(oMsg)
		Dim outStr
		outStr = ""
		If oMsg = "" Or IsNull(oMsg) Then
			goStr = outStr
		Else
			outStr = outStr & "<script language=""javascript"" type=""text/javascript"" charset=""gb2312"">" & vbcrlf
			outStr = outStr & "alert('"&oMsg&"');" & vbcrlf
			outStr = outStr & "history.go(-1);" & vbcrlf
			outStr = outStr & "</script>" & vbcrlf
		End If
		goStr = outStr
	End Function
	'==============================
	'获取错误描述
	'==============================
	Private Function GetErr(ByVal Num)
		Select Case Num
			Case 0
				GetErr = goStr("数据处理完毕!")
			Case 1
				GetErr = goStr("上传数据超过" & GetSize(vMaxSize) & "限制!可设置MaxSize属性来改变限制!")
			Case 2
				GetErr = goStr("未设置上传表单enctype属性为multipart/form-data或者未设置method属性为Post,上传无效!")
			Case 3
				GetErr = goStr("含有非法扩展名(" & vErrExe & ")文件!只能上传扩展名为" & Replace(vExe, "|", ",") & "的文件")
			Case 4
				GetErr = goStr("对不起,程序不允许使用相同name属性的文件域!")
			Case 5
				GetErr = goStr("单个文件大小超出" & GetSize(vSingleSize) & "的上传限制!")
		End Select
	End Function
	'==============================
	'函数名：NumRand
	'作  用：生成n位随机数字
	'==============================
	Private Function NumRand(n)
		For i = 1 To n
			Randomize
			temp = CInt(9 * Rnd)
			temp = temp + 48
			NumRand = NumRand & Chr(temp)
		Next
	End Function
	'==============================
	'函数名：NumRand
	'作  用：生成n位随机小写字母
	'==============================
	Private Function LCharRand(n)
		For i = 1 To n
			Randomize
			temp = CInt(25 * Rnd)
			temp = temp + 97
			LCharRand = LCharRand & Chr(temp)
		Next
	End Function
	'==============================
	'函数名：NumRand
	'作  用：生成n位随机大写字母
	'==============================
	Private Function UCharRand(n)
		For i = 1 To n
			Randomize
			temp = CInt(25 * Rnd)
			temp = temp + 65
			UCharRand = UCharRand & Chr(temp)
		Next
	End Function
	'==============================
	'函数名：NumRand
	'作  用：生成n位随机数字大写字母组合
	'==============================
	Private Function allRand(n)
		For i = 1 To n
			Randomize
			temp = CInt(25 * Rnd)
			If temp Mod 2 = 0 Then
				temp = temp + 97
			ElseIf temp < 9 Then
				temp = temp + 48
			Else
				temp = temp + 65
			End If
			allRand = allRand & Chr(temp)
		Next
	End Function
	'==============================
	'函数名：NumRand
	'作  用：生成n位随机数字小写字母组合
	'==============================
	Private Function LallRand(n)
		For i = 1 To n
			Randomize
			temp = CInt(25 * Rnd)
			If temp Mod 2 = 0 Then
				temp = temp + 97
			ElseIf temp < 9 Then
				temp = temp + 48
			Else
				temp = temp + 97
			End If
			LallRand = LallRand & Chr(temp)
		Next
	End Function
	'==============================
	'根据日期生成随机文件名
	'==============================
	Private Function GetNameIng()
		Dim y, m, d, h, mm, S, r
		Randomize
		y = Year(Now)
		m = Month(Now): If m < 10 Then m = "0" & m
		d = Day(Now): If d < 10 Then d = "0" & d
		h = Hour(Now): If h < 10 Then h = "0" & h
		mm = Minute(Now): If mm < 10 Then mm = "0" & mm
		S = Second(Now): If S < 10 Then S = "0" & S
		r = NumRand(10)
		GetNameIng = y & m & d & h & mm & S & r
	End Function
	Public Function NumIng()
		Select Case vIsNum
			Case 1:
				NumIng = NumRand(vNum)
			Case 2:
				NumIng = LCharRand(vNum)
			Case 3:
				NumIng = UCharRand(vNum)
			Case 4:
				NumIng = allRand(vNum)
			Case 5:
				NumIng = LallRand(vNum)
			Case 6:
				NumIng = GetNameIng()
			Case Else
				NumIng = NumRand(vNum)
		End Select
	End Function
	'==============================
	'根据选择的参数生成随机文件名
	'==============================
	Private Function Getname()
		Getname = NumIng()
	End Function
	'==============================
	'检测上传类型是否为multipart/form-data
	'==============================
	Private Function checkEntryType()
		Dim ContentType, ctArray, bArray,RequestMethod
		RequestMethod=trim(LCase(Request.ServerVariables("REQUEST_METHOD")))
		if RequestMethod="" or RequestMethod<>"post" then
			checkEntryType = False
			exit function
		end if
		ContentType = LCase(Request.ServerVariables("HTTP_CONTENT_TYPE"))
		ctArray = Split(ContentType, ";")
		if ubound(ctarray)>=0 then
			If Trim(ctArray(0)) = "multipart/form-data" Then
			checkEntryType = True
			vboundary = Split(ContentType,"boundary=")(1)
			Else
			checkEntryType = False
			End If
		else
			checkEntryType = False
		end if
	End Function
	
	'==============================
	'获取上传表单值,参数可选,如果为-1则返回一个包含所有表单项的一个dictionary对象
	'==============================
	Public Function Forms(ByVal formname)
		If trim(formname) = "-1" Then
			Set Forms = Form
		Else
			If Form.Exists(LCase(formname)) Then
				Forms = Form(LCase(formname))
			Else
				Forms = ""
			End If
		End If
	End Function
	
	'==============================
	'获取上传的文件类,参数可选,如果为-1则返回一个包含所有上传文件类的一个dictionary对象
	'==============================
	Public Function Files(ByVal formname)
		If trim(formname) = "-1" Then
			Set Files = Fils
		Else
			If Fils.Exists(LCase(formname)) Then
				Set Files = Fils(LCase(formname))
			Else
				Set Files = Nothing
			End If
		End If
	End Function
End Class
%>
<script language="jscript" runat="server">
function getNewFileObj(){
	return new UploadFile();	
}
function UploadFile(){
	this.FormName="";
	this.NewName = "";
	this.LocalName="";
	this.FileName="";
	this.UserSetName="";
	this.ContentType="";
	this.Size=0;
	this.value=null;
	this.Path = "";
	this.extend="";
}

//保存文件的方法
UploadFile.prototype.SaveToFile=function(){
	var arg = arguments;
	var Path ,Option, OverWrite
	if(arg.length==0){return {error:true,description:'参数错误,请传递至少一个参数'};}
	if(arg.length==1){Path = arg[0];Option=0;OverWrite=true;}
	if(arg.length==2){Path = arg[0];Option=arg[1];OverWrite=true;}
	if(arg.length==3){Path = arg[0];Option=arg[1];OverWrite=arg[2];}
	if(arg.length>3){return {error:true,description:'参数错误,最多传递3个参数'};}
	try{
		var IsP = (Path.indexOf(":")==1)
		if(!IsP){
			Path = server.MapPath(Path);	
		}
		Path = Path.replace("/","\\");
		if(Path.substr(Path.length-1,1)!="\\"){Path = Path + "\\";}
		this.CreateFolder(Path);
		this.Path = Path;
		if(Option==1){
			Path = Path + this.LocalName;this.FileName = this.LocalName;
		}else{
			if(Option==-1 && this.UserSetName!=""){
				Path = Path + this.UserSetName + "." + this.extend;this.FileName = this.UserSetName + "." + this.extend;
			}else{
				Path = Path + this.NewName;this.FileName = this.NewName;
			}
		}
		if(!OverWrite){
			Path = this.GetFilePath();
		}
		var tmpStrm;
		var s1="str"
		var s2="eam"
		tmpStrm = server.CreateObject("adodb."+s1+s2);
		tmpStrm.mode=3;
		tmpStrm.type= 1;
		tmpStrm.open();
		var Info = server.CreateObject("ADODB.Recordset");
		Info.Fields.Append("value", 205,-1);
		Info.open();
		Info.addNew();
		Info("value").appendChunk(this.value);
		tmpStrm.write(Info("value"));
		Info("value").appendChunk(null);
		Info.update();
		Info.Close();
		Info = null;
		Path = Path.replace(/\u0000/igm,"");
		tmpStrm.saveToFile(Path,2);
		tmpStrm.close();
		tmpStrm = null;
		return {error:false,description:'成功保存文件'};
	}catch(ex){
		return {error:true,description:ex.description};
	}
};
//获取二进制数据的方法
UploadFile.prototype.GetBytes=function(){
	return this.value
};

UploadFile.prototype.CreateFolder=function(folderPath){
	var oFSO;
	oFSO =server.CreateObject("Scripting.FileSystemObject" );
	var sParent;
	sParent = oFSO.GetParentFolderName( folderPath );
	if(sParent == ""){return;}
	if(!oFSO.FolderExists(sParent)){this.CreateFolder( sParent );}
	if(!oFSO.FolderExists(folderPath)){oFSO.CreateFolder(folderPath);}
	oFSO = null;
};

UploadFile.prototype.GetFilePath=function(){
	var oFSO,Fname,FNameL,i=0;
	oFSO =server.CreateObject("Scripting.FileSystemObject" );
	Fname = this.Path + this.FileName;
	FNameL = this.FileName.substr(0,this.FileName.lastIndexOf("."));
	while(oFSO.FileExists(Fname)){
		Fname = this.Path + FNameL + "(" + i + ")." + this.extend;
		this.FileName = FNameL + "(" + i + ")." + this.extend
		i++;
	}
	oFSO = null;
	return Fname;
};
</script>