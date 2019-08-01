<%
'=========================================================
 '����: AnUpLoad(����������ϴ���)
 '����: Anlige
 '�汾: ����ASP������ϴ����Ż���(V9.11.1)
 '��������: 2008-4-12
 '�޸�����: 20010-5-24
 '��ҳ: http://dev.mo.cn
 'Email: i@ruboy.com
 'QQ: 1034555083
'=========================================================
Class AnUpLoad
	Private Form, Fils
	Private vCharSet, vMaxSize, vSingleSize, vErr, vVersion, vTotalSize
	Private vExe, pID, vOP, vErrExe,vboundary, vLostTime, vMode, vFileCount
	Private vIsNum ,vNum
	'==============================
	'���úͶ�ȡ���Կ�ʼ
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
	'���úͶ�ȡ���Խ�������ʼ����
	'==============================
	
	Private Sub Class_Initialize()
		set Form = server.createobject("Scripting.Dictionary")
		set Fils = server.createobject("Scripting.Dictionary")
		vVersion = "����ASP������ϴ����Ż���(V9.11.1)"
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
	'������:GetData
	'����:����ͻ����ύ������������
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
		'����3��ע�͵��ˣ���Ϊ��IIS5.0�У�����ϴ���С�������ƴ�С���ļ��������һֱû�ҵ�����������������IIS5���ϵİ汾ʹ�ã�����ȡ������3���ע��
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
					If vExe <> "" Then '�ж���չ��
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
					If vSingleSize > 0 And (valueend - formend - 6) > vSingleSize Then '�ж��ϴ������ļ���С
						vErr = 5
						tempdata = empty
						Exit Sub
					End If
					If vMaxSize > 0 And vTotalSize > vMaxSize Then '�ж��ϴ������ܴ�С
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
	'�ж���չ��
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
	'������ת��Ϊ�ļ���С��ʾ��ʽ
	'==============================
	Public Function GetSize(ByVal iSize)
		Dim sRet,KB,MB,S
		KB = 1024 : MB = KB * KB
		If Not IsNumeric(iSize) Then
			GetSize = "δ֪"
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
	'����������ת��Ϊ�ַ�
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
	'������ʾ��Ϣ��
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
	'��ȡ��������
	'==============================
	Private Function GetErr(ByVal Num)
		Select Case Num
			Case 0
				GetErr = goStr("���ݴ������!")
			Case 1
				GetErr = goStr("�ϴ����ݳ���" & GetSize(vMaxSize) & "����!������MaxSize�������ı�����!")
			Case 2
				GetErr = goStr("δ�����ϴ���enctype����Ϊmultipart/form-data����δ����method����ΪPost,�ϴ���Ч!")
			Case 3
				GetErr = goStr("���зǷ���չ��(" & vErrExe & ")�ļ�!ֻ���ϴ���չ��Ϊ" & Replace(vExe, "|", ",") & "���ļ�")
			Case 4
				GetErr = goStr("�Բ���,��������ʹ����ͬname���Ե��ļ���!")
			Case 5
				GetErr = goStr("�����ļ���С����" & GetSize(vSingleSize) & "���ϴ�����!")
		End Select
	End Function
	'==============================
	'��������NumRand
	'��  �ã�����nλ�������
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
	'��������NumRand
	'��  �ã�����nλ���Сд��ĸ
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
	'��������NumRand
	'��  �ã�����nλ�����д��ĸ
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
	'��������NumRand
	'��  �ã�����nλ������ִ�д��ĸ���
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
	'��������NumRand
	'��  �ã�����nλ�������Сд��ĸ���
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
	'����������������ļ���
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
	'����ѡ��Ĳ�����������ļ���
	'==============================
	Private Function Getname()
		Getname = NumIng()
	End Function
	'==============================
	'����ϴ������Ƿ�Ϊmultipart/form-data
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
	'��ȡ�ϴ���ֵ,������ѡ,���Ϊ-1�򷵻�һ���������б����һ��dictionary����
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
	'��ȡ�ϴ����ļ���,������ѡ,���Ϊ-1�򷵻�һ�����������ϴ��ļ����һ��dictionary����
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

//�����ļ��ķ���
UploadFile.prototype.SaveToFile=function(){
	var arg = arguments;
	var Path ,Option, OverWrite
	if(arg.length==0){return {error:true,description:'��������,�봫������һ������'};}
	if(arg.length==1){Path = arg[0];Option=0;OverWrite=true;}
	if(arg.length==2){Path = arg[0];Option=arg[1];OverWrite=true;}
	if(arg.length==3){Path = arg[0];Option=arg[1];OverWrite=arg[2];}
	if(arg.length>3){return {error:true,description:'��������,��ഫ��3������'};}
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
		return {error:false,description:'�ɹ������ļ�'};
	}catch(ex){
		return {error:true,description:ex.description};
	}
};
//��ȡ���������ݵķ���
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