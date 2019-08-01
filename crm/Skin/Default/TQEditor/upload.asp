<!--#include file="UpLoad_Class.asp"-->
<%
'------------------------------------------------------
Dim oUpLoadType, oAction, oFileExe, ooFileSize, oFileSize, UpLoadFile
'------------------------------------------------------
Response.Buffer = True
Response.ExpiresAbsolute = Now() -1
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.Charset = "GB2312"
'------------------------------------------------------
oUpLoadType = Trim(Request("oUpLoadType"))
'------------------------------------------------------
Select Case oUpLoadType
    '------------------------------------------- 图片
    Case "Images"
        oFileExe = "jpg|jpeg|gif|bmp|swf|png"
        sFileSize = 2048
        UpLoadFile = oCreateFolder("../../../UpLoad/"&year(now())&"/Images/")
	'------------------------------------------- 图片
    Case "Img"
        oFileExe = "jpg|jpeg|gif|bmp|png"
        sFileSize = 2048
        UpLoadFile = oCreateFolder("/UpLoadPic/Img/")
    '------------------------------------------- 广告图片
    Case "Ad"
        oFileExe = "jpg|jpeg|gif|bmp|swf|png"
        sFileSize = 500
        UpLoadFile = oCreateFolder("/UpLoadPic/Ad/")
    '------------------------------------------- Flash
    Case "Flash"
        oFileExe = "swf|flv"
        sFileSize = 10240
        UpLoadFile = oCreateFolder("/UpLoadPic/Flash/")
    '------------------------------------------- 音乐
    Case "Music"
        oFileExe = "mp3"
        sFileSize = 10240
        UpLoadFile = oCreateFolder("/UpLoadFile/Music/")
    '------------------------------------------- 视频
    Case "Video"
        oFileExe = "wmv"
        sFileSize = 51200
        UpLoadFile = oCreateFolder("/UpLoadFile/Video/")
    '------------------------------------------- 连接
    Case "Link"
        oFileExe = "*"
        sFileSize = 51200
        UpLoadFile = oCreateFolder("/UpLoadFile/Link/")
    '------------------------------------------- 文件
	Case "UFile"
        oFileExe = "rar|zip|7z|tar|exe"
        sFileSize = 51200
        UpLoadFile = oCreateFolder("/UpLoadFile/UFile/")
    '------------------------------------------- 其它
    Case Else
        oFileExe = "jpg|gif|swf|png|rar|zip|7z"
        sFileSize = 2048
        UpLoadFile = oCreateFolder("/UpLoadPic/Others/")
End Select
oFileSize = 1024 * sFileSize

oUpLoad()

Sub oUpLoad()
    Dim Upload, sPath, tempCls, fName, UpLoadFileName, sSmallPath
    '===============================================================================
    Set Upload = New AnUpLoad                     '创建类实例
    Upload.IsNum = 1                              '设置随即的样式
    Upload.IsNumIng = 12                          '设置随即值,建议使用18位以上
    Upload.SingleSize = oFileSize                 '设置单个文件最大上传限制,按字节计；默认为不限制
    Upload.MaxSize = 1024 * 1024 * 1024           '设置最大上传限制,按字节计；默认为不限制
    Upload.Exe = oFileExe                         '设置合法扩展名,以|分割,忽略大小写
    Upload.Charset = "gb2312"                     '设置文本编码，默认为gb2312
    Upload.openProcesser = False                  '禁止进度条功能，如果启用，需配合客户端程序
    Upload.GetData()                              '获取并保存数据,必须调用本方法
    '===============================================================================
    If Upload.ErrorID>0 Then                      '判断错误号,如果myupload.Err<=0表示正常
        Response.Write Upload.Description         '如果出现错误,获取错误描述
        Response.End()
    Else
        If Upload.Files( -1).Count > 0 Then       '这里判断你是否选择了文件
            sPath = Server.MapPath(UpLoadFile)    '文件保存路径
            Set tempCls = Upload.Files("file")
            tempCls.SaveToFile sPath, 0
            fName = tempCls.FileName
            Set tempCls = Nothing
			UpLoadFileName = UpLoadFile&fName
			Response.Write("{url:'"&UpLoadFileName&"',  error:'0',message:'上传成功,请勿修改上传后的路径!', width:'宽度',height:'高度'}")
			Response.End()
        Else
            Response.Write("{url:'',  error:'1',message:'上传失败,您没有上传任何文件!', width:'宽度',height:'高度'}")
            Response.End()
        End If
    End If
    Set Upload = Nothing
End Sub
'================================================
'函数名：oCreateFolder
'作  用：创建多级目录，可以创建不存在的根目录
'参  数：sPath为绝对路径
'================================================
Function oCreateFolder(ByVal sPath)
	On Error Resume Next
    Dim IsPath
    IsPath = sPath
    sPath = Replace(sPath, "/", "\")
    sPath = Replace(sPath, "\\", "\")
    Dim strHostPath, strPath
    Dim sPathItem, sTempPath
    Dim i
    Set Fso = Server.CreateObject("Scripting.FileSystemObject")
    strHostPath = Server.MapPath("/")
    If InStr(sPath, ":") = 0 Then sPath = Server.MapPath(sPath)
    If Fso.FolderExists(sPath) Or Len(sPath) < 3 Then
        oCreateFolder = IsPath
        Exit Function
    End If
    strPath = Replace(sPath, strHostPath, vbNullString, 1, -1, 1)
    sPathItem = Split(strPath, "\")
    If InStr(LCase(sPath), LCase(strHostPath)) = 0 Then
        sTempPath = sPathItem(0)
    Else
        sTempPath = strHostPath
    End If
    For i = 1 To UBound(sPathItem)
        If sPathItem(i) <> "" Then
            sTempPath = sTempPath & "\" & sPathItem(i)
            If Fso.FolderExists(sTempPath) = False Then
                Fso.CreateFolder sTempPath
            End If
        End If
    Next
    If Err.Number <> 0 Then Err.Clear
    oCreateFolder = IsPath
End Function
'===============================================
'函数名：IsObjInstalled
'作  用：检测组件
'参  数：strClassString = 组件名
'返回值：True/False
'===============================================
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
%>