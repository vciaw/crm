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
    '------------------------------------------- ͼƬ
    Case "Images"
        oFileExe = "jpg|jpeg|gif|bmp|swf|png"
        sFileSize = 2048
        UpLoadFile = oCreateFolder("../../../UpLoad/"&year(now())&"/Images/")
	'------------------------------------------- ͼƬ
    Case "Img"
        oFileExe = "jpg|jpeg|gif|bmp|png"
        sFileSize = 2048
        UpLoadFile = oCreateFolder("/UpLoadPic/Img/")
    '------------------------------------------- ���ͼƬ
    Case "Ad"
        oFileExe = "jpg|jpeg|gif|bmp|swf|png"
        sFileSize = 500
        UpLoadFile = oCreateFolder("/UpLoadPic/Ad/")
    '------------------------------------------- Flash
    Case "Flash"
        oFileExe = "swf|flv"
        sFileSize = 10240
        UpLoadFile = oCreateFolder("/UpLoadPic/Flash/")
    '------------------------------------------- ����
    Case "Music"
        oFileExe = "mp3"
        sFileSize = 10240
        UpLoadFile = oCreateFolder("/UpLoadFile/Music/")
    '------------------------------------------- ��Ƶ
    Case "Video"
        oFileExe = "wmv"
        sFileSize = 51200
        UpLoadFile = oCreateFolder("/UpLoadFile/Video/")
    '------------------------------------------- ����
    Case "Link"
        oFileExe = "*"
        sFileSize = 51200
        UpLoadFile = oCreateFolder("/UpLoadFile/Link/")
    '------------------------------------------- �ļ�
	Case "UFile"
        oFileExe = "rar|zip|7z|tar|exe"
        sFileSize = 51200
        UpLoadFile = oCreateFolder("/UpLoadFile/UFile/")
    '------------------------------------------- ����
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
    Set Upload = New AnUpLoad                     '������ʵ��
    Upload.IsNum = 1                              '�����漴����ʽ
    Upload.IsNumIng = 12                          '�����漴ֵ,����ʹ��18λ����
    Upload.SingleSize = oFileSize                 '���õ����ļ�����ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
    Upload.MaxSize = 1024 * 1024 * 1024           '��������ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
    Upload.Exe = oFileExe                         '���úϷ���չ��,��|�ָ�,���Դ�Сд
    Upload.Charset = "gb2312"                     '�����ı����룬Ĭ��Ϊgb2312
    Upload.openProcesser = False                  '��ֹ���������ܣ�������ã�����Ͽͻ��˳���
    Upload.GetData()                              '��ȡ����������,������ñ�����
    '===============================================================================
    If Upload.ErrorID>0 Then                      '�жϴ����,���myupload.Err<=0��ʾ����
        Response.Write Upload.Description         '������ִ���,��ȡ��������
        Response.End()
    Else
        If Upload.Files( -1).Count > 0 Then       '�����ж����Ƿ�ѡ�����ļ�
            sPath = Server.MapPath(UpLoadFile)    '�ļ�����·��
            Set tempCls = Upload.Files("file")
            tempCls.SaveToFile sPath, 0
            fName = tempCls.FileName
            Set tempCls = Nothing
			UpLoadFileName = UpLoadFile&fName
			Response.Write("{url:'"&UpLoadFileName&"',  error:'0',message:'�ϴ��ɹ�,�����޸��ϴ����·��!', width:'���',height:'�߶�'}")
			Response.End()
        Else
            Response.Write("{url:'',  error:'1',message:'�ϴ�ʧ��,��û���ϴ��κ��ļ�!', width:'���',height:'�߶�'}")
            Response.End()
        End If
    End If
    Set Upload = Nothing
End Sub
'================================================
'��������oCreateFolder
'��  �ã������༶Ŀ¼�����Դ��������ڵĸ�Ŀ¼
'��  ����sPathΪ����·��
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
'��������IsObjInstalled
'��  �ã�������
'��  ����strClassString = �����
'����ֵ��True/False
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