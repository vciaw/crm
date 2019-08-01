<%
Dim title,SiteUrl,Skinurl,DataPageSize,YNalert,CRTypeEnd,YnUserLog,YnDDNum,YnHTNum,YnDelReason,SaveOldUser,YNRecycler,ClientOnly,SelectCharset,language,uploadtype,Keeponline,CookieKey
'全局配置
title="客户管理系统" '系统名称
SiteUrl="/" '安装目录
Skinurl="Skin/default/" '风格路径
DataPageSize=10 '分页数量
YNalert=1 '操作提示
CRTypeEnd="" '跟进流程结束状态
YnUserLog=1 '记录登录日志
YnDDNum=1 '订单编号生成方式
YnHTNum=1 '合同编号生成方式
YnDelReason=1 '删除客户档案是否需要填写原因
SaveOldUser=1 '客户转移后，是否保留原有业务员
YNRecycler=1 '公海申请客户是否需要审核
ClientOnly="100" '判断客户唯一的标准
SelectCharset=1 '处理乱码
language="zh-cn" '系统语言
uploadtype="gif/jpg/png/bmp/doc/xls/ppt/rar/zip" '上传文件后缀
Keeponline=1 '保持在线
CookieKey="01A6A1F4EB945D24" '识别码

gdzy="300000000000000" '跟单转移

%>
