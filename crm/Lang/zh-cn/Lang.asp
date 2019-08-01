<!-- #include file="Alert.asp" -->
<!-- #include file="Client.asp" -->
<!-- #include file="Linkmans.asp" -->
<!-- #include file="Records.asp" -->
<!-- #include file="Order.asp" -->
<!-- #include file="Hetong.asp" -->
<!-- #include file="Service.asp" -->
<!-- #include file="Expense.asp" -->
<!-- #include file="File.asp" -->
<!-- #include file="Logfile.asp" -->
<!-- #include file="Customer.asp" -->
<% 
'EasyCrm 中文语言包

'常用按钮
	L_Add="添加"
	L_Back="返回"
	L_Clear="清空"
	L_Del="删除"
	L_Reply="回复"
	L_Edit="修改"
	L_Export="导出"
	L_No_data="无数据"
	L_Print="打印"
	L_Search="搜索"
	L_Submit="提交"
	L_Select="请选择"
	L_No_select="未选择"
	L_Transfer ="转移"
	L_Refresh ="刷新"
	L_Retreat ="后退"
	L_Forward ="前进"
	L_Download ="下载"
	L_Transfer_to ="转给"
	L_Backlist ="返回列表"
	L_Transfer_check ="转移所选"
	L_Transfer_all ="转移所有"
	L_Del_check ="批量删除"
	L_go_info_list ="列表视图"
	L_go_rl_list ="日历视图"
	L_ReConfirm="通过审核"
	L_ReDenied="拒绝申请"
	L_ReDel="撤销删除"
	L_RealDel="彻底删除"
	L_ReApp="申请"

'数据表
	L_Client = "客户档案"
	L_Linkmans = "联系人"
	L_Records = "跟单记录"
	L_Order = "订单记录"
	L_Order_Products = "订单详情"
	L_Hetong = "合同记录"
	L_Hetong_Renew = "合同续费"
	L_Service = "售后记录"
	L_Expense = "费用记录"
	L_File = "附件记录"
	L_Share = "客户共享"
	L_Logfile = "操作记录"

'头部
	L_Header_title = "系统首页"
	L_Header_company = "客户管理"
	L_Header_oa = "办公OA"
	L_Header_plugin = "功能插件"
	L_Header_manage = "系统设置"
	L_Header_help = "帮助中心"
	L_Header_no_login = "未登录"
	L_Header_logout = "退出登录"

'当前页面
	L_Here = "当前位置"
	L_Company = "客户管理"
	L_Page_Company = "新增客户"
	L_Page_Listall = "所有客户"
	L_Page_Records = "跟单记录管理"
	L_Page_RecordsPlan = "预约记录管理"
	L_Page_Hetong = "合同记录管理"
	L_Page_Export = "导出Excel"
	L_Page_Notice = "内部公文"
	L_Page_Recycler = "系统公海"
	L_Page_Search = "高级搜索"
	L_Page_TransData = "客户转移"
	L_Page_OA = "办公OA"
	L_Page_Calendar = "个人日历"
	L_Page_Contact = "通讯录"
	L_Page_Receive = "站内短信"
	L_Page_Report = "工作报告"
	L_Page_Report_add = "写报告"
	L_Page_Report_view = "阅读报告"
	L_Page_Plugin = "插件管理"

'表格头部
	L_Top_Add_Company="录入基本档案"
	L_Top_Edit_Company="修改基本档案"
	L_Top_View_Company="客户基本档案"
	L_Top_Search="按条件筛选"
	L_Top_Notice_add = "添加公文"
	L_Top_Notice_edit = "修改公文"
	L_Top_Mms_add = "编写短信"
	L_Top_Mms_reply = "回复短信"
	L_Top_Mms_view = "查看短信"
	L_Top_Plugin = "已安装插件"
	L_Top_Manage="管理"
	
'左侧菜单
	lmquick   = "快捷菜单"
	lmliall   = "客户管理"
	lmkhtj    = "客户统计"
	lmnbgw    = "内部公文"
	lmzndx    = "站内短信"
	lmgzbg    = "工作报告"
	lmzygx    = "资源共享"
	lmyhgl    = "用户管理"
	lmgncj    = "功能插件"
	lmxtgl    = "系统管理"
	lmlog     = "日志管理"
	lmhelp    = "帮助中心"

'描述文字
	L_Tip_Info_01="唯一标识，录入后不可修改"
	L_Tip_Info_02="例：010-12345678" '联系电话提示
	L_Tip_Info_03="例：010-12345678" '传真号码提示
	L_Tip_Info_04="加：http://" '企业网站提示
	L_Tip_Info_05="例：master@email.com" '电子邮件提示
	L_Tip_Info_06="限：数字"
	L_Tip_Info_07="生效日期为今天" 
	L_Tip_Info_08="修改 → 审核 → 存档" 
	L_Tip_Info_09="新增/<b style='color:#f60'>修改</b> → 审核 → 存档" 
	L_Tip_Info_10="审核通过后，合同不能修改" 
	L_Tip_Info_11="合同当前状态不允许修改" 

'下拉框提示
	L_Please_choose_01="请选择"
	L_Please_choose_02="未选择大类"

'无信息提示
	L_Notfound = "Sorry，没有找到符合条件的信息！"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'导入导出 Export Improt
	L_Export_content = "导出项目"
	L_Export_rState = "当前状态"
	L_Export_rState_all = "全部"
	L_Export_rState_0 = "待处理"
	L_Export_rState_1 = "过期预约"
	L_Export_hOwed_all = "全部"
	L_Export_hOwed_0 = "无欠款"
	L_Export_hOwed_1 = "有欠款"
	L_Export_text="\\ 生成的Excel文件存档在【办公OA—文件柜—导出存档】类别中。"
	L_Export_alert="导出成功！"
	L_Improt_template="客户档案模版"
	L_Improt_template_alert="生成模版成功，请右键另存！"
	L_Improt_alert="导入成功！"

'插件表 Plugin
	L_Plugin_id ="编号"
	L_Plugin_pTitle ="插件名称"
	L_Plugin_pUrl ="安装路径"
	L_Plugin_pAuthor ="插件开发"
	L_Plugin_pVersion ="版本"
	L_Plugin_pContent ="功能说明"
	L_Plugin_pTime ="时间"
	L_Plugin_pYn ="是否启用"
	L_Plugin_pYn_0 ="已禁用"
	L_Plugin_pYn_1 ="已启用"

'公文 Notice
	L_Notice_ONid ="编号"
	L_Notice_ONclass ="分类"
	L_Notice_ONStar ="星标"
	L_Notice_ONtitle ="标题"
	L_Notice_ONcontent ="内容"
	L_Notice_ONIsread ="是否阅读"
	L_Notice_ONuser ="作者"
	L_Notice_ONedittime ="编辑时间"
	L_Notice_ONaddtime ="添加时间"

'站内信 Mms
	L_Mms_id ="编号"
	L_Mms_oReceiver ="收件人"
	L_Mms_oSender ="发件人"
	L_Mms_oTitle ="标题"
	L_Mms_oContent ="内容"
	L_Mms_oIsread ="阅读"
	L_Mms_oIsread_0 ="未读"
	L_Mms_oIsread_1 ="已读"
	L_Mms_oIsread_all ="全部"
	L_Mms_oTime ="时间"
	
'工作报告 Report
	L_Report_id ="编号"
	L_Report_oClass ="分类"
	L_Report_oTitle ="标题"
	L_Report_oReport ="工作总结"
	L_Report_oPlan ="工作计划"
	L_Report_oReply ="领导批注"
	L_Report_oUser ="提交人"
	L_Report_oIsread ="阅读"
	L_Report_oIsread_0 ="未读"
	L_Report_oIsread_1 ="已读"
	L_Report_oIsread_all ="全部"
	L_Report_oTime ="时间"
	L_Ribao="日报"
	L_Zhoubao="周报"
	L_Yuebao="月报"
	L_Jibao="季报"
	L_Nianbao="年报"
	L_whoswork="的工作"
	
'日历记录表
	L_Calendar_ID ="编号"
	L_Calendar_calendarDate ="录入时间"
	L_Calendar_calendarText ="详细内容"
	L_Calendar_calendaruser ="业务员"
	L_Calendar_this_month="显示当月"
	L_Calendar_today="今天"
	L_Calendar_view="点击查看详情"
	L_Calendar_add="添加记事"
	L_Calendar_Date="日期"
	L_Calendar_Content="内容"
	L_Calendar_User="创建人"
	L_Calendar_Time="创建时间"
	
	
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	

'字段名称
	
	'用户表
	L_User_uId ="编号"
	L_User_uAccount ="登录帐号"
	L_User_uPassword ="密码"
	L_User_uName ="真实姓名"
	L_User_uGroup ="部门"
	L_User_uLevel ="默认角色"
	L_User_uQxflag ="详细权限"
	L_User_uMobile ="手机号码"
	L_User_uEmail ="电子邮箱"
	L_User_uAddress ="详细地址"
	L_User_uBirthday ="生日"
	L_User_uCard ="身份证号"
	L_User_uAddtime ="入司时间"
	L_User_uManagerange ="权限范围"
	

'操作记录
	L_insert_action_01 ="新增"
	L_insert_action_02 ="修改"
	L_insert_action_03 ="删除"
	L_insert_action_04 ="彻底删除"

'其它杂项
	L_Export_soft = "导出存档"
	L_Export_content = "导出项目"
	L_Export_rState = "当前状态"
	L_Export_rState_all = "全部"
	L_Export_rState_0 = "待处理"
	L_Export_rState_1 = "过期预约"
	L_Export_hOwed_all = "全部"
	L_Export_hOwed_0 = "无欠款"
	L_Export_hOwed_1 = "有欠款"
	L_Export_text="\\ 生成的Excel文件存档在【办公OA—文件柜—导出存档】类别中。"
	L_Export_alert="导出成功！"
	L_Improt_template="客户档案模版"
	L_Improt_template_alert="生成模版成功，请右键另存！"
	L_Improt_alert="导入成功！"
	L_Recycler_reDel_check="批量还原"
	L_Recycler_true_del="彻底删除"
	L_Recycler_sh="审核"
	L_PageNum="分页数量"
	L_File_Type="允许类型"
	L_Hetong_bt1="概况"
	L_Hetong_bt2="财务"
	L_Botion="无权限"
	L_Pizhu="批注"
	L_More="多个"
	L_Day="天"
	L_Yuan="元"
	L_Wu="无"
	L_You="有"
	L_Hao="号"
	L_Shi="是"
	L_Fou="否"
%>