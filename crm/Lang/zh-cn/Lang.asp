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
'EasyCrm �������԰�

'���ð�ť
	L_Add="���"
	L_Back="����"
	L_Clear="���"
	L_Del="ɾ��"
	L_Reply="�ظ�"
	L_Edit="�޸�"
	L_Export="����"
	L_No_data="������"
	L_Print="��ӡ"
	L_Search="����"
	L_Submit="�ύ"
	L_Select="��ѡ��"
	L_No_select="δѡ��"
	L_Transfer ="ת��"
	L_Refresh ="ˢ��"
	L_Retreat ="����"
	L_Forward ="ǰ��"
	L_Download ="����"
	L_Transfer_to ="ת��"
	L_Backlist ="�����б�"
	L_Transfer_check ="ת����ѡ"
	L_Transfer_all ="ת������"
	L_Del_check ="����ɾ��"
	L_go_info_list ="�б���ͼ"
	L_go_rl_list ="������ͼ"
	L_ReConfirm="ͨ�����"
	L_ReDenied="�ܾ�����"
	L_ReDel="����ɾ��"
	L_RealDel="����ɾ��"
	L_ReApp="����"

'���ݱ�
	L_Client = "�ͻ�����"
	L_Linkmans = "��ϵ��"
	L_Records = "������¼"
	L_Order = "������¼"
	L_Order_Products = "��������"
	L_Hetong = "��ͬ��¼"
	L_Hetong_Renew = "��ͬ����"
	L_Service = "�ۺ��¼"
	L_Expense = "���ü�¼"
	L_File = "������¼"
	L_Share = "�ͻ�����"
	L_Logfile = "������¼"

'ͷ��
	L_Header_title = "ϵͳ��ҳ"
	L_Header_company = "�ͻ�����"
	L_Header_oa = "�칫OA"
	L_Header_plugin = "���ܲ��"
	L_Header_manage = "ϵͳ����"
	L_Header_help = "��������"
	L_Header_no_login = "δ��¼"
	L_Header_logout = "�˳���¼"

'��ǰҳ��
	L_Here = "��ǰλ��"
	L_Company = "�ͻ�����"
	L_Page_Company = "�����ͻ�"
	L_Page_Listall = "���пͻ�"
	L_Page_Records = "������¼����"
	L_Page_RecordsPlan = "ԤԼ��¼����"
	L_Page_Hetong = "��ͬ��¼����"
	L_Page_Export = "����Excel"
	L_Page_Notice = "�ڲ�����"
	L_Page_Recycler = "ϵͳ����"
	L_Page_Search = "�߼�����"
	L_Page_TransData = "�ͻ�ת��"
	L_Page_OA = "�칫OA"
	L_Page_Calendar = "��������"
	L_Page_Contact = "ͨѶ¼"
	L_Page_Receive = "վ�ڶ���"
	L_Page_Report = "��������"
	L_Page_Report_add = "д����"
	L_Page_Report_view = "�Ķ�����"
	L_Page_Plugin = "�������"

'���ͷ��
	L_Top_Add_Company="¼���������"
	L_Top_Edit_Company="�޸Ļ�������"
	L_Top_View_Company="�ͻ���������"
	L_Top_Search="������ɸѡ"
	L_Top_Notice_add = "��ӹ���"
	L_Top_Notice_edit = "�޸Ĺ���"
	L_Top_Mms_add = "��д����"
	L_Top_Mms_reply = "�ظ�����"
	L_Top_Mms_view = "�鿴����"
	L_Top_Plugin = "�Ѱ�װ���"
	L_Top_Manage="����"
	
'���˵�
	lmquick   = "��ݲ˵�"
	lmliall   = "�ͻ�����"
	lmkhtj    = "�ͻ�ͳ��"
	lmnbgw    = "�ڲ�����"
	lmzndx    = "վ�ڶ���"
	lmgzbg    = "��������"
	lmzygx    = "��Դ����"
	lmyhgl    = "�û�����"
	lmgncj    = "���ܲ��"
	lmxtgl    = "ϵͳ����"
	lmlog     = "��־����"
	lmhelp    = "��������"

'��������
	L_Tip_Info_01="Ψһ��ʶ��¼��󲻿��޸�"
	L_Tip_Info_02="����010-12345678" '��ϵ�绰��ʾ
	L_Tip_Info_03="����010-12345678" '���������ʾ
	L_Tip_Info_04="�ӣ�http://" '��ҵ��վ��ʾ
	L_Tip_Info_05="����master@email.com" '�����ʼ���ʾ
	L_Tip_Info_06="�ޣ�����"
	L_Tip_Info_07="��Ч����Ϊ����" 
	L_Tip_Info_08="�޸� �� ��� �� �浵" 
	L_Tip_Info_09="����/<b style='color:#f60'>�޸�</b> �� ��� �� �浵" 
	L_Tip_Info_10="���ͨ���󣬺�ͬ�����޸�" 
	L_Tip_Info_11="��ͬ��ǰ״̬�������޸�" 

'��������ʾ
	L_Please_choose_01="��ѡ��"
	L_Please_choose_02="δѡ�����"

'����Ϣ��ʾ
	L_Notfound = "Sorry��û���ҵ�������������Ϣ��"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���뵼�� Export Improt
	L_Export_content = "������Ŀ"
	L_Export_rState = "��ǰ״̬"
	L_Export_rState_all = "ȫ��"
	L_Export_rState_0 = "������"
	L_Export_rState_1 = "����ԤԼ"
	L_Export_hOwed_all = "ȫ��"
	L_Export_hOwed_0 = "��Ƿ��"
	L_Export_hOwed_1 = "��Ƿ��"
	L_Export_text="\\ ���ɵ�Excel�ļ��浵�ڡ��칫OA���ļ��񡪵����浵������С�"
	L_Export_alert="�����ɹ���"
	L_Improt_template="�ͻ�����ģ��"
	L_Improt_template_alert="����ģ��ɹ������Ҽ���棡"
	L_Improt_alert="����ɹ���"

'����� Plugin
	L_Plugin_id ="���"
	L_Plugin_pTitle ="�������"
	L_Plugin_pUrl ="��װ·��"
	L_Plugin_pAuthor ="�������"
	L_Plugin_pVersion ="�汾"
	L_Plugin_pContent ="����˵��"
	L_Plugin_pTime ="ʱ��"
	L_Plugin_pYn ="�Ƿ�����"
	L_Plugin_pYn_0 ="�ѽ���"
	L_Plugin_pYn_1 ="������"

'���� Notice
	L_Notice_ONid ="���"
	L_Notice_ONclass ="����"
	L_Notice_ONStar ="�Ǳ�"
	L_Notice_ONtitle ="����"
	L_Notice_ONcontent ="����"
	L_Notice_ONIsread ="�Ƿ��Ķ�"
	L_Notice_ONuser ="����"
	L_Notice_ONedittime ="�༭ʱ��"
	L_Notice_ONaddtime ="���ʱ��"

'վ���� Mms
	L_Mms_id ="���"
	L_Mms_oReceiver ="�ռ���"
	L_Mms_oSender ="������"
	L_Mms_oTitle ="����"
	L_Mms_oContent ="����"
	L_Mms_oIsread ="�Ķ�"
	L_Mms_oIsread_0 ="δ��"
	L_Mms_oIsread_1 ="�Ѷ�"
	L_Mms_oIsread_all ="ȫ��"
	L_Mms_oTime ="ʱ��"
	
'�������� Report
	L_Report_id ="���"
	L_Report_oClass ="����"
	L_Report_oTitle ="����"
	L_Report_oReport ="�����ܽ�"
	L_Report_oPlan ="�����ƻ�"
	L_Report_oReply ="�쵼��ע"
	L_Report_oUser ="�ύ��"
	L_Report_oIsread ="�Ķ�"
	L_Report_oIsread_0 ="δ��"
	L_Report_oIsread_1 ="�Ѷ�"
	L_Report_oIsread_all ="ȫ��"
	L_Report_oTime ="ʱ��"
	L_Ribao="�ձ�"
	L_Zhoubao="�ܱ�"
	L_Yuebao="�±�"
	L_Jibao="����"
	L_Nianbao="�걨"
	L_whoswork="�Ĺ���"
	
'������¼��
	L_Calendar_ID ="���"
	L_Calendar_calendarDate ="¼��ʱ��"
	L_Calendar_calendarText ="��ϸ����"
	L_Calendar_calendaruser ="ҵ��Ա"
	L_Calendar_this_month="��ʾ����"
	L_Calendar_today="����"
	L_Calendar_view="����鿴����"
	L_Calendar_add="��Ӽ���"
	L_Calendar_Date="����"
	L_Calendar_Content="����"
	L_Calendar_User="������"
	L_Calendar_Time="����ʱ��"
	
	
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	

'�ֶ�����
	
	'�û���
	L_User_uId ="���"
	L_User_uAccount ="��¼�ʺ�"
	L_User_uPassword ="����"
	L_User_uName ="��ʵ����"
	L_User_uGroup ="����"
	L_User_uLevel ="Ĭ�Ͻ�ɫ"
	L_User_uQxflag ="��ϸȨ��"
	L_User_uMobile ="�ֻ�����"
	L_User_uEmail ="��������"
	L_User_uAddress ="��ϸ��ַ"
	L_User_uBirthday ="����"
	L_User_uCard ="���֤��"
	L_User_uAddtime ="��˾ʱ��"
	L_User_uManagerange ="Ȩ�޷�Χ"
	

'������¼
	L_insert_action_01 ="����"
	L_insert_action_02 ="�޸�"
	L_insert_action_03 ="ɾ��"
	L_insert_action_04 ="����ɾ��"

'��������
	L_Export_soft = "�����浵"
	L_Export_content = "������Ŀ"
	L_Export_rState = "��ǰ״̬"
	L_Export_rState_all = "ȫ��"
	L_Export_rState_0 = "������"
	L_Export_rState_1 = "����ԤԼ"
	L_Export_hOwed_all = "ȫ��"
	L_Export_hOwed_0 = "��Ƿ��"
	L_Export_hOwed_1 = "��Ƿ��"
	L_Export_text="\\ ���ɵ�Excel�ļ��浵�ڡ��칫OA���ļ��񡪵����浵������С�"
	L_Export_alert="�����ɹ���"
	L_Improt_template="�ͻ�����ģ��"
	L_Improt_template_alert="����ģ��ɹ������Ҽ���棡"
	L_Improt_alert="����ɹ���"
	L_Recycler_reDel_check="������ԭ"
	L_Recycler_true_del="����ɾ��"
	L_Recycler_sh="���"
	L_PageNum="��ҳ����"
	L_File_Type="��������"
	L_Hetong_bt1="�ſ�"
	L_Hetong_bt2="����"
	L_Botion="��Ȩ��"
	L_Pizhu="��ע"
	L_More="���"
	L_Day="��"
	L_Yuan="Ԫ"
	L_Wu="��"
	L_You="��"
	L_Hao="��"
	L_Shi="��"
	L_Fou="��"
%>