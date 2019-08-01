<% Function AddCompanyInput()%>
	<script language="JavaScript">
	<!--
	function CheckInput()
	{
		if (<%=Addbt001%>=="1"){
		if(document.all.Company.value == ""){
			alert("公司名称不能为空！");
			document.all.Company.focus();
			return false;
		}
		}
		
		if (<%=Addbt002%>=="1"){
		if(document.all.Area.value == ""){
			alert("地区大类不能为空！");
			document.all.Area.focus();
			return false;
		}
		if(document.all.Squares.value == ""){
			alert("地区小类不能为空！");
			document.all.Squares.focus();
			return false;
		}
		if(document.all.Address.value == ""){
			alert("详细地址不能为空！");
			document.all.Address.focus();
			return false;
		}
		}
		
		if (<%=Addbt003%>=="1"){
		if(document.all.Tel.value == ""){
			alert("联系电话不能为空！");
			document.all.Tel.focus();
			return false;
		}
		}
		
		if (<%=Addbt004%>=="1"){
		if(document.all.Fax.value == ""){
			alert("传真号码不能为空！");
			document.all.Fax.focus();
			return false;
		}
		}
		
		if (<%=Addbt005%>=="1"){
		if(document.all.Homepage.value == ""){
			alert("企业网站不能为空！");
			document.all.Homepage.focus();
			return false;
		}
		}
		
		if (<%=Addbt006%>=="1"){
		if(document.all.Email.value == ""){
			alert("电子邮件不能为空！");
			document.all.Email.focus();
			return false;
		}
		}
		
		if (<%=Addbt007%>=="1"){
		if(document.all.Trade.value == ""){
			alert("产品分类不能为空！");
			document.all.Trade.focus();
			return false;
		}
		if(document.all.Strades.value == ""){
			alert("产品子类不能为空！");
			document.all.Strades.focus();
			return false;
		}
		}
		
		if (<%=Addbt008%>=="1"){
		if(document.all.Type.value == ""){
			alert("客户类型不能为空！");
			document.all.Type.focus();
			return false;
		}
		}
		
		if (<%=Addbt009%>=="1"){
		if(document.all.Start.value == ""){
			alert("客户级别不能为空！");
			document.all.Start.focus();
			return false;
		}
		}
		
		if (<%=Addbt010%>=="1"){
		if(document.all.Source.value == ""){
			alert("客户来源不能为空！");
			document.all.Source.focus();
			return false;
		}
		}
	}

	-->
	</script>
<%End Function%>

<% Function AddLinkmansInput()%>
	<script language="JavaScript">
	<!--
	function CheckInput()
	{
		if (<%=Addbt011%>=="1"){
		if(document.all.lname.value == ""){
			alert("姓名不能为空！");
			document.all.lname.focus();
			return false;
		}
		}
		
		if (<%=Addbt012%>=="1"){
		if(document.all.lZhiwei.value == ""){
			alert("职位不能为空！");
			document.all.lZhiwei.focus();
			return false;
		}
		}
		
		if (<%=Addbt013%>=="1"){
		if(document.all.lQQ.value == ""){
			alert("腾讯QQ不能为空！");
			document.all.lQQ.focus();
			return false;
		}
		}
		
		if (<%=Addbt014%>=="1"){
		if(document.all.lMSN.value == ""){
			alert("MSN不能为空！");
			document.all.lMSN.focus();
			return false;
		}
		}
		
		if (<%=Addbt015%>=="1"){
		if(document.all.lMobile.value == ""){
			alert("常用手机不能为空！");
			document.all.lMobile.focus();
			return false;
		}
		}
		
		if (<%=Addbt016%>=="1"){
		if(document.all.lALWW.value == ""){
			alert("阿里旺旺不能为空！");
			document.all.lALWW.focus();
			return false;
		}
		}
		
		if (<%=Addbt017%>=="1"){
		if(document.all.lTel.value == ""){
			alert("联系电话不能为空！");
			document.all.lTel.focus();
			return false;
		}
		}
		
		if (<%=Addbt018%>=="1"){
		if(document.all.lEmail.value == ""){
			alert("个人邮箱不能为空！");
			document.all.lEmail.focus();
			return false;
		}
		}
	}

	-->
	</script>
<%End Function%>

<% Function AddRecordsInput()%>
	<script language="JavaScript">
	<!--
	function CheckInput()
	{
		if (<%=Addbt019%>=="1"){
		if(document.all.rType.value == ""){
			alert("跟单类型不能为空！");
			document.all.rType.focus();
			return false;
		}
		}
		
		if (<%=Addbt020%>=="1"){
		if(document.all.rState.value == ""){
			alert("当前进度不能为空！");
			document.all.rState.focus();
			return false;
		}
		}
		
		if (<%=Addbt021%>=="1"){
		if(document.all.rContent.value == ""){
			alert("备注说明不能为空！");
			document.all.rContent.focus();
			return false;
		}
		}
	}
	-->
	</script>
<%End Function%>

<% Function AddRecordsPlanInput()%>
	<script language="JavaScript">
	<!--
	function CheckInput()
	{
		if (<%=Addbt022%>=="1"){
		if(document.all.rDate.value == ""){
			alert("预约时间不能为空！");
			document.all.rDate.focus();
			return false;
		}
		}
		
		if (<%=Addbt023%>=="1"){
		if(document.all.rType.value == ""){
			alert("预约类型不能为空！");
			document.all.rType.focus();
			return false;
		}
		}
		
		if (<%=Addbt024%>=="1"){
		if(document.all.rlinkman.value == ""){
			alert("预约对象不能为空！");
			document.all.rlinkman.focus();
			return false;
		}
		}
		
		if (<%=Addbt025%>=="1"){
		if(document.all.rContent.value == ""){
			alert("备注说明不能为空！");
			document.all.rContent.focus();
			return false;
		}
		}
	}
	-->
	</script>
<%End Function%>

<% Function AddHetongInput()%>
	<script language="JavaScript">
	<!--
	function CheckInput()
	{
		if (<%=Addbt026%>=="1"){
		if(document.all.hEdate.value == ""){
			alert("合同截至不能为空！");
			document.all.hEdate.focus();
			return false;
		}
		}
		
		if (<%=Addbt027%>=="1"){
		if(document.all.hType.value == ""){
			alert("合同分类不能为空！");
			document.all.hType.focus();
			return false;
		}
		}
		
		if (<%=Addbt028%>=="1"){
		if(document.all.hPBigclass.value == ""){
			alert("产品分类不能为空！");
			document.all.hPBigclass.focus();
			return false;
		}
		if(document.all.Strades.value == ""){
			alert("产品子类不能为空！");
			document.all.Strades.focus();
			return false;
		}
		}
		
		if (<%=Addbt029%>=="1"){
		if(document.all.hMoney.value == ""){
			alert("总金额不能为空！");
			document.all.hMoney.focus();
			return false;
		}
		}
		
		if (<%=Addbt030%>=="1"){
		if(document.all.hRevenue.value == ""){
			alert("预付款不能为空！");
			document.all.hRevenue.focus();
			return false;
		}
		}
		
		if (<%=Addbt031%>=="1"){
		if(document.all.hInvoice.value == ""){
			alert("发票不能为空！");
			document.all.hInvoice.focus();
			return false;
		}
		}
		
		if (<%=Addbt032%>=="1"){
		if(document.all.hContent.value == ""){
			alert("附加条款不能为空！");
			document.all.hContent.focus();
			return false;
		}
		}
	}
	-->
	</script>
<%End Function%>

<% Function AddReceiveInput()%>
<script language="JavaScript">
function CheckInput(){
	if(document.all.oReceiver.value == ""){alert("请填写收件人！");document.all.oReceiver.focus();return false;}
	if(document.all.oTitle.value == ""){alert("请填写标题！");document.all.oTitle.focus();return false;}
	if(document.all.oContent.value == ""){alert("请填写内容");document.all.oContent.focus();return false;}
}
</script>
<%End Function%>

<% Function AddNoticeInput()%>
<script language="JavaScript">
function CheckInput(){
	if(document.all.ONclass.value == ""){alert("请选择分类！");document.all.ONclass.focus();return false;}
	if(document.all.ONStar.value == ""){alert("请选择星标！");document.all.ONStar.focus();return false;}
	if(document.all.ONtitle.value == ""){alert("请填写标题");document.all.ONtitle.focus();return false;}
	if(document.all.ONcontent.value == ""){alert("请填写内容");document.all.ONcontent.focus();return false;}
}
</script>
<%End Function%>