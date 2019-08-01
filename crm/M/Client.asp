<!--#include file="inc.asp"--><%=Header%>
    
    <!-- start searchbox -->
    <div class="searchbox">
   	  <form id="form1" name="form1" method="post" action="">
      	<input type="text" name="textfield" id="textfield" class="txtbox" />
   	  </form>
    </div>
    <!-- end searchbox -->
    
    
    
    <!-- start page -->
    <div class="page">
    
    		
            <!-- start profile box -->
            <div class="profilebox">
            	<img src="img/avatar.png" width="19" height="20" alt="avatar" class="avatar"/> 欢迎 <b><%=Session("CRM_name")%></b> 登录系统
                <a href="#" class="logout" title="退出">退出</a>
                <div class="clear"></div>
            </div>
            <!-- end profile box -->
            
            
            
            <!-- start menu -->
           	 <ul id="menu">
             	<li><a href="Listall.asp"><img src="img/icons/files.png" width="21" height="21" alt="icon" class="m-icon"/><b>所有客户</b></a></li>
             	<li><a href="form.html"><img src="img/icons/bubble.png" width="29" height="21" alt="icon" class="m-icon"/><b>表单组件</b></a></li>
   	         	<li><a href="statistics.html"><img src="img/icons/graph.png" width="24" height="21" alt="icon" class="m-icon"/><b>数据统计 <span>9</span> </b></a></li>
             	<li><a href="alert-boxes.html"><img src="img/icons/alert.png" width="25" height="21" alt="icon" class="m-icon"/><b>提示信息 <span class="red">15</span> </b></a></li>
   	         	<li><a href="typo.html"><img src="img/icons/personal-folder.png" width="29" height="21" alt="icon" class="m-icon"/><b>文字排版</b></a></li>
             	<li><a href="gallery.html"><img src="img/icons/photo-gallery.png" width="29" height="21" alt="icon" class="m-icon"/><b>图片相册<span>93</span> </b></a></li>
             	<li><a href="table.html"><img src="img/icons/blocks.png" width="26" height="21" alt="icon" class="m-icon"/><b>数据表格</b></a></li>
             	<li><a href="simple-page.html"><img src="img/icons/page.png" width="26" height="21" alt="icon" class="m-icon"/><b>内容页面</b></a></li>
             	<li><a href="error-page.html"><img src="img/icons/error.png" width="26" height="21" alt="icon" class="m-icon"/><b>错误页面  <span class="red">1</span></b></a></li>
             </ul>
            <!-- end menu -->
            
            
            <!-- start top button -->
            <div class="topbutton"><a href="#"><span>Top</span></a></div>
            <!-- end top button -->
            
            
            
		<%=Footer%>
            
    
    <div class="clear"></div>
    </div>
    <!-- end page -->
<script type="text/javascript" src="js/frame.js"></script>
</body>
</html>
