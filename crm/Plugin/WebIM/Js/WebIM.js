var _webIM = null;
function IntWebIM()
{
	if(!$("windowContainerBorder"))
	{
		new WebIM().Initialize();
	}
}
function TempChatMain() {
            IntWebIM()
}

function WebIM()
{
	var _me = (_webIM=this);
	this.Sys = null;
	this.Profile = null;  //���˵���
	this.Config  = null;  //��������
	this.Friend  = null;  //�����б�
	this.Group   = null;  //������Ϣ
	this.Win     = null;  //�������
	this.Version = "V3.1.1";   //�汾��
	//����
	this.Common = {
		//��������
		playSound:function(soundname)
		{
			var oDiv = $("divSound");
			if(!oDiv)return;
			Elem.Value(oDiv,"<embed id=\"sound\" name=\"sound\" pluginspage=\"http://www.macromedia.com/go/getflashplayer\" src=\"sound/"+soundname+".swf\" width=\"1\" height=\"1\" type=\"application/x-shockwave-flash\" autoplay=\"true\" quality=\"high\" loop=\"False\"></embed>");
		},
		//���ַ������滻��ͼƬ
		replaceFaceFromStr:function(str)
		{
			var faces = [];
			faces = _me.Data.getMsnFaceInfo();
			for(var i=0;i<faces.length;i++)
			{
				var face = faces[i];
				str = str.replaceAll(face[2],"<img style='height:24px;width:24px' src='qqface/"+face[0]+".gif'/>");
				if(face.length>3)str = str.replaceAll(face[3],"<img style='height:24px;width:24px' src='qqface/"+face[0]+".gif'/>");
			}
			return str;
		},
		//����UserStatus���ض�Ӧ�ַ���
		getUserStatusStr:function(n)
		{
			return ["����","æµ","���ϻ���","�뿪","ͨ����","����Ͳ�","��ʾΪ����","ע��"][n];
		},
		//����OnlineStatus���ض�Ӧ�ַ���
		getOnlineStatusStr:function(n)
		{
			return ["����","æµ","�뿪","����"][n];
		},
		//����UserStatus��������״̬OnlineStatus
		getOnlineStatus:function(n)
		{
			return [0,1,2,2,1,2,3,3,3][n];
		},
		//��GroupID�õ�Group����
		getGroupById:function(id)
		{
			var g = null;
			for(var i=0;i<_me.Group.length;i++)
			{
				if(_me.Group[i].ID==id)
				{
					g = _me.Group[i];
					break;
				}
			}
			return g;
		},
		//��GroupName�õ�Group����
		getGroupByName:function(gname)
		{
			var g = null;
			for(var i=0;i<_me.Group.length;i++)
			{
				if(_me.Group[i].Name==gname)
				{
					g = _me.Group[i];
					break;
				}
			}
			return g;
		},
		//��xml����User����
		getUserFromXml:function(xml)
		{
			var users = [];
			var items=$T(xml,'item');
			if(!items||items.length<1)return null;
			for(var i=0;i<items.length;i++)
			{
				var item = items[i];
				var face      = Xml.First(item,"f");
				var id        = parseInt(Xml.First(item,"id"));
				var name      = Xml.First(item,"n");
				var email     = Xml.First(item,"e");
				var sign      = Xml.First(item,"sn");
				var status    = Xml.First(item,"s");
				var groupid   = Xml.First(item,"g");
				var isblocked = Xml.First(item,"b")=="1";
				var customname= Xml.First(item,"cn");
				var gender    = Xml.First(item,"u");
				if(!_me.Common.getUserFromArr(id,users))//��ֹ�ظ�
					users.add(new _me.Model.User(face,id,name,email,sign,parseInt(status),parseInt(groupid),isblocked,customname,gender));
			}
			return users;
		},
		//����UserID��User�����з���User����
		getUserFromArr:function(id,users)
		{
			for(var i=0;i<users.length;i++)
			{
				if(users[i].UserID==id)return users[i];
			}
			return null;
		},
		//��xml����Group����
		getGroupFromXml:function(xml)
		{
			var groups = [];
			var items=$T(xml,'item');
			for(var i=0;i<items.length;i++)
			{
				var item = items[i];
				var name      = Xml.First(item,"Name");
				var id        = Xml.First(item,"ID");
				groups.add(new _me.Model.Group(name,parseInt(id)));
			}
			if(groups.length<1)groups.add(new _me.Model.Group("Ĭ��",1));
			return groups;
		},
		//��xml����Msg����
		getMsgFromXml:function(xml)
		{
			var msgs = [];
			var items=$T(xml,'item');
			for(var i=0;i<items.length;i++)
			{
				var item = items[i];
				var from      = Xml.First(item,"From");
				var to        = Xml.First(item,"To");
				var content      = Xml.First(item,"Content");
				var type        = Xml.First(item,"Type");
				var isconfirm      = Xml.First(item,"IsConfirm");
				var addtime        = Xml.First(item,"AddTime");
				msgs.add(new _me.Model.Msg(from,to,content,type,isconfirm,addtime));
			}
			return msgs;
		},
		//��xml����Config����
		getConfigFromXml:function(xml)
		{
			var items=$T(xml,'item');
			if(!items||items.length<1)return null;
			var item = items[0];
			var distype      = Xml.First(item,"DisType");
			var ordertype        = Xml.First(item,"OrderType");
			var chatside      = Xml.First(item,"ChatSide");
			var msgsendkey     = Xml.First(item,"MsgSendKey");
			var msgshowtime     = Xml.First(item,"MsgShowTime");
			var userpower     = Xml.First(item,"UserPower");
			return new _me.Model.Config(parseInt(distype),parseInt(ordertype),parseInt(chatside),parseInt(msgsendkey),parseInt(msgshowtime),parseInt(userpower));
		},
		//��������
		searchFriendList:function(k)
		{
			var result = [];
			for(var i=0;i<_me.Friend.length;i++)
			{
				var p = _me.Friend[i];
				if(p.UserName.indexOfEx(k)>-1||p.UserEmail.indexOfEx(k)>-1||p.CustomName.indexOfEx(k)>-1)//�ǳơ����䡢�Զ����ǳ�
				{
					result.add(p);
				}
			}
			return result;
		},
		//�Ժ����б�����
		sortFriendList:function()
		{
			switch(_me.Config.OrderType)
			{
				case 1://״̬
					_me.Friend.sort(
						function(a,b)
						{
							if(a.OnlineStatus==b.OnlineStatus)
							{
								return a.UserName.localeCompare(b.UserName);
							}
							else
							{
								return a.OnlineStatus<b.OnlineStatus?-1:1;
							}
						});
					return;
				case 2://����
					_me.Friend.sort(
						function(a,b)
						{
							if(a.GroupID==b.GroupID)
							{
								if(a.OnlineStatus==b.OnlineStatus)
								{
									return a.UserName.localeCompare(b.UserName);
								}
								else
								{
									return a.OnlineStatus<b.OnlineStatus?-1:1;
								}
							}
							else
							{
								return a.GroupID<b.GroupID?-1:1;
							}
						});
					return;
			}
		},
		//�ı�ѡ��item
		changeSelItem:function(o,evt)
		{
			var objSel = _me.Sys["ObjSel"];
			var e = evt||event;
			if(objSel&&objSel!=o)
			{
				objSel.style.backgroundColor = "";
			}
			if(Other.Browser()=="opera")
			{
				if($("btnOpera"))Elem.Del("btnOpera");
				var btn = Elem.New("input","btnOpera");
				btn.type = "button";
				var bs = btn.style;
				bs.zIndex = "9999";
				bs.opacity = "0.01";
				bs.height = "20px";
				bs.width = "20px";
				bs.top = (Evt.Top(e)-5)+"px";
				bs.left = (Evt.Left(e)-5)+"px";
				bs.position = "absolute";
				Elem.Add(btn);
			}
			o.style.backgroundColor = "#D2EAF6";
			_me.Sys["ObjSel"] = o;
		},
		//��ϵ���Ҽ��˵�
		showContextMenu:function(o,evt)
		{
			Other.Browser()=="opera"&&Elem.Del("btnOpera");
			var e = evt||event;
			if(parseInt(e.button)==2)//�Ҽ��˵�
			{
				var id = _me.Win.id;
				var t = _me.Win.type;
				var uid = o.getAttribute("uid");
				var gid = o.getAttribute("gid");
				if(uid)
				{
					var u = _me.Common.getUserFromArr(uid,_me.Friend);
					if(!u)return;
					var arr = [];
					arr.add(u.OnlineStatus<3?"<strong>���ͼ�ʱ��Ϣ</strong>|_webIM.CMD.openChatWindow("+uid+",true)":"<strong>�����ѻ���ʱ��Ϣ</strong>|_webIM.CMD.openChatWindow("+uid+",true)");
					//arr.add("���͵����ʼ�|_webIM.Common.showLink(250,200,(Elem.Height()-200)/2,(Elem.Width()-250)/2,11,'���͵����ʼ�','page/sendemail.asp?e="+u.UserEmail+"')");
					//arr.add("");
					//arr.add("�鿴��ϵ�˿�Ƭ|_webIM.CMD.showCard("+uid+",$('wMainUserItemId"+id+"Type"+t+"No"+uid+"'))");
					arr.add("�鿴��Ϣ��ʷ��¼|_webIM.CMD.showMsgHistory("+uid+")");
					arr.add(u.CustomName==u.UserName?"�޸ı�ע����|_webIM.CMD.editCustomName("+uid+")":"�޸ı�ע����|_webIM.CMD.editCustomName("+uid+")");
					if(_me.Config.OrderType==2)arr.add("�༭����|_webIM.CMD.eidtUserGroup("+uid+")");
					//arr.add(u.IsBlocked?"ȡ����ֹ��ϵ��|_webIM.CMD.blockFriend("+uid+",2)":"��ֹ��ϵ��|_webIM.CMD.blockFriend("+uid+",1)");
					arr.add("ɾ����ϵ��|_webIM.CMD.delFriend("+uid+")");
				}
				else if(gid)
				{
					var arr = [];
					arr.add("��������|_webIM.CMD.editGroup("+gid+")");
					arr.add("ɾ����|_webIM.CMD.delGroup("+gid+")");
					arr.add("");
					arr.add("��������|_webIM.CMD.addGroup()");
				}
				else
				{
					return;
				}
				var m = new SysMenu("menuUserNameId"+id+"Type"+t);
				m.Data = arr;
				m.Width = 130;
				m.E = e;
				m.Show();
			}
		},
		//����һ��itemͷ,obj:����,title:����,no:���
		addUserHeader:function(obj,title,no)
		{
			var id = _me.Win.id;
			var t = _me.Win.type;
			var header = Elem.New("div","wMainUserHeaderId"+id+"Type"+t+"No"+no,"wMainUserHeaderEx",title);
			header.onmouseover = function()
			{
				this.className = "wMainUserHeaderEx wMainUserHeaderExHover";
			};
			header.onmouseout = function()
			{
				this.className = "wMainUserHeaderEx";
			};
			if(_me.Config.OrderType==2)header.setAttribute("gid",no);
			header.onmousedown = function(e){_me.Common.changeSelItem(this,e)};
			header.onmouseup = function(e){_me.Common.showContextMenu(this,e)};
			header.onclick = function()
			{
				if(this.className=="wMainUserHeader wMainUserHeaderHover")
				{
					this.className = "wMainUserHeaderEx wMainUserHeaderExHover";
					this.onmouseout = function(){this.className = "wMainUserHeaderEx";};
					this.onmouseover = function(){this.className = "wMainUserHeaderEx wMainUserHeaderExHover";};
					Elem.Show(this.nextSibling);
				}
				else
				{
					this.className = "wMainUserHeader wMainUserHeaderHover";
					this.onmouseout = function(){this.className = "wMainUserHeader";};
					this.onmouseover = function(){this.className = "wMainUserHeader wMainUserHeaderHover";};
					Elem.Hid(this.nextSibling);
				}
			};
			var container = Elem.New("div","wMainUserContainerId"+id+"Type"+t+"No"+no,"wMainUserContainer");
			Elem.Add(obj,header,container);
		},
		//����һ��item,n:1��Ҫ2��ϸ,u:User����
		createUserItem:function(n,u)
		{
			var id = _me.Win.id;
			var t = _me.Win.type;
			var ban = u.IsBlocked?"b":"";
			var str = new StringBuilder();
			switch(n)
			{
				case 1:
					str.add("<div id=\"wMainUserItemId"+id+"Type"+t+"No"+u.UserID+"\" class=\"wMainUserItem\" uid=\""+u.UserID+"\" onmousedown=\"_webIM.Common.changeSelItem(this,event)\" onmouseup=\"_webIM.Common.showContextMenu(this,event)\" oncontextmenu=\"function(){return !!0}\" ondblclick=\"var uid = this.getAttribute('uid');if(uid)_webIM.CMD.openChatWindow(uid,true);\">");
					str.add("<div class=\"wMainListButton\" onmouseover=\"this.className='wMainListButton wMainListButtonHover'\" onmouseout=\"this.className='wMainListButton'\">");
					str.add("<img src=\"images/m"+u.OnlineStatus+ban+".gif\" style=\"height:16px;width:16px\" />");
					//str.add("<img src=\"images/m"+u.OnlineStatus+ban+".gif\" title=\"�鿴���˵���ϵ�˿�Ƭ\" style=\"height:19px;width:19px\" onclick=\"_webIM.CMD.showCard("+u.UserID+",this.parentNode)\"/>");
					str.add("</div>");
					str.add("<div class=\"wMainUserItemText\">"+u.CustomName+"&nbsp;��&nbsp;<span style=\"color:#777\">"+u.UserSign+"</span></div>");
					str.add("</div>");
					break;
				case 2:
					str.add("<div id=\"wMainUserItemId"+id+"Type"+t+"No"+u.UserID+"\" class=\"wMainUserItemBig\" uid=\""+u.UserID+"\" onmousedown=\"_webIM.Common.changeSelItem(this,event)\" onmouseup=\"_webIM.Common.showContextMenu(this,event)\" oncontextmenu=\"function(){return !!0}\" ondblclick=\"var uid = this.getAttribute('uid');if(uid)_webIM.CMD.openChatWindow(uid,true);\">");
					str.add("<div class=\"wMainUserItemFaceBG\"><img style=\"width:46px;height:45px\" src=\"images/userface/"+u.UserFace+"\" /></div>");
					str.add("<div class=\"wMainUserItemBigRight\">");
					str.add("<div class=\"wMainUserItemStatus\" ><img style=\"height:19px;width:19px\" src=\"images/m"+u.OnlineStatus+ban+".gif\"/></div>");
					str.add("<div style=\"padding-top:2px\">"+u.CustomName+"</div>");
					str.add("<div style=\"padding-top:2px;color:#777\">"+u.UserSign+"</div>");
					str.add("<div style=\"padding-top:2px\">"+u.UserEmail+"</div>");
					str.add("</div>");
					str.add("</div>");
					break;
				case 3:
					str.add("<div class=\"wMainUserItemBig\" style=\"padding-left:0\">");
					str.add("<div class=\"wMainUserItemFaceBG\"><img style=\"width:46px;height:45px\" src=\"images/userface/"+u.UserFace+"\" /></div>");
					str.add("<div class=\"wMainUserItemBigRight\">");
					str.add("<div class=\"wMainUserItemStatus\" ><img style=\"height:19px;width:19px\" src=\"images/m"+u.OnlineStatus+ban+".gif\"/></div>");
					str.add("<div style=\"padding-top:2px\">"+u.CustomName+"</div>");
					str.add("<div style=\"padding-top:2px;color:#777\">"+u.UserSign+"</div>");
					str.add("<div style=\"padding-top:2px\">"+u.UserEmail+"</div>");
					str.add("</div>");
					str.add("</div>");
					break;
			}
			return str.toString();
		},
		//��ʾһ�������
		showAlert:function(msg,title,closecb,url,icon,loadedcb)
		{
			if($("windowCover"))
			{
				var cs = $("windowCover").style;
				cs.display = "block";
				cs.height = Elem.Height()+"px";
				cs.width = Elem.Width()+"px";
				cs.zIndex = ++window.zIndex;
			}
			var w = new WebForm();
			w.Title = title||"ע��";
			w.Icon = icon||"warning.gif";
			w.Type = 3;
			w.UserID = 4;
			w.Height = 120;
			w.Width = 210;
			w.CanControl = !!0;
			w.Resizeable = !!0;
			w.ShowCorner = !!0;
			w.ContentUrl = url||"layout/winalert.htm";
			w.Left = (Elem.Width()-210)/2;
			w.Top = (Elem.Height()-130)/2;
			w.RepaintMethod = function(w,h,id,t){$("wOtherMainId"+id+"Type"+t).style.height = h+"px";if($("divMsgId"+id+"Type"+t))Elem.Value("divMsgId"+id+"Type"+t,msg);};
			w.CloseCallback = closecb||null;
			w.LoadedCallback = function(id,t)
			{
				$("btnSelId"+id+"Type"+t).focus();
				if(loadedcb)loadedcb(id,t);
			};
			w.Show();
			w.Focus();
			return w;
		},
		//��ʾһ��ȷ�Ͽ�,�ش�cb�ĵ�������������:��ȷ��Ϊtrue;ȡ��Ϊfalse
		showConfirm:function(msg,title,cb)
		{
			_me.Common.showAlert(msg,title,cb,"layout/winconfirm.htm","confirm.gif");
		},
		//��ʾһ���ı������,�ش�cb�ĵ�������������:��������
		showPrompt:function(value,title,cb)
		{
			_me.Common.showAlert(value,title,cb,"layout/winPrompt.htm");
		},
		//��ʾһ��������
		showSelect:function(title,loadedcb,closecb)
		{
			_me.Common.showAlert(null,title,closecb,"layout/winSelect.htm",null,loadedcb);
		},
		//��ʾһ�����ڣ���������iframe����ʽ��ʾ������ҳ
		showLink:function(_w,_h,_t,_l,_id,_title,_url,_icon,_closecb)
		{
			var wLink = WinManage.GetWindow(_id,3);//����
			if(wLink)
			{
				if(wLink.isMin)wLink.win.Minimize();
				wLink.win.Focus();
				return;
			}
			var w = new WebForm();
			w.Title = _title||"MYCrm";
			w.Icon = _icon||"defaulticon.gif";
			w.Type = 3;
			w.UserID = _id;
			w.Height = _h;
			w.Width = _w;
			w.Left = _l;
			w.Top = _t;
			w.Resizeable = !!0;
			w.ShowCorner = !!0;
			w.Content = "<iframe scrolling='no' id='ifrId[id]Type[type]' frameborder='0' src='"+_url+"'></iframe>";
			w.RepaintMethod = function(w,h,id,t){$("ifrId"+id+"Type"+t).style.height = h+"px";$("ifrId"+id+"Type"+t).width=w-2+"px";};
			if(_closecb)w.CloseCallback = _closecb;
			w.LoadedCallback = function(){w.ShowLoading();};
			w.Show();
			w.Focus();
		}
	};
	//ģ��
	this.Model = {
		//�û�ģ��
		User:function(_face,_id,_name,_email,_sign,_userStatus,_groupId,_isBlocked,_customname,_gender)
		{
			this.UserFace = _face;
			this.UserID = _id;
			this.UserName =_name;
			this.UserEmail = _email;
			this.UserSign = _sign;
			this.OnlineStatus = _me.Common.getOnlineStatus(_userStatus);
			this.UserStatus = _userStatus;
			this.GroupID = _groupId;
			this.IsBlocked = _isBlocked;
			this.CustomName= _customname||_name;
			this.UserGender= _gender;
		},
		//����ģ��
		Group:function(_name,_id)
		{
			this.Name = _name;
			this.ID = _id;
		},
		//�û�����
		Config:function(_dis,_order,_side,_sendkey,_showtime,_userpower)
		{
			this.DisType    = _dis;//��ʾ��ʽ1Ĭ��2��ͼ��
			this.OrderType  = _order;//����ʽ1״̬2����
			this.ChatSide   = _side;//���촰���Ƿ���ʾͼƬ1��2����
			this.MsgSendKey = _sendkey;//������Ϣ�ȼ�,1enter2ctrl+enter
			this.MsgShowTime = _showtime;//����ʱ��,1��ʾ2����ʾ
			this.UserPower = _userpower; //�û�Ȩ��
		},
		//��Ϣ
		Msg:function(_f,_t,_c,_type,_i,_time)
		{
			this.From = _f; //��Ϣ��Դ
			this.To   = _t; //��Ϣȥ��
			this.Content = _c; //��Ϣ����
			this.Type    = _type;//1�ı�2����3���Ӻ���4ɾ������5״̬�ı�6���
			this.IsConfirm = _i; //��Ϣ�Ƿ�ȷ��
			this.AddTime = _time;//��Ϣ����ʱ��
		},
		//ϵͳ
		Sys:function(_c)
		{
			this.Code = _c;//У����
			this.IntervalTime = 3500; //ѭ�����
			this.IntervalID = null;   //ѭ��ID
		}
	};
	//����
	this.Data = {
		//msn������Ϣ
		getMsnFaceInfo:function()
		{
			return [["1","Ц��","/xl"],["2","Ʋ��","/pz"],["3","ɫ","/se"],["4","����","/fd"],["5","����","/dy"],["6","����","/ll"],["7","����","/hx"],["8","����","/bz"],["9","˯","/shui"],["10","���","/dk"],["11","����","/gg"],["12","��ŭ","/fn"],["13","��Ƥ","/tp"],["14","����","/zy"],["15","����","/jy"],["16","�ѹ�","/ng"],["17","��","/kuk"],["18","�亹","/lengh"],["19","ץ��","/zk"],["20","��","/tu"],["21","͵Ц","/tx"],["22","�ɰ�","/ka"],["23","����","/by"],["24","����","/am"],["25","����","/jie"],["26","��","/kun"],["27","����","/jk"],["28","����","/lh"],["29","��Ц","/hanx"],["30","���","/db"],["31","�ܶ�","/fengd"],["32","����","/zm"],["33","����","/yw"],["34","��...","/xu"],["35","��","/yun"],["36","��ĥ","/zhem"],["37","˥","/shuai"],["38","����","/kl"],["39","�ô�","/qd"],["40","�ټ�","/zj"],["41","����","/ch"],["42","�ٱ�","/kb"],["43","����","/gz"],["44","�ܴ���","/qdl"],["45","��Ц","/huaix"],["46","��ߺ�","/zhh"],["47","�Һߺ�","/yhh"],["48","��Ƿ","/hq"],["49","����","/bs"],["50","ί��","/wq"],["51","�����","/kkl"],["52","����","/yx"],["53","����","/qq"],["54","��","/xia"],["55","����","/kel"]];
		},
		//���ظ�����Ϣ��User����
		getMyUserInfo:function()
		{
			new Ajax().get("data/getmyinfo.asp",
				function(o)
				{
					if(!o)return;
					if(!$T(o.responseXML,"list"))return;
					_me.Profile = _me.Common.getUserFromXml($T(o.responseXML,"list").item(0))[0];
				});
		},
		//���غ�����Ϣ��User����
		getFriendUserInfo:function(uid,cb)
		{
			new Ajax().get("data/getuserinfo.asp?id="+uid,cb);
		},
		//���ظ�������
		getMyConfig:function()
		{
			new Ajax().get("data/getmyconfig.asp",
				function(o)
				{
					if(!o)return;
					if(!$T(o.responseXML,"list"))return;
					_me.Config =_me.Common.getConfigFromXml($T(o.responseXML,"list").item(0));
				});
		},
		//���غ����б���User��������
		getMyFriendList:function(cb)
		{
			new Ajax().get("data/getmyfriend.asp",
				function(o)
				{
					if(!o)return;
					if(!$T(o.responseXML,"list"))return;
					_me.Friend =_me.Common.getUserFromXml($T(o.responseXML,"list").item(0));
					if(cb)cb(o);
				});
		},
		//���ط����б���Group��������
		getMyGroupList:function(cb)
		{
			new Ajax().get("data/getmygroup.asp",
				function(o)
				{
					if(!o)return;
					if(!$T(o.responseXML,"list"))return;
					_me.Group = _me.Common.getGroupFromXml($T(o.responseXML,"list").item(0));
					if(cb)cb(o);
				});
		},
		//������Ϣ�б�msg��������
		getMyMsgList:function(cb)
		{
			new Ajax().get("data/getmymsg.asp?code="+_me.Sys.Code,cb);
		},
		//�޸��û�Profile
		setUserProfile:function()
		{
			var data = "userface="+_me.Profile.UserFace;
			data+="&username="+_me.Profile.UserName.escapeEx();
			data+="&usersign="+_me.Profile.UserSign.escapeEx();
			data+="&userstatus="+_me.Profile.UserStatus;
			new Ajax().post("data/service.asp?t=4",data);
		},
		//��½
		setUserLogin:function(email,pass,us,cb)
		{
			new Ajax().post("data/service.asp?t=0","us="+us+"&email="+email+"&pass="+pass,cb);
		},
		//�˳���½
		setUserLogout:function()
		{
			new Ajax().post("data/service.asp?t=2");
		},
		//������Ϣ
		sendMessage:function(msg)
		{
			new Ajax().post("data/service.asp?t=3","from="+msg.From+"&to="+msg.To+"&content="+msg.Content.escapeEx()+"&type="+msg.Type);
		},
		//�������Ӻ�������
		acceptAddFriend:function(uid,cb)
		{
			new Ajax().post("data/service.asp?t=5","to="+uid,cb);
		},
		//ɾ������
		deleteFriend:function(uid)
		{
			new Ajax().post("data/service.asp?t=6","to="+uid);
		},
		//�޸�����״̬
		blockFriend:function(uid,status)
		{
			new Ajax().post("data/service.asp?t=7","to="+uid+"&s="+status);
		},
		//�޸ĺ����ǳ�
		editCustomName:function(uid,name)
		{
			new Ajax().post("data/service.asp?t=8","to="+uid+"&n="+name.escapeEx());
		},
		//�޸ĺ��ѷ���
		editUserGroup:function(uid,gid)
		{
			new Ajax().post("data/service.asp?t=12","id="+uid+"&gid="+gid);
		},
		//������
		addGroup:function(gname,cb)
		{
			new Ajax().post("data/service.asp?t=9","n="+gname.escapeEx(),cb);
		},
		//ɾ��
		delGroup:function(gid,cb)
		{
			new Ajax().post("data/service.asp?t=10","id="+gid,cb);
		},
		//�޸�
		editGroup:function(gid,gname)
		{
			new Ajax().post("data/service.asp?t=11","id="+gid+"&n="+gname.escapeEx());
		}
	};
	//����
	this.CMD = {
		//���ӷ���
		addGroup:function()
		{
			_me.Common.showPrompt("","����������",
				function()
				{
					var gname = arguments[2].strip();
					if(gname=="")return;
					if(!_me.Common.getGroupByName(gname))//������
					{
						_me.Data.addGroup(gname,
						function(o)
						{
							_me.Data.getMyGroupList(
								function(o)
								{
									_me.Data.getMyFriendList(function(){_me.CMD.renderMyFriend(null,null,!0);});
								});
						});
					}
					else
					{
						_me.Common.showAlert("��������Ѿ�����!","��ʾ");
					}
				});
		},
		//�༭����
		editGroup:function(gid)
		{
			if(gid==1)
			{
				_me.Common.showAlert("����������Ĭ�Ϸ���!","��ʾ");
				return;
			}
			var g = _me.Common.getGroupById(gid);
			if(!g)return;
			_me.Common.showPrompt(g.Name,"����������",
				function()
				{
					var gname = arguments[2].strip();
					if(gname==g.Name||gname=="")return;
					if(!_me.Common.getGroupByName(gname))//������
					{
						g.Name = gname;
						_me.Data.editGroup(gid,gname);
						_me.CMD.renderMyFriend();
					}
					else
					{
						_me.Common.showAlert("��������Ѿ�����!","��ʾ");
					}
				});
		},
		//ɾ������
		delGroup:function(gid)
		{
			if(gid==1)
			{
				_me.Common.showAlert("����ɾ��Ĭ�Ϸ���!","��ʾ");
				return;
			}
			var g = _me.Common.getGroupById(gid);
			_me.Common.showConfirm("ȷ��Ҫɾ����"+g.Name+"���������(������ѻ�����Ĭ����)��","��ʾ",
				function()
				{
					if(arguments[2])
					{
						_me.Data.delGroup(gid,
							function(o)
							{
								_me.Data.getMyGroupList(
									function(o)
									{
										_me.Data.getMyFriendList(function(){_me.CMD.renderMyFriend(null,null,!0);});
									});
							});
					}
				});
		},
		//��ʾ��ϵ�˿�Ƭ
		//�޸� showCard:function(uid,o)
		showCard:function(uid1,o)
		{
			var u = uid==_me.Profile.UserID?_me.Profile:_me.Common.getUserFromArr(uid,_me.Friend);
			if(!u)return;
			var wWin = _me.Win;
			var top = Elem.GetY($(o),$("windowContainer"))+2-$(o).parentNode.parentNode.parentNode.scrollTop;
			var left = wWin.win.Left<233?wWin.win.Left+wWin.win.Width+2:wWin.win.Left-230;
			var oCard = $("divCardBorder");
			if(!oCard)
			{
				oCard = Elem.New("div","divCardBorder","wCardBorder");
				Elem.Add("windowContainer",oCard);
			}
			var os = oCard.style;
			os.top = top+"px";
			os.left = left+"px";
			Elem.Value(oCard);
			var oCardContent = Elem.New("div","","wCardContainer",_me.Common.createUserItem(3,u));
			var ban = u.IsBlocked?"b":"";
			var img = Elem.New("img","","wCardCloseImage");
			img.src = "images/close.gif";
			img.title = "�ر�";
			img.onmouseover = function(){this.src="images/closehover.gif";};
			img.onmouseout = function(){this.src="images/close.gif";};
			img.onclick = function(){Elem.Hid("divCardBorder");};
			Elem.Add(oCardContent,img);
			Elem.Add(oCard,oCardContent);
			Elem.Show(oCard);
		},
		//�޸��û�����
		eidtUserGroup:function(uid)
		{
			var u = _me.Common.getUserFromArr(uid,_me.Friend);
			if(!u)return;
			_me.Common.showSelect("��ѡ�����",
				function(id,t)
				{
					var oSel = $("divSelectId"+id+"Type"+t);
					for(var i=0;i<_me.Group.length;i++)
					{
						var g = _me.Group[i];
						var opt = Elem.New("option");
						opt.value = g.ID;
						opt.text = g.Name;
						oSel.options.add(opt);
					}
					Elem.Value(oSel,u.GroupID);
				},
				function()
				{
					var gid = arguments[2];
					if(gid==""||gid==u.GroupID)return;//�޸ı�
					_me.Data.editUserGroup(uid,gid);
					u.GroupID = gid;
					_me.CMD.renderMyFriend(null,null,!0);
				});
		},
		//�޸ĺ����ǳ�
		editCustomName:function(uid)
		{
			var u = _me.Common.getUserFromArr(uid,_me.Friend);
			if(!u)return;
			_me.Common.showPrompt(u.CustomName,"�������ǳ�",
				function()
				{
					var name = arguments[2].trim();
					if(name!="")
					{
						if(name==u.CustomName)return;//�޸ı�
						_me.Data.editCustomName(uid,name);
						u.CustomName = name;
						_me.CMD.renderMyFriend();
					}
					else//����ǳ�
					{
						_me.Data.editCustomName(uid,"");
						u.CustomName = u.UserName;
						_me.CMD.renderMyFriend();
					}
				});
		},
		//��ʾ�����¼
		showMsgHistory:function(uid)
		{
			var url = "page/Message.asp?v="+Math.random();
			if(uid)url+="&id="+uid;
			_me.Common.showLink(550,450,50,(Elem.Width()-450)/2,9,"�����¼",url,"toolmsghistory.gif");
		},
		//��ʾ����ҳ��
		showManage:function(uid)
		{
			var url = "page/usermanage.asp?v="+Math.random();
			_me.Common.showLink(550,450,50,(Elem.Width()-450)/2,12,"����",url,"toolmanage.gif");
		},
		//���κ���
		blockFriend:function(uid,isblock)
		{
			_me.Data.blockFriend(uid,isblock);
			var users = [];
			users = _me.Friend;
			var u = _me.Common.getUserFromArr(uid,users);
			u.IsBlocked = isblock==1;
			if($("wChatButtonBlockId"+uid+"Type1"))
			{
				var oBtn = $("wChatButtonBlockId"+uid+"Type1");
				var isblock = 3-parseInt(oBtn.getAttribute("b"));
				$T(oBtn,"img")[0].src = isblock == 1?"images/chatbuttoncancelblock.gif":"images/chatbuttonblock.gif";
				$T(oBtn,"img")[0].title = isblock == 1?"ȡ����ֹ����ϵ��":"��ֹ����ϵ��";
				oBtn.setAttribute("b",isblock);
			}
			_me.CMD.renderMyFriend();
		},
		//ɾ������
		delFriend:function(uid)
		{
			var users = [];
			users = _me.Friend;
			var u = _me.Common.getUserFromArr(uid,users);
			_me.Common.showConfirm("ȷ��Ҫɾ����"+u.CustomName+"����λ���ѣ�","��ʾ",
			function()
			{
				if(arguments[2])
				{
					var wChat = WinManage.GetWindow(uid,1);
					if(wChat)wChat.win.Close();
					_me.Friend.remove(u);
					_me.CMD.renderMyFriend();
					_me.Data.deleteFriend(uid);
				}
			});
		},
		//ѭ����ȡ��Ϣ
		getMsgInterval:function()
		{
			if(!_me.Sys.IntervalID)
			{
				_me.Sys.IntervalID = setInterval(
					function()
					{
						_me.Data.getMyMsgList(
							function(o)
							{
								if(!o)return;
								if(!$T(o.responseXML,"list"))return;
								var msgs = _me.Common.getMsgFromXml($T(o.responseXML,"list").item(0));
								if(msgs)
								{
									for(var i=0;i<msgs.length;i++)
									{
										_me.CMD.showChatContent(msgs[i].From,msgs[i]);
									}
								}
							});
					}
					,_me.Sys.IntervalTime);
			}
		},
		//ֹͣѭ����ȡ��Ϣ
		stopMsgInterval:function()
		{
			clearInterval(_me.Sys.IntervalID);
			_me.Sys.IntervalID = null;
		},
		//��ʾע�ᴰ��
		showRegWindow:function()
		{
			var wLogin = WinManage.GetWindow(6,3);
			if(wLogin)wLogin.win.Minimize();
			_me.Common.showLink(350,320,50,(Elem.Width()-350)/2,7,"ע�����ʺ�","page/reg.asp?v="+Math.random(),null,
				function()
				{
					if(wLogin.isMin)wLogin.win.Minimize();
				});
		},
		//��½���ڳ�ʼ��
		intLoginWindow:function()
		{
			var wMain = WinManage.GetWindow(0,2);//����������
			if(wMain)
			{
				if(wMain.isMin)wMain.win.Minimize();
				wMain.win.Focus();
				return;
			}
			var wLogin = WinManage.GetWindow(6,3);//���ڵ�½����
			if(wLogin)
			{
				if(wLogin.isMin)wLogin.win.Minimize();
				wLogin.win.Focus();
				return;
			}
			var w = new WebForm();
			w.Title = "�ڲ�������";
			w.Icon = "defaulticon.gif";
			w.Type = 3;
			w.UserID = 6;
			w.Height = 450;
			w.Width = 210;
			w.Left = (Elem.Width()-240)/1;
			w.ContentUrl = "layout/winlogin.htm";
			w.Resizeable = !!0;
			w.RepaintMethod = function(w,h,id,t){$("wOtherMainId"+id+"Type"+t).style.height = h+"px";};
			w.LoadedCallback = function(id,t)
			{
				$("divStatusId6Type3").onclick = function(e)
				{
					var m = new SysMenu("menuLoginStatus6Type3");
					m.Data = ["����|_webIM.CMD.changeLoginStatus(0)|images/m0.gif",
						"æµ|_webIM.CMD.changeLoginStatus(1)|images/m1.gif",
						"���ϻ���|_webIM.CMD.changeLoginStatus(2)|images/m2.gif",
						"�뿪|_webIM.CMD.changeLoginStatus(3)|images/m2.gif",
						"ͨ����|_webIM.CMD.changeLoginStatus(4)|images/m1.gif",
						"����Ͳ�|_webIM.CMD.changeLoginStatus(5)|images/m2.gif",
						"����|_webIM.CMD.changeLoginStatus(6)|images/m3.gif"];
					m.E = e;
					m.HasIcon = true;
					m.Width = 120;
					m.Top = w.Top+157;
					m.Left = w.Left+102;
					m.Show();
				};
				Elem.Value("tbEmailId6Type3",Other.GetCookie("stremail"));
				//Elem.Value("tbPassId6Type3",Other.GetCookie("strpass"));
				$("cbSaveUserId6Type3").checked = Other.GetCookie("saveemail")=="1";
				$("cbSavePassId6Type3").checked = Other.GetCookie("savepass")=="1";
				$("cbAutoLoginId6Type3").checked = Other.GetCookie("autologin")=="1";
				if(Other.GetCookie("loginstatus").trim()!="")_webIM.CMD.changeLoginStatus(parseInt(Other.GetCookie("loginstatus")));
				new CheckBox("cbSaveUserId6Type3","�����ҵ���Ϣ(B)").Render();
				new CheckBox("cbSavePassId6Type3","��ס�ҵ�����(R)").Render();
				new CheckBox("cbAutoLoginId6Type3","�Զ�Ϊ�ҵ�½(N)").Render();
				$("tbPassId6Type3").onkeydown = function(e)
				{
					var e = e||event;
					if(e.keyCode==13)//�س�
					{
						_me.CMD.loginWebIM();
					}
				}
				$("btnLoginId6Type3").onclick=_me.CMD.loginWebIM;
				$("linkDelCookieId6Type3").onclick = function()
				{
					Other.SetCookie("stremail","");
					Other.SetCookie("strpass","");
					Other.SetCookie("saveemail","2");
					Other.SetCookie("savepass","2");
					Other.SetCookie("autologin","2");
					Other.SetCookie("loginstatus","0");
					_me.CMD.destroyMainWindow();
					IntWebIM();
				};
				if(Other.GetCookie("autologin")=="1")
				{
					_me.CMD.loginWebIM();//�Զ���½
				}
			};
			w.CloseCallback = _me.CMD.destroyMainWindow;
			w.Show();
			w.Focus();
		},
		//��½
		loginWebIM:function()
		{
			var w = WinManage.GetWindow(6,3).win;
			var user = $F("tbEmailId6Type3").trim();
			var pass = $F("tbPassId6Type3").trim();
			if(user=="")
			{
				_me.Common.showAlert("��������ĵ����ʼ�!","��ʾ",
					function()
					{
						w.Focus();
						$("tbEmailId6Type3").focus();
					});
				return;
			}
			//if(pass=="")
			//{
			//	_me.Common.showAlert("�������������!","��ʾ",
			//		function()
			//		{
			//			w.Focus();
			//			$("tbPassId6Type3").focus();
			//		});
			//	return;
			//}
			var us = $("divStatusId6Type3").getAttribute("us");
			Elem.Toggle("divLoginId6Type3","divLoginingId6Type3");
			_me.Data.setUserLogin(user,pass,us,function(o)
			{
				var result = parseInt(Xml.First($T(o.responseXML,"result").item(0),"num"));
				if(result==1)
				{
					Other.SetCookie("stremail",$("cbSaveUserId6Type3").checked?user:"");
					Other.SetCookie("strpass",$("cbSavePassId6Type3").checked?pass:"");
					Other.SetCookie("saveemail",$("cbSaveUserId6Type3").checked?"1":"2");
					Other.SetCookie("savepass",$("cbSavePassId6Type3").checked?"1":"2");
					Other.SetCookie("autologin",$("cbAutoLoginId6Type3").checked?"1":"2");
					Other.SetCookie("loginstatus",us);
					w.Close();
					_me.Sys.Code = parseInt(Xml.First($T(o.responseXML,"result").item(0),"code"));
					_me.CMD.intMainWindow(w.Top,w.Left);
				}
				else
				{
					w.Focus();
					$("tbEmailId6Type3").focus();
					_me.Common.showAlert("��½ʧ��,����������ʻ���Ϣ�Ƿ���ȷ!","��ʾ");
					Elem.Toggle("divLoginId6Type3","divLoginingId6Type3");
					return;
				}
			});
		},
		//�������ʼ��
		intMainWindow:function(top,left)
		{
			var w = new WebForm();
			w.Title = "<span style='color:#fff;padding-left:5px;'>�ڲ�������</span>";
			w.Icon = "defaulticon.gif";
			w.Type = 2;
			w.Height = 450;
			w.Width = 220;
			w.Top = top||null;
			w.Left = left||2600;
			w.MinHeight = 250;
			w.RepaintMethod = function(w,h,id,t)
			{
				$("wMainMainId"+id+"Type"+t).style.height = (h-150)+"px";
				$("wMainUserInfoRightId"+id+"Type"+t).style.width = (w-82)+"px";
				$("wMainSearchUserId"+id+"Type"+t).style.width = (w-65)+"px";
				$("inputSearchId"+id+"Type"+t).style.width = (w-80)+"px";
				$("txtUserSignId"+id+"Type"+t).style.width = (w-85)+"px";
			};
			w.ClosingCallback = function()
			{
				_me.CMD.stopMsgInterval();//��ͣ��Ϣ��ʱ��
						//�޸�_me.Data.setUserLogout();
						_me.Win.win.ClosingCallback = null;
						_me.Win.win.Close();
				//_me.Common.showConfirm("���Ҫ�˳���","��ʾ",
				//function()
				//{
				//	if(arguments[2])
				//	{
				//		_me.Data.setUserLogout();
				//		_me.Win.win.ClosingCallback = null;
				//		_me.Win.win.Close();
				//	}
				//	else
				//	{
				//		_me.CMD.getMsgInterval();//����Ϣ��ʱ��
				//	}
				//});
			};
			w.CloseCallback = function()
			{
				_me.CMD.destroyMainWindow();
			};
			w.LoadedCallback = function(id,t)
				{
					w.ShowLoading();
					_me.Win = WinManage.GetWindow(id,t);
					_me.Data.getMyConfig();
					_me.Data.getMyUserInfo();
					_me.Data.getMyGroupList();
					_me.Data.getMyFriendList();
					var intInteralID = setInterval(
						function()
						{
							if(_me.Config&&_me.Profile&&_me.Group&&_me.Friend)
							{
								clearInterval(intInteralID);
								_me.CMD.renderMyUserInfo();
								_me.CMD.renderMyFriend();
								_me.CMD.getMsgInterval();
								$("wMainUserSignId"+id+"Type"+t).onclick = function()
								{
									Elem.Show("txtUserSignId"+id+"Type"+t);
									Elem.Value("txtUserSignId"+id+"Type"+t,_me.Profile.UserSign);
									$("txtUserSignId"+id+"Type"+t).focus();
									$("txtUserSignId"+id+"Type"+t).onblur = function()
									{
										var sign = $F(this).trim();
										if(sign!=""&&_me.Profile.UserSign!=sign)//ǩ���и��ľ���Ҫupdate
										{
											_me.Profile.UserSign = sign;
											_me.CMD.renderMyUserInfo();
											_me.Data.setUserProfile();
										}
										Elem.Hid("txtUserSignId"+id+"Type"+t);
									};
								};
								$("wMainMainId"+id+"Type"+t).onclick = function(e)
								{
									var e = e||window.event;
									var tar = e.srcElement||e.target;
									if(tar)
									{

									}
								};
								$("wMainUserNameId"+id+"Type"+t).parentNode.onclick = function(e)//����״̬�˵�
								{
									var m = new SysMenu("menuUserNameId"+id+"Type"+t);
									m.Data = ["����|_webIM.CMD.changeUserStatus(0)|images/m0.gif",
											"æµ|_webIM.CMD.changeUserStatus(1)|images/m1.gif",
											"���ϻ���|_webIM.CMD.changeUserStatus(2)|images/m2.gif",
											"�뿪|_webIM.CMD.changeUserStatus(3)|images/m2.gif",
											"ͨ����|_webIM.CMD.changeUserStatus(4)|images/m1.gif",
											"����Ͳ�|_webIM.CMD.changeUserStatus(5)|images/m2.gif",
											"����|_webIM.CMD.changeUserStatus(6)|images/m3.gif",
											"",
											//"ע��|_webIM.CMD.logoutWebIM()|images/m3.gif"];
											"ע��|_webIM.CMD.changeUserStatus(7)|images/m3.gif"];
									m.E = e;
									m.HasIcon = true;
									m.Width = 120;
									m.Top = w.Top+59;
									m.Left = w.Left+80;
									m.Show();
								};
								$("inputSearchId"+id+"Type"+t).onfocus = function()//��������
								{
									if(parseInt(this.getAttribute("b"))==1)return;
									var value = $F(this);
									var IntID = setInterval(
										function()
										{
											if(value != $F("inputSearchId"+id+"Type"+t))
											{
												value = $F("inputSearchId"+id+"Type"+t);
												_me.CMD.searchMyFriend(value);
											}
										},50);
									this.setAttribute("b","1");
									this.onblur = function()
									{
										this.setAttribute("b","2");
										clearInterval(IntID);
									};
								};
								$("imgSearchId"+id+"Type"+t).onclick = function()//�������ͼƬ
								{
									Elem.Hid("imgSearchId"+id+"Type"+t);
									Elem.Value("inputSearchId"+id+"Type"+t);
									_me.CMD.renderMyFriend();
								};
								$("wMainAddFriendId"+id+"Type"+t).onclick = function()//������ϵ�˴���
								{
									_me.Common.showLink(350,285,w.Top,w.Left<355?w.Left+w.Width+2:w.Left-352,1,"������ϵ��","page/addfriend.asp?v="+Math.random(),"tooladdfriend.gif");
								};
								$("wMainOptionId"+id+"Type"+t).onclick = function()//ѡ���
								{
									_me.Common.showLink(420,400,w.Top,w.Left<425?w.Left+w.Width+2:w.Left-422,8,"ѡ��","page/option.asp?v="+Math.random(),"tooloption.gif");
								};
								$("wMainProfileId"+id+"Type"+t).onclick = function()//������Ϣ����
								{
									_me.Common.showLink(350,300,w.Top,w.Left<355?w.Left+w.Width+2:w.Left-352,10,"�༭��������","page/profile.asp?v="+Math.random(),"toolprofile.gif");
								};
								$("wMainShowHistoryId"+id+"Type"+t).onclick = function()//�����¼����
								{
									_me.CMD.showMsgHistory();
								};
								$("wMainShowFocusId"+id+"Type"+t).onclick = function()//���ս��㴰��
								{
									_me.Common.showLink(400,400,w.Top,w.Left<405?w.Left+w.Width+2:w.Left-402,2,"���ս���","page/todayfocus.asp?v="+Math.random(),"toolshowfocus.gif");
								};
								$("wMainUserFaceId"+id+"Type"+t).onclick = function()//��ʾ�Լ��Ŀ�Ƭ
								{
									_me.CMD.showCard(_me.Profile.UserID,this);
								}
								$("wMainListTypeId"+id+"Type"+t).onclick = function(e)//������ϵ�˰�ť
								{
									var m = new SysMenu("menuUserNameId"+id+"Type"+t);
									m.Data = ["��״̬����|_webIM.CMD.renderMyFriend(1,_webIM.Config.DisType,true)",
											"��������|_webIM.CMD.renderMyFriend(2,_webIM.Config.DisType,true)",
											"",
											"��ʾ��Ҫ��Ϣ|_webIM.CMD.renderMyFriend(_webIM.Config.OrderType,1,true)"];
											//�޸Ĺ����� ,"��ʾ��ϸ��Ϣ|_webIM.CMD.renderMyFriend(_webIM.Config.OrderType,2,true)"
									m.Data[_me.Config.OrderType-1]+="|images/selected.gif|images/selectedhover.gif";
									m.Data[_me.Config.DisType+2]+="|images/selected.gif|images/selectedhover.gif";
									m.E = e;
									m.HasIcon = true;
									m.Width = 120;
									m.Top = w.Top+157;
									m.Left = w.Left+w.Width-23;
									m.Show();
								};
								if(_me.Config.UserPower<2)
								{
									var o = $("wMainManageId"+id+"Type"+t);
									Elem.Show(o);
									o.onclick = _me.CMD.showManage;
								}
								if(!$("divFaceList"))//�����������
								{
									var oDiv = Elem.New("div","divFaceList","wChatFaceContainer");
									var faces = [];
									faces = _me.Data.getMsnFaceInfo();
									for(var i=0;i<faces.length;i++)
									{
										var face = faces[i];
										var oItem = Elem.New("div","","wChatFaceItem","<img f='"+face[2]+"' src='qqface/"+face[0]+".gif' title='"+face[1]+" "+face[2]+"'/>");
										oItem.onmouseover = function(){this.className="wChatFaceItem wChatFaceItemHover";};
										oItem.onmouseout  = function(){this.className="wChatFaceItem";};
										oItem.onclick = function()
										{
											var chatid = parseInt(window.cWindow.replace("winChat",""));
											if(!chatid)return;
											Elem.Append("wChatInputId"+chatid+"Type1",$T(this,"img")[0].getAttribute("f"));
											Elem.Hid("divFaceList");
										};
										Elem.Add(oDiv,oItem);
									}
									Elem.Add("windowContainer",oDiv);
								}
								w.HideLoading();
							}
						},50);
				};
			w.Show();
		},
		//ע������
		logoutWebIM:function()
		{
			_me.CMD.stopMsgInterval();//��ͣ��Ϣ��ʱ��
			Other.SetCookie("autologin","2");
			_me.Data.setUserLogout();
			_me.Win.win.ClosingCallback = null;
			_me.Win.win.Close();
			IntWebIM();
		},
		//��䱾����Ϣ
		renderMyUserInfo:function()
		{
			var user = _me.Profile;
			var id = _me.Win.id;
			var t = _me.Win.type;
			$("wMainUserFaceId"+id+"Type"+t).src = "images/userface/"+user.UserFace;
			$("wMainUserNameId"+id+"Type"+t).innerHTML = user.UserName+" <span style=\"font-size:12px;color:#B9DDE7\">("+_me.Common.getUserStatusStr(user.UserStatus)+")</span>";
			$("wMainUserSignId"+id+"Type"+t).innerHTML = user.UserSign;
		},
		//�޸�UserStatus
		changeUserStatus:function(n)
		{
			if(_me.Profile.UserStatus == n)return;//�����޸�
			_me.Profile.UserStatus = n;
			_me.Profile.OnlineStatus = _me.Common.getOnlineStatus(n);
			_me.CMD.renderMyUserInfo();
			_me.Data.setUserProfile();
		},
		//�޸ĵ�½״̬
		changeLoginStatus:function(n)
		{
			$("divStatusId6Type3").innerHTML = _me.Common.getUserStatusStr(n);
			$("divStatusId6Type3").setAttribute("us",n);
		},
		//���������壬������Ӵ���
		destroyMainWindow:function()
		{
			clearInterval(WinManage.WinListInteralID);
			for(var f = 0 ;f < WinManage.WindowsList.length;f++)
			{
				var w = WinManage.WindowsList[f];
				Elem.Del(WinManage.GetObjByWindow(w));
			}
			WinManage.WindowsList.clear();
			Elem.Del("windowContainerBorder");
			window.onscroll = window.onresize = null;
		},
		//��������
		searchMyFriend:function(k)
		{
			var id = _me.Win.id;
			var t = _me.Win.type;
			if(k=="")
			{
				Elem.Hid("imgSearchId"+id+"Type"+t);
				_me.CMD.renderMyFriend();
			}
			else
			{
				var obj = $("wMainMainId"+id+"Type"+t);
				Elem.Show("imgSearchId"+id+"Type"+t);
				Elem.Value(obj);//���
				var result = _me.Common.searchFriendList(k);
				if(result&&result.length>0)
				{
					var title = Elem.New("div","","","<div class=\"wMainUserItemText\" style=\"padding-left:20px;width:100%;color:#aca899\">�ҵ�:</div>");
					title.style.height = "20px";
					var strHtml = new StringBuilder();
					for(var i=0;i<result.length;i++)
					{
						var u = result[i];
						strHtml.add(_me.Common.createUserItem(_me.Config.DisType,u));
					}
					var container = Elem.New("div","wMainUserContainerId"+id+"Type"+t+"Search","wMainUserContainer",strHtml.toString());
					Elem.Add(obj,title,container);
				}
				else
				{
					Elem.Value(obj,"<div class=\"wMainUserItemText\" style=\"padding:10px 0 0 10px;width:100%;\">���������</div>");
				}
			}
		},
		//�������б�
		renderMyFriend:function(orderType,disType,clearSearch)
		{
			var id = _me.Win.id;
			var t = _me.Win.type;
			var searchKey = $F("inputSearchId"+id+"Type"+t);
			if(searchKey!="") //������Ϊ��
			{
				if(clearSearch) //���
				{
					Elem.Value("inputSearchId"+id+"Type"+t);
					Elem.Hid("imgSearchId"+id+"Type"+t);
				}
				else
				{
					_me.CMD.searchMyFriend(searchKey);
					return;
				}
			}
			if(orderType&&disType)
			{
				if(orderType!=_me.Config.OrderType||disType!=_me.Config.DisType)
				{
					_me.Config.OrderType = orderType;
					_me.Config.DisType = disType;
				}
				else
				{
					return;//����Ҫ�������
				}
			}
			_me.Common.sortFriendList();
			var obj = $("wMainMainId"+id+"Type"+t);
			var nullGroupMsg = "<div class=\"wMainUserItemText\" style=\"padding-left:25px;width:100%;color:#aca899\">������û����ϵ��</div>";
			Elem.Value(obj);
			var itemHeight = _me.Config.DisType==1?30:63;
			var groupUsers = {};
			if(_me.Config.OrderType == 1)
			{
				_me.Common.addUserHeader(obj,"����",0);
				_me.Common.addUserHeader(obj,"����",1);
				var num = [0,0];
				for(var i=0;i<_me.Friend.length;i++)
				{
					var u = _me.Friend[i];
					var no = u.OnlineStatus==3?1:0;
					if(!(no in groupUsers))groupUsers[no]=[];
					groupUsers[no].add(u);
					num[no]++;
				}
				for(var i=0;i<num.length;i++)
				{
					if(num[i]==0)
					{
						Elem.Value("wMainUserContainerId"+id+"Type"+t+"No"+i,nullGroupMsg);
						$("wMainUserContainerId"+id+"Type"+t+"No"+i).style.height = "30px";
						continue;
					}
					var strHtml = new StringBuilder();
					for(var q=0;q<groupUsers[i].length;q++)
					{
						strHtml.add(_me.Common.createUserItem(_me.Config.DisType,groupUsers[i][q]));
					}
					Elem.Value($("wMainUserContainerId"+id+"Type"+t+"No"+i),strHtml.toString());
					$("wMainUserHeaderId"+id+"Type"+t+"No"+i).innerHTML +=" ( "+num[i]+" ) ";
					$("wMainUserContainerId"+id+"Type"+t+"No"+i).style.height = num[i]*itemHeight+"px";
				}
			}
			else
			{
				var num1 = {},num2 = {};
				for(var i=0;i<_me.Group.length;i++)
				{
					var g = _me.Group[i];
					_me.Common.addUserHeader(obj,g.Name,g.ID);
					groupUsers[g.ID]=[];
					num1[g.ID]=0;
					num2[g.ID]=0;
				}
				for(var i=0;i<_me.Friend.length;i++)
				{
					var u = _me.Friend[i];
					groupUsers[u.GroupID].add(u);
					num1[u.GroupID]++;
					if(u.OnlineStatus!=3)num2[u.GroupID]++;
				}
				for(var i=0;i<_me.Group.length;i++)
				{
					var gid = _me.Group[i].ID;
					if(num1[gid]==0)
					{
						Elem.Value("wMainUserContainerId"+id+"Type"+t+"No"+gid,nullGroupMsg);
						$("wMainUserContainerId"+id+"Type"+t+"No"+gid).style.height = "20px";
						continue;
					}
					var strHtml = new StringBuilder();
					for(var q=0;q<groupUsers[gid].length;q++)
					{
						strHtml.add(_me.Common.createUserItem(_me.Config.DisType,groupUsers[gid][q]));
					}
					Elem.Value($("wMainUserContainerId"+id+"Type"+t+"No"+gid),strHtml.toString());
					$("wMainUserHeaderId"+id+"Type"+t+"No"+gid).innerHTML +=" ( "+num2[gid]+" / "+num1[gid]+" ) ";
					$("wMainUserContainerId"+id+"Type"+t+"No"+gid).style.height = num1[gid]*itemHeight+"px";
				}
			}
		},
		//��ʾһ����Ϣ,uid������userid,msg��Ϣ��
		showChatContent:function(uid,msg)
		{
			_me.CMD.stopMsgInterval();
			var users = [];
			users = _me.Friend;
			users.add(_me.Profile);
			var userFrom = _me.Common.getUserFromArr(msg.From,users);
			var userTo = _me.Common.getUserFromArr(msg.To,users);
			users.remove(_me.Profile);
			if(msg.Type<3)
			{
				_me.CMD.openChatWindow(uid,false);//����uid���촰��
				var _interID = setInterval(function()
						{
							var objChat = $("wChatViewId"+uid+"Type1");
							if(objChat)
							{
								clearInterval(_interID);
								var winChat = WinManage.GetWindow(uid,1);
								winChat.win.Flash();//��ʾ������Ϣ
								switch(parseInt(msg.Type))
								{
									case 1://�ı���Ϣ
										var msgTitle = userFrom.CustomName+" ˵";
										if(_me.Config.MsgShowTime==1)
										{
											var msgTime = !msg.AddTime?new Date():new Date(Date.parse(msg.AddTime.replace(/-/g,"/")));
											msgTitle+=" ("+msgTime.getHours().toString().padLeft("0",2)+"��"+msgTime.getMinutes().toString().padLeft("0",2)+")";
										}
										msgTitle +="��";
										Elem.Add(objChat,Elem.New("div",null,"wChatMsgTitle",msgTitle));
										Elem.Add(objChat,Elem.New("div",null,"wChatMsgContent",_me.Common.replaceFaceFromStr(msg.Content).replace(/{br}/img,"<br />")));
										if(msg.From!=_me.Profile.UserID)_me.Common.playSound("newmessage");
										_me.CMD.getMsgInterval();
										break;
									case 2://������Ϣ
										switch(msg.Content)
										{
											case "FLASH"://����
												var oChild = objChat.childNodes;
												if(oChild.length==0||(oChild.length>0&&oChild[oChild.length-1].className!="wChatMsgSplit"))Elem.Add(objChat,Elem.New("div",null,"wChatMsgSplit"));
												Elem.Add(objChat,Elem.New("div",null,"wChatMsgSpecial",userFrom.CustomName+"������һ��������"),Elem.New("div",null,"wChatMsgSplit"));
												WinManage.FlashWindow(WinManage.GetWindow(uid,1));
												_me.Common.playSound("flash");
											break;
										}
										_me.CMD.getMsgInterval();
								}
								objChat.scrollTop = objChat.scrollHeight;
							}
						},100);
			}
			else
			{
				switch(parseInt(msg.Type))
				{
					case 3://���Ӻ���
						_me.Data.getFriendUserInfo(msg.From,
							function(o)
							{
								if(!o)return;
								if(!$T(o.responseXML,"list"))return;
								var profile = _me.Common.getUserFromXml($T(o.responseXML,"list").item(0))[0];
								if(!_me.Common.getUserFromArr(msg.From,users))//��ǰ�б��в�����������
								{
									_me.Friend.add(profile);
									_me.CMD.renderMyFriend();
								}
								_me.CMD.getMsgInterval();
							});
						break;
					case 4://ɾ������
						_me.Friend.remove(userFrom);
						_me.CMD.renderMyFriend();
						_me.CMD.getMsgInterval();
						break;
					case 5://�û���Ϣ�ı�
						_me.Friend.remove(userFrom);
						_me.Data.getFriendUserInfo(msg.From,
							function(o)
							{
								if(!o)return;
								if(!$T(o.responseXML,"list"))return;
								var result =  _me.Common.getUserFromXml($T(o.responseXML,"list").item(0));
								if(!result||result.length<1)return;
								var profile =result[0];
								_me.Friend.add(profile);
								_me.CMD.renderMyFriend();
								_me.Common.playSound("friendonline");
								_me.CMD.getMsgInterval();
							});
						break;
					case 7://�Ӻ�������
						_me.Common.showConfirm(msg.Content+"�������Ϊ����!ͬ�⣿","��ʾ",
							function()
							{
								if(arguments[2])
								{
									_me.Data.acceptAddFriend(msg.From,function()
									{
										_me.Data.getFriendUserInfo(msg.From,
										function(o)
										{
											if(!o)return;
											if(!$T(o.responseXML,"list"))return;
											var profile = _me.Common.getUserFromXml($T(o.responseXML,"list").item(0))[0];
											_me.Friend.add(profile);
											_me.CMD.renderMyFriend();
										});
									});
								}
								_me.CMD.getMsgInterval();//����Ϣ��ʱ��
							});
						break;
					case 8://�˳�
						_me.Common.showAlert(msg.Content,"��ʾ",
							function()
							{
								_me.CMD.destroyMainWindow();
							});
				}
			}
		},
		//������Ϣ
		sendMessage:function(uid)
		{
			var objContent = $("wChatInputId"+uid+"Type1");
			var msg = "";
			if($F(objContent).trim()=="")
			{
				msg = "���ܷ��Ϳ���Ϣ!";
			}
			var users = [];
			users = _me.Friend;
			var userTo = _me.Common.getUserFromArr(uid,users);
			if(userTo&&userTo.IsBlocked)
			{
				msg = "�����������κ��ѷ�����Ϣ��";
				Elem.Value(objContent);
			}
			if(msg!="")
			{
				_me.Common.showAlert(msg,"��ʾ",function()
				{
					objContent.focus();
				});
				return;
			}
			var msg = new _me.Model.Msg(_me.Profile.UserID,uid,$F(objContent).replace(/\n/img,"{br}").escapeHTML(),1,2);
			_webIM.CMD.showChatContent(uid,msg);
			_webIM.Data.sendMessage(msg);
			Elem.Value(objContent);
		},
		//������
		sendFlashMsg:function(uid,o)
		{
			var msg = "";
			if(o.getAttribute("b")=="1")
			{
				msg = "������Ƶ���ط��������񶯡�";
			}
			var users = [];
			users = _me.Friend;
			var userTo = _me.Common.getUserFromArr(uid,users);
			if(userTo&&userTo.OnlineStatus==3)
			{
				msg = "������ϵ�˴����ѻ�״̬����������ܷ��������񶯡�";
			}
			if(userTo&&userTo.IsBlocked)
			{
				msg = "�����������κ��ѷ��������񶯡�";
			}
			if(msg!="")
			{
				_me.Common.showAlert(msg,"��ʾ");
				return;
			}
			o.setAttribute("b","1");
			var msg = new _me.Model.Msg(_me.Profile.UserID,uid,"FLASH",2,2);
			_webIM.CMD.showChatContent(uid,msg);
			_webIM.Data.sendMessage(msg);
			setTimeout(function()
			{
				o.setAttribute("b","0");
			},10000);
		},
		//��ʾ�����б�
		showFaceList:function(id,e)
		{
			Elem.Show("divFaceList");
			var wChat = WinManage.GetWindow(id,1);
			var os = $("divFaceList").style;
			os.top = (wChat.win.Top+wChat.win.Height-223-parseInt($("wChatInputId"+id+"Type1").style.height))+"px";
			os.left = (wChat.win.Left+6)+"px";
			os.width = "342px";
			os.height = "156px";
			document.onmousedown = function(e)
			{
				var ex = Evt.Left(e);
				var ey = Evt.Top(e);
				if(!(ex>parseInt(os.left)&&ex<parseInt(os.left)+parseInt(os.width)+2&&ey>parseInt(os.top)&&ey<parseInt(os.top)+parseInt(os.height)+2))
				{
					Elem.Hid("divFaceList");
					document.onmousedown = null;
				}
			};
		},
		//�򿪺�uid�����촰��
		openChatWindow:function(uid,isfocus)
		{
			var winChat = WinManage.GetWindow(uid,1);
			var wWin = _me.Win;
			if(winChat)//�����Ѵ���
			{
				if(isfocus)
				{
					if(winChat.isMin)winChat.win.Minimize();//������ڴ�����С��״̬��ԭ
					winChat.win.Focus();//����
				}
			}
			else//���ڲ������򴴽�
			{
				var u = _me.Common.getUserFromArr(uid,_me.Friend);
				if(!u)return;
				var w = new WebForm();
				w.Title = u.CustomName;
				w.Icon = "chaticon.gif";
				w.Type = 1;
				w.UserID = uid;
				w.Height = 420;
				w.Width = 420;
				w.MinWidth = 420;
				w.MinHeight = 420;
				var lastChatWin = WinManage.GetLastWindow(1);
				if(!lastChatWin)
				{
					w.Left = wWin.win.Left<425?wWin.win.Left+wWin.win.Width+2:wWin.win.Left-422;
					w.Top = wWin.win.Top;
				}
				else
				{
					w.Left = lastChatWin.win.Left+15;
					w.Top = lastChatWin.win.Top+15;
				}
				w.RepaintMethod = function(w,h,id,t)
					{
						$("wChatMainId"+id+"Type"+t).style.height = (h-6)+"px";
						$("wChatSideId"+id+"Type"+t).style.height = (h-6)+"px";
						$("wChatSideBarId"+id+"Type"+t).style.height = (h-6)+"px";
						_me.Config.ChatSide==1?Elem.Hid($("wChatSideId"+id+"Type"+t)):Elem.Hid($("wChatSideId"+id+"Type"+t));
						$("wChatMainId"+id+"Type"+t).style.width = _me.Config.ChatSide==1?(w-10)+"px":(w-10)+"px";
						$("wChatResizeBarId"+id+"Type"+t).style.width = _me.Config.ChatSide==1?(w-20)+"px":(w-20)+"px";
						$("wChatInputId"+id+"Type"+t).style.width = _me.Config.ChatSide==1?(w-25)+"px":(w-25)+"px";
						$("wChatSignHolderId"+id+"Type"+t).style.width = (w-60)+"px";
						var viewHeight = (h-parseInt($("wChatInputId"+id+"Type"+t).style.height)-71);
						if(viewHeight<10)
						{
							$("wChatInputId"+id+"Type"+t).style.height = (h-parseInt($("wChatViewId"+id+"Type"+t).style.height)-71)+"px";
						}
						else
						{
							$("wChatViewId"+id+"Type"+t).style.height = viewHeight+"px";
						}
					};
				w.LoadedCallback = function(id,t)
					{
						$T("wChatButtonBlockId"+id+"Type"+t,"img")[0].src = u.IsBlocked?"images/chatbuttoncancelblock.gif":"images/chatbuttonblock.gif";
						$T("wChatButtonBlockId"+id+"Type"+t,"img")[0].title = u.IsBlocked?"ȡ����ֹ����ϵ��":"��ֹ����ϵ��";
						$("wChatButtonBlockId"+id+"Type"+t).setAttribute("b",u.IsBlocked?"1":"2");
						$("wChatSignId"+id+"Type"+t).innerHTML = u.UserSign+"";
						//�޸� $("wChatSignId"+id+"Type"+t).innerHTML = u.UserSign+" &lt;"+u.UserEmail+">";
						var sideBtn = $T("wChatSideBarId"+id+"Type"+t,"img")[0];
						sideBtn.src=_me.Config.ChatSide==1?"images/chatsidebutton.gif":"images/chatsidebutton1.gif";
						sideBtn.title = _me.Config.ChatSide==1?"���ز����ߵ���ʾͼƬ":"��ʾ�����ߵ���ʾͼƬ";
						sideBtn.onclick = function()
						{
							_me.Config.ChatSide=_me.Config.ChatSide==1?2:1;
							w.RepaintMethod(w.Width,w.Height-24,w.UserID,w.Type);
							this.src=_me.Config.ChatSide==1?"images/chatsidebutton.gif":"images/chatsidebutton1.gif";
							this.title = _me.Config.ChatSide==1?"���ز����ߵ���ʾͼƬ":"��ʾ�����ߵ���ʾͼƬ";
						};
						$("wChatButtonBlockId"+id+"Type"+t).onclick = function()
						{
							var isblock = 3-parseInt(this.getAttribute("b"));
							_me.CMD.blockFriend(id,isblock);
						};
						$("wChatButtonHistoryId"+id+"Type"+t).onclick = function()
						{
							_me.CMD.showMsgHistory(id);
						};
						$("wChatResizeId"+id+"Type"+t).onmousedown = function()
						{
							var resizeBar = $("wChatResizeBarId"+id+"Type"+t);
							var rs = resizeBar.style;
							rs.top = (90+parseInt($("wChatViewId"+id+"Type"+t).style.height))+"px";
							Elem.Show(resizeBar);
							document.onmousemove = function(e)
							{
								var _top = parseInt(rs.top)+(Evt.Top(e)-(parseInt(rs.top)+w.Top+5));
								_top = _top<140?140:_top;
								_top = _top>w.Height-100?w.Height-100:_top;
								rs.top = _top+"px";
							};
							document.onmouseup = function(e)
							{

								$("wChatViewId"+id+"Type"+t).style.height = (parseInt(rs.top)-90)+"px";
								$("wChatInputId"+id+"Type"+t).style.height = (w.Height-parseInt(rs.top)-75)+"px";
								document.onmousemove = document.onmouseup = null;
								Elem.Hid(resizeBar);
							};
						};
						$("wChatInputId"+id+"Type"+t).onkeydown = function(e)
						{
							var e = e||event;
							if(e.keyCode==13)//�س�
							{
								if(_me.Config.MsgSendKey==1)//enter����
								{
									if(!e.ctrlKey&&!e.shiftKey)//û��ͬʱ������ctrl������Ϣ
									{
										_me.CMD.sendMessage(id);
										return false;
									}
								}
								else//ctrl+enter����
								{
									if(e.ctrlKey)//ͬʱ������ctrl������Ϣ
									{
										_me.CMD.sendMessage(id);
										return false;
									}
								}
							}
							if(_me.Config.MsgSendKey==1&&(!e.ctrlKey && e.keyCode==13))
							{
							}
							else if(_me.Config.MsgSendKey==2&&(e.ctrlKey && e.keyCode==13))
							{
								_me.CMD.sendMessage(id);
								return false;
							}
						};
						$("wChatBtnSendId"+id+"Type"+t).onclick = function()
						{
							_me.CMD.sendMessage(id);
						};
						$("wChatFaceButtonFlashId"+id+"Type"+t).onclick = function()
						{
							_me.CMD.sendFlashMsg(id,this);
						};
						$("wChatFaceButtonFaceId"+id+"Type"+t).onclick = function(e)
						{
							_me.CMD.showFaceList(id,e);
						};
						$("wChatBtnOptionId"+id+"Type"+t).onclick = function(e)
						{
							var m = new SysMenu("menuMsgSendKey"+id+"Type"+t);
							m.Data = ["Ŷ|Elem.Value('wChatInputId"+id+"Type"+t+"','Ŷ');_webIM.CMD.sendMessage("+id+")",
									"�ð�|Elem.Value('wChatInputId"+id+"Type"+t+"','�ð�');_webIM.CMD.sendMessage("+id+")",
									"�ҵ����ˣ��ݰ�!|Elem.Value('wChatInputId"+id+"Type"+t+"','�ҵ����ˣ��ݰ�!');_webIM.CMD.sendMessage("+id+")",
									"",
									"��Enter ��������Ϣ|_webIM.Config.MsgSendKey=1",
									"��Ctrl+Enter ��������Ϣ|_webIM.Config.MsgSendKey=2"];
							m.Data[_me.Config.MsgSendKey+3]+="|images/selected.gif|images/selectedhover.gif";
							m.E = e;
							m.HasIcon = true;
							m.Width = 175;
							m.Show();
						};
								$("wChatFaceFriendId"+id+"Type"+t).src = "images/userface/"+u.UserFace;
								$("wChatFaceMeId"+id+"Type"+t).src = "images/userface/"+_me.Profile.UserFace;
								$("wChatInputId"+id+"Type"+t).focus();
					};
				w.Show();
				w.Focus();
			}
		}
	};
	//��ʼ��
	this.Initialize = function()
	{
		if(!Other.TestCookie())
		{
			alert("ע�⣺�����������֧��Cookie�����ֹ��ܽ��޷�ʹ��!");
		}
		_me.CMD.intLoginWindow();
		_me.Sys = new _me.Model.Sys();
	}
}