/*************************************************
* TQEditor
* Author      : TengQiu Li
* E-Mail      : litqqs@163.com
* Copyright   : TengQiu Li
* Version     : 2.2.4 
* Site        : http://www.e512.net
*************************************************/
(function(window,undefined){
if(window.TQE) return;
var version='2.2.4',//�༭���汾
	ua = navigator.userAgent.toLowerCase(),
	document=window.document,
	isIE =!!window.ActiveXObject, //ie  // ua.indexOf('trident')>=0 || ua.indexOf('msie 6')>=0 ; 
	isIE6 = isIE && !window.XMLHttpRequest,//ie6
	isIE9 = isIE && !!document.createDocumentFragment,
	isWebkit = ua.indexOf('webkit')>=0,//chrome,safari
	isOpera = ua.indexOf('presto')>=0, //opera
	isGecko = !isWebkit && !isOpera && ua.indexOf('gecko')>=0, //firefox
	//�����Դ
	resStrCreateLink='�������',
	resStrTargetOptions='<option value="">Ĭ��</option><option value="_blank">�´���</option><option value="_self">��ǰ����</option><option value="_parent">������</option><option value="_top">���㴰��</option>',
	resStrInsertLinkUI='��ַ:<input type="text" id="url" value="" size=45 /><br />����:<input type="text" id="title" value="" size=45 /><br />��:<select id="target">'+resStrTargetOptions+'</select>',
	resStrInsertImage='����ͼƬ',
	resStrInsertImageUI='ͼƬ��ַ:<input type="text" id="url" value="" size=40 /><br />�滻�ı�:<input name="alt" type="text" id="alt" size="20" /><br />�ߴ�:<input name="width" type="text" id="width" size="4" /> &times; <input name="height" type="text" id="height" size="4" /> px &nbsp;<br />����:<select name="align" id="align"><option value="" >Ĭ��</option><option value="top">����</option><option value="textTop">�ı��Ϸ�</option><option value="middle">����</option><option value="absMiddle">���Ծ���</option><option value="baseline">����</option><option value="bottom">�ײ�</option><option value="absBottom">���Եײ�</option><option value="left">��</option><option value="right">��</option></select> &nbsp; &nbsp; �߿�:<input name="border" type="text" id="border" size="5" />',
	resStrImageContextUI='<form class="ePopForm">��ַ: <input type="text" id="url" size="50"><br>�ߴ�: <input  type="text" id="width" style="width:30px; overflow:visible;"> &times; <input  type="text" id="height" style="width:30px; overflow:visible;">px <a href="javascript:void(0)" id="resetWH" >�����С</a><label><input name="lock" type="checkbox" id="lock" value="1" checked>��������</label><br>�滻�ı�: <input type="text" id="alt" /> �߿�: <input type="text" id="border" style="width:30px; overflow:visible;" /><br>�Ű�: <a href="javascript:void(0)" id="alignDefault">Ĭ��</a> | <a href="javascript:void(0)" id="alignLeft">�󸡶�</a> | <a href="javascript:void(0)" id="alignC">����</a> | <a href="javascript:void(0)" id="alignRight">�Ҹ���</a> | <a href="javascript:void(0)" id="alignTop">�϶���</a> | <a href="javascript:void(0)" id="alignAbsMiddle">�ж���</a> | <a href="javascript:void(0)" id="alignBottom">�¶���</a><div><hr>����: <input type="text" id="link" size="40"><select id="target">'+resStrTargetOptions+'</select></div></form>',
	resStrInsertFlash='���붯��',
	resStrInsertFlashUI='Flash��ַ:<input type="text" id="url" value="" size=40 /><br />�ߴ�:<input name="width" type="text" id="width" size="4" value="200" />&times;<input name="height" type="text" id="height" size="4" value="200" /> px',
	resStrInsertMusic='������Ƶ',
	resStrInsertMusicUI='��Ƶ��ַ:<input type="text" id="url" value="" size=40 /><br />�Զ�����:<select id="auto_start" name="auto_start"><option value="0">�ֶ���ʼ</option><option value="1">�Զ���ʼ</option></select>',
	resStrInsertVideo='������Ƶ',
	resStrInsertVideoUI='��Ƶ��ַ:<input type="text" id="url" value="" size=40 /><br />�ߴ�:<input name="width" type="text" id="width" size="4"  value="320"/>&times;<input name="height" type="text" id="height" size="4" value="240" /> px<br />�Զ�����:<select id="auto_start" name="auto_start"><option value="0">�ֶ���ʼ</option><option value="1">�Զ���ʼ</option></select><br />��Ƶ����:<select id="video_type" name="video_type"><option value="auto">�Զ�ʶ��</option><option value="flv">FLV</option><option value="rm">RMVB</option><option value="wm">������Ƶ</option></select>',
	resStrInsertRow='������',
	resStrInsertCol='������',
	resStrDeleteRow='ɾ����',
	resStrDeleteCol='ɾ����',
	resStrDeleteTable='ɾ�����',
	resStrTableContextUI='<div><a class="eMenuItem" style="width:168px;" ><span id="eMergeCells" class="eButton"></span>�ϲ���Ԫ��</a><a style="width:168px;" class="eMenuItem" ><span id="eSplitCells" class="eButton"></span>��ֵ�Ԫ��</a><hr width="166"><div style="padding:1px 1px 1px 3px;">����Ű�: <a id="eTableAlignLeft">����</a> | <a id="eTableAlignCenter">����</a> | <a id="eTableAlignRight">����</a></div><table border="0" cellSpacing="2" cellPadding="0"><tbody><tr><td width="65" align="center" valign="top" bgcolor="#FFFFFF" style="border:#999 solid 1px;">���ݶ���<br><a id="eCellAlignLT" class="eButton eCellAlign"></a><a id="eCellAlignCT" class="eButton eCellAlign"></a><a id="eCellAlignRT" class="eButton eCellAlign"></a><br><a id="eCellAlignLC" class="eButton eCellAlign"></a><a id="eCellAlignCC" class="eButton eCellAlign"></a><a id="eCellAlignRC" class="eButton eCellAlign"></a><br><a id="eCellAlignLB" class="eButton eCellAlign"></a><a id="eCellAlignCB" class="eButton eCellAlign"></a><a id="eCellAlignRB" class="eButton eCellAlign"></a></td><td width="90" align="center" valign="top">���<input id="width" name="width" size="3" type="text"><br>�߿�<input id="border" name="border" size="3" type="text"><br>���<input id="cellSpacing" name="cellSpacing" size="3" type="text"><br>���<input id="cellPadding" name="cellPadding" size="3" type="text"></td></tr></tbody></table></div><div><hr width="166">������</div>',
	resStrNoColor='��ɫ',
	resStrOK='ȷ��',
	resStrCenal='ȡ��',
	resStrUpload='�ϴ�',
	resStrClose='�ر�',
	resStrVersion='�汾',
	resStrAuthor='����',
	resStrSite='����',
	resStrMail='����',

TQE=function(objId,userConfig)
{
	var $=this,config=TQE.clone(TQE.config);
	$.objId = objId;
	if(userConfig){
		for(var k in userConfig){
			config[k] = userConfig[k];
		}
	}

	if(typeof config.toolbar =='string'){
		var tm=TQE.toolbarMode[config.toolbar];
		if(!tm) tm=TQE.toolbarMode.normal;
		config.toolbar=tm.left;
		if(!userConfig || undefined===userConfig.toolbarRight)config.toolbarRight=tm.right;
	}

	$.config = config;
	$.hasFocus=false;
	$._containerDiv=null;
	$._toolbarDiv=null;
	$._clientDiv=null;
	$._statusDiv=null;
	//Methods
	$.init=function(){
		$.obj = TQE.find($.objId);
		if(!$.obj){
			if($.config.debug) throw('ȱ��IDΪ '+$.objId+' ��Ԫ��');
			return;
		}
		if($.obj.editor) return;//���󲻴��ڻ���ִ�й�
		$.obj.editor = $;

		if (!$.config.width || $.config.width == "auto") {
			if($.obj.style.width){$.config.width = $.obj.style.width; }
			else if($.obj.cols){ $.config.width = ($.obj.cols * 8)+'px'; }
			else{ $.config.width = '100%'; }
		}
		if (!$.config.height || config.height == "auto") {
			if($.obj.style.height){ $.config.height = $.obj.style.height; }
			else if($.obj.rows){ $.config.height = ($.obj.rows * 17)+'px' }
			else{ $.config.height = '200'; }
		}

		//toolbar
		var i,j, btn,code ='<div class="eToolBar" id="'+$.objId+'_EditorToolBar">';
		//
		if($.config.toolbarRight.length){
			code +='<ul class="eToolBarGroupRight" >';
			for(i=0; i < $.config.toolbarRight.length; i++){
				code += $._btnCode($.config.toolbarRight[i], false);
			}
			code +='</ul>';
		}
		//������
		code +='<ul class="eToolBarGroup" >';
		if($.config.showToolbarGroupHandle){
			btn = TQE.toolbarBottons['group'];
			code +='<div class="eButton" id="eBTNgroup"  style="background-position: -'+btn.left+'px 0;width:'+btn.width+'px;" ></div>';
		}
		for(i=0; i < $.config.toolbar.length; i++){
			code += $._btnCode($.config.toolbar[i], true);
		}
		code += '</ul>'+	   //��������
			'<div class="eClear"></div></div>'+
			//�༭��
			'<div id="'+$.objId+'_EditorClient" class="eEditorClient">'+
			'<iframe id="'+$.objId+'_Editor" style="height:'+$.config.height+';" frameborder=0 ></iframe>'+
			'</div>'+ 
			'</div>'+
			'<div class="eStatusBar"><div class="eResizeWH" ></div><div class="eResizeH" ></div></div>';

		if (!$.config.debug) {
			$.obj.style.display = "none";
		}
		var e =TQE.CE('DIV','TQEditorContainer_'+$.objId, code );
		e.className='eEditor';
		e.style.width=$.config.width;
		$.obj.parentNode.insertBefore(e, $.obj);

		$._containerDiv = e;
		$._toolbarDiv = TQE.find($.objId+'_EditorToolBar');
		$._clientDiv = TQE.find($.objId+'_EditorClient');

		TQE._noSel($._toolbarDiv);
		//TQE._noSel($._toolbarDiv,'DIV,UL,LI');

		var btns = TQE._tags($._toolbarDiv,'LI');
		for(i=btns.length-1;i>=0;i--){
			$._ieMouseStyle(btns[i],'eButton','eButtonUp','eButtonDown');	//ie6

			TQE.addListener(btns[i],'click',	$._toolbarBottonClick);
			TQE.addListener(btns[i],'mouseout',	function(){
					var o=TQE.activePop;
					this.className=	(o && o.relateObj==this)? 'eButtonDropDown':'eButton';
					});
		}
		if(isIE){//IE6��IE8����ģʽʱ��Ҫ
			var groups = TQE._tags($._toolbarDiv,'UL'),w;
			for(i=groups.length-1;i>=0;i--){
				w=3;
				btns=groups[i].childNodes;
				for(j=btns.length-1;j>=0;j--){w += btns[j].offsetWidth+2;}
				groups[i].style.minWidth=w+'px';
			}
		}
		//--
		if('function'===typeof $.config.oninit) $.config.oninit.call($);
		$._initContent($.obj.value);
		$.obj.focus = function(){this.editor.focus();}
		//ie
		//$._containerDiv.onresize=$._fixWidth;
		//if(!isIE && $.config.width.indexOf('%')){
		//	TQE.addListener(window,'resize',$._fixWidth);
		//}
		//--
		$._statusDiv = e.lastChild;
		if(false==$.config.resize)$._statusDiv.style.display='none';
		$._statusDiv.onmousedown=function(a){
			var x,y,obj,maskDiv,
				o=this,
				
				//$=this.context,
				p=TQE.pos(this), 
				scrollTop = TQE._docScrollTop(),
				scrollLeft = TQE._docScrollLeft(),
				h =parseInt( $.config.height,10),
				w=$._containerDiv.offsetWidth-2,
				bChangeWidth = this.style.cursor == 'se-resize',
				oldTop = TQE.pos($._clientDiv).y,
				documentMM=document.onmousemove,
				documentMU=document.onmouseup;
			
			$._showTableContext(0);

			if(isIE){
				a=window.event;
				obj=a.srcElement;
				x=scrollLeft + a.clientX-p.x;
				y=scrollTop + a.clientY-p.y;
			}else{
				obj=a.target;
				x=a.pageX-p.x;
				y=a.pageY-p.y;

			}
			bChangeWidth = bChangeWidth || obj===o.firstChild;
			
			o.mouseout=o.onmouseleave=null;
			//$.setOpacity(o,80);
			if(isWebkit || isIE6){
				maskDiv=TQE.CE('DIV','','',true);
				maskDiv.className='eDragMaskDiv';
				maskDiv.style.cursor= bChangeWidth ? 'se-resize':'s-resize';
				maskDiv.style.height=Math.max(document.documentElement.clientHeight, document.body.scrollHeight)+'px';
			}
			if(o.setCapture)
				o.setCapture();
			else if(window.captureEvents)
				window.captureEvents(Event.MOUSEMOVE|Event.MOUSEUP);

			document.onmousemove=function(a){
				if(isIE){
					a=window.event;
					a.pageX=scrollLeft + a.clientX;
					a.pageY=scrollTop + a.clientY;
				}
				var deltaH=oldTop - TQE.pos($._clientDiv).y, ty=Math.max(50-h+p.y-deltaH, a.pageY-y),tx=Math.max(100-w+p.x, a.pageX-x);
				if(bChangeWidth){
					//o.style.width= tx-p.x + w +'px';
					$.setWidth(tx-p.x + w);
				}
				//o.style.top=ty +'px';
				$.setHeight(ty-p.y + h + deltaH );
			};

			document.onmouseup=function(a){
				if(o.releaseCapture)
					o.releaseCapture();
				else if(window.releaseEvents)
					window.releaseEvents(Event.MOUSEMOVE|Event.MOUSEUP);
				document.onmousemove=documentMM;
				document.onmouseup=documentMU;

				if(maskDiv) document.body.removeChild(maskDiv);
				maskDiv=null;
				$.focus();
				$._popTableContext();
				//$.setOpacity(o,100);
			};
		};
	};
	$._btnCode=function(tn, groupFlag){
		var submenu='',btn,code,styleCode;
		if('||'===tn) tn='group';
		else if('|'===tn) tn='separator';
		else if('string'!==typeof tn){//�Ӳ˵�
			submenu= 'submenu="'+tn.join(',')+'" ';
			tn='submenu';
		}
		btn = TQE.toolbarBottons[tn];
		if(!btn) return '';
		styleCode=function(btn){return 'id="eBTN'+tn+'" class="eButton"' + (btn.css ? ' style="'+btn.css+'"':'' ); };
		switch(tn){
			case 'group':
				if(!groupFlag) return '';
				code='</ul><ul class="eToolBarGroup" >';
				if($.config.showToolbarGroupHandle){
					code+= '<div '+styleCode(btn)+' ></div>';
				}
				return code;
			case 'separator':
				return '<div '+styleCode(btn)+' ></div>'; // id="eBTN'+tn+'" 
			default:
				return'<li name="'+tn+'" title="'+btn.text+'" '+submenu+styleCode(btn)+' ></li>';//id="'+$.objId+'_'+tn+'"
		}
	};
	//ie6
	$._ieMouseStyle=function(obj,normal,hover,down){
		if(!isIE || !isIE6 && document.compatMode == "CSS1Compat") return;
		TQE.addListener(obj,'mouseover',	function(){this.className=hover;});
		TQE.addListener(obj,'mouseup',	function(){this.className=hover;});
		TQE.addListener(obj,'mousedown',	function(){this.className=down;});
		TQE.addListener(obj,'mouseout',	function(){this.className=normal;});
	};
	$._keepBookmark=function(){
		if(isIE){
			var _bookmark,
				win = TQE.find($.objId+'_Editor');
			TQE.addListener(win,'beforedeactivate',function(){_bookmark=$._getRange();});
			TQE.addListener(win,'activate',function(){
				if(null==_bookmark) return;
				var sel=$._getSelection();
				if(sel.addRange) sel.addRange(_bookmark);
				else _bookmark.select();
				_bookmark=null;
			});
		}
	};
	$._toolbarBottonClick=function(){
		var e=isIE ? event : arguments[0],
			cmd = this.getAttribute('name'),
			btn=TQE.toolbarBottons[cmd];
		e.cancelBubble = true;
		$.focus();
		if(btn && btn.click){
			btn.click($,this);
		}else{
			TQE.hidePop();
			$.exeCmd(cmd);
		}
	};
	$.__selectOption=function(selectObj,value){
		for(var i=selectObj.options.length-1;i>=0;i--){
			if(selectObj.options[i].value==value){
				selectObj.selectedIndex=i;
				return;
			}
		}
	};
	$.exeCmd=function(cmd,para){
		$.focus();
		var cells=$.selectedCells(),
			d=$._getDoc(),
			sel=$._getSelection(),
			i, r, fcs;
		if(cells.length){
			r= d.createRange ? d.createRange() : d.body.createTextRange();
			for(i=0;i<cells.length;i++){
				if('backcolor'==cmd || 'hilitecolor'==cmd){
					cells[i].style.backgroundColor=para;
					continue;
				}
				if(sel.empty) sel.empty();	//ie6,7,8
				else if(sel.removeAllRanges) sel.removeAllRanges();
				
				if(sel.selectAllChildren){
					sel.selectAllChildren(cells[i])
				}else if(r.moveToElementText){ //ie6,7,8
					r.moveToElementText(cells[i]);
					r.select();
				//}else{
				//	r.selectNode(cells[i]);
				//	sel.addRange(r);
				}
				try{
					d.execCommand(cmd,false, para!=null?para:false);
					if('bold'==cmd || 'italic'==cmd || 'underline'==cmd || 
						'strikethrough'==cmd || 'subscript'==cmd || 'superscript'==cmd){
						if(i==0) fcs = d.queryCommandState(cmd);
						else if(fcs != d.queryCommandState(cmd)) d.execCommand(cmd,false, para!=null?para:false);
					}
				}catch(e){}
			}
			if(sel.empty) sel.empty();	//ie6,7,8
			else if(sel.removeAllRanges) sel.removeAllRanges();
		}else{
			try{return d.execCommand(cmd,false, para!=null?para:false);}catch(e){}
		}
	};
	$.focus=function(){
		if($.hasFocus) return ;
		if('visual'==$.currentMode()){
			var cells=$.selectedCells(),r,sel;
			if(cells.length<1){
				$._editorObj().contentWindow.focus();
				var r=$._getRange(), sel = $._getSelection();
				if(isWebkit && sel.rangeCount<1){
					r.selectNodeContents($._getDoc().body);
					r.collapse(true);
					sel.removeAllRanges();
					sel.addRange(r);
				}
			}else{
				var r=$._getRange();
				if(r.select){ //ie6,7,8
					r.collapse(false);
					r.select();
				}else{
					r.selectNodeContents(cells[0]);
					r.collapse(false);
					sel = $._getSelection();
					sel.removeAllRanges();
					sel.addRange(r);
				}
			}
		}else{
			$._editorObj().focus();
		}
		$.hasFocus=true;

	};
	$.focusNode=function(){
		if('visual'!==$.currentMode() || !$.hasFocus ) return null;
		var r,node = $._getSelection().focusNode;
		if(!node){//ie6,7,8 
			r = $._getRange();
			return (r.length)? r.item(0).parentNode : r.parentElement? r.parentElement():null;
		}
		return node;
	};
	$._showColorSelectUI=function(sender,cmd){
		var o=TQE.find($.objId+'ColorSelect'+cmd);
		if(null==o){
			o = TQE.CE('DIV',$.objId+'ColorSelect'+cmd, TQE._colorPickerUI() ,true);
			o.className='ePopPanel';
			TQE._noSel(o);//,'TABLE,TR,TD,DIV');
			TQE.listenTags(o,'TD','click',function(){TQE.hidePop();$.exeCmd(cmd, this.bgColor);});
			o.lastChild.lastChild.onclick=function(){TQE.hidePop();$.exeCmd(cmd, '');};

			TQE.listenTags(o,'TD','mouseover',function(){
				var self=this,
					iid = window.setInterval(function(){self.style.backgroundColor=(self.style.backgroundColor!='')?'':'#FFFFFF'},500),
					p=this.parentNode.parentNode.parentNode.parentNode.lastChild;
				p.firstChild.style.backgroundColor=this.bgColor;
				p.firstChild.nextSibling.innerHTML=this.bgColor;

				this.style.backgroundColor='#FFFFFF';
				this.onmouseout=function(){window.clearInterval(iid);this.style.backgroundColor='';}
			});
		}
		TQE.pop(o,sender);
	};
	$._popAsMenu=function(sender, btns){
		var cmd=sender.getAttribute('name'), 
			o=TQE.find($.objId+'PopMenu'+cmd), 
			s='',
			es,i,tn,btn;
		if(null==o){
			o = TQE.CE('DIV',$.objId+'PopMenu'+cmd, '' ,true);
			o.className='ePopPanel';
		}
		for(i=0; i < btns.length; i++){
			tn = btns[i];
			if('separator'==tn || '|'==tn){
				s +='<hr style="width:120px; clear:both;" />';
				continue;
			}
			btn = TQE.toolbarBottons[tn];
			if(btn){
				s +='<div  name="'+tn+'" class="eMenuItem" ><span class="eButton" id="eBTN'+tn+'" ></span>'+btn.text+'</div>';
			}
		}
		o.innerHTML=s;
		TQE._noSel(o);
		if(isOpera)TQE._noSel(o,'DIV');
		TQE.listenTags(o,'DIV','click',$._toolbarBottonClick);
		if(isIE6){ 
			es = TQE._tags(o,'DIV');
			for(i=es.length-1;i>=0;i--){
				$._ieMouseStyle(es[i],'eMenuItem','eMenuItemUp','eMenuItem');
			}
		}
		TQE.pop(o,sender);
	};
	$._popFormUI=function(id, caption, htm, cbOK, intInputNames, uploadUrl, uploadFileTypes, uploadFileSize, uploadType){
		var s='<form class="ePopForm">';
		if(uploadUrl){
			s+='<div class="eTabBar"><a href="javascript:void(0)" class="active" HIDEFOCUS="true">'+caption+'</a><a href="javascript:void(0)" HIDEFOCUS="true">'+resStrUpload+'</a></div>';
			s+='<div>';
		}
		s += htm;
		s+='<div class="eBar"><button id="OK" name="OK" type="submit" class="eBtn"  HIDEFOCUS="true"/>'+resStrOK+'</button> &nbsp; <button id="Cancel" name="Cancel" type="button" class="eBtn"  HIDEFOCUS="true">'+resStrCenal+'</button></div>';
		if(uploadUrl){
			s+='</div><div style="display:none">';
			var url=uploadUrl;
			if(uploadUrl.substr(0,1)=='#'){//�Զ��ϴ�����
				url=uploadUrl.substring(1);
			}else{
				//���������ַ
				if(!('http://'==uploadUrl.substr(0,7).toLowerCase() || '/'==uploadUrl.substr(0,1) || 'https://'==uploadUrl.substr(0,6).toLowerCase())){
					if('?'==uploadUrl.substr(0,1)){
						uploadUrl = location.pathname+uploadUrl;
					}else{
						var a=location.pathname.split('/');
						a[a.length-1]= uploadUrl;
						uploadUrl = a.join('/');
					}
				}
				url=TQE.url+'upload.htm?url='+encodeURIComponent(uploadUrl)+'&field_name='+$.config.uploadFieldName+'&exts='+encodeURIComponent(uploadFileTypes)+'&max_size='+encodeURIComponent(uploadFileSize)+'&debug='+$.config.debug+'&upload_type='+uploadType;
			}
			s+='<iframe src="'+url+'" frameborder=0 width=358 height=45 ALLOWTRANSPARENCY=true marginwidth=0 marginheight=0 ></iframe>';
			s+='</div>';
		}
		
		s +='</form>';

		var o= TQE.CE('DIV',id, s ,true);
		o.className='ePopPanel';

		if(uploadUrl){
			var tabClick=function(){
				this.className='active';
				var pn= this.parentNode.nextSibling;
				if(this.nextSibling){
					this.nextSibling.className='';
					pn.style.display='';
					pn.nextSibling.style.display='none';
					pn.parentNode.getElementsByTagName('INPUT')[0].focus();
				}else if(this.previousSibling){
					this.previousSibling.className='';
					pn.style.display='none';
					pn.nextSibling.style.display='';
				}
				},
				tabbar = o.firstChild.firstChild;
			tabbar.firstChild.onclick=tabClick;
			tabbar.lastChild.onclick=tabClick;
		}
		var f= TQE._tags(o,'FORM')[0],es = f.elements;
		f.srcEditor=$;
		es['Cancel'].onclick= TQE.hidePop;
		es['OK'].onclick=cbOK;
		f.onsubmit=function(){
			this.elements['OK'].click();
			return false;
		}
		if(intInputNames) for(i=intInputNames.length-1;i>=0;i--) $._intEdit(es[intInputNames[i]]);
		TQE._noSel(o);
		if(isOpera)TQE._noSel(o,'FORM,DIV');
		return o;
	};
	$.insertMusic=function(url,auto_start){
		var s = $.config.tplMusic.replace(/\{\$url\}/g,url).replace(/\{\$auto_start\}/g,auto_start);
		$.insertHtml(s);
	};
	$.insertFlash=function(url,width,height){
		var s = $.config.tplFlash.replace(/\{\$url\}/g,url);
		s = s.replace(/\{\$width\}/g,width);
		s = s.replace(/\{\$height\}/g,height);
		$.insertHtml(s);
	};
	$.insertVideo=function(url,auto_start,width,height,video_type){
		var tpl='',ext;
		switch(video_type){
			case 'rm':tpl = $.config.tplRm;break;
			case 'flv':tpl = $.config.tplFlv;break;
			case 'wm':tpl = $.config.tplVideo;break;
			default:
				ext = url.substr(url.lastIndexOf('.')).toLowerCase();
				if('.flv'==ext){
					tpl = $.config.tplFlv;
				}else if('.rm'==ext || '.rmvb'==ext){
					tpl = $.config.tplRm;
				}else{
					tpl = $.config.tplVideo;
				}
			break;
		}
		tpl = tpl.replace(/\{\$url\}/g,url).replace(/\{\$auto_start\}/g,auto_start).replace(/\{\$width\}/g,width).replace(/\{\$height\}/g,height);
		$.insertHtml(tpl);
	};
	//����ģʽʱ��Ч
	//position ��ѡֵ[before|after|replace|-1|1|0]
	$.insertHtml=function(htm, position){//
		$.focus();
		if('code'==$.currentMode()){
			var editor=$._editorObj();
			if(document.selection){//ie
				var r = document.selection.createRange();
				if('before'==position || -1==position) r.text = htm+r.text;
				else if('after'==position || 1==position) r.text+=htm;
				else r.text=htm;
				r.select();
			}else if(editor.selectionStart || editor.selectionStart=='0'){
				var startPos = editor.selectionStart,
					endPos = editor.selectionEnd,
					scrollTop = editor.scrollTop,
					str=editor.value;
				if('before'==position || -1==position) htm = str.substring(startPos,endPos)+htm;
				else if('after'==position || 1==position)  htm += str.substring(startPos,endPos);

				editor.value=str.substr(0,startPos)+htm+str.substr(endPos, str.length);
				editor.focus();
				editor.selectionStart=editor.selectionEnd= startPos+htm.length;
				editor.scrollTop=scrollTop;
			}else{
				editor.value+=htm;
				editor.focus();
			}
			$.save();
			return;
		}
		//$.unselectAllCells();
		if(isGecko){
			htm = htm.replace(/<td([^>]*)>\s*<\/td/ig, '<td$1 ><br></td');
		}
		var r = $._getRange();
		if(r.length){//ie
			var d=$._getDoc(),
				e=r.item(0), 
				tr=d.body.createTextRange();
			tr.moveToElementText(e);
			if('before'==position || -1==position){
				tr.collapse();
			}else if('after'==position || 1==position){
				tr.collapse(false);
			}else{
				r.execCommand('Delete');
			}
			tr.pasteHTML(htm);
			tr.select();//������
		}else if(r.insertNode){//ie9,firefox,chrome,opera
			var cf,n,sel;
			if(r.createContextualFragment){//IE9 δʵ�ִ˷���
				cf=r.createContextualFragment(htm);
			}else{
				cf =$._getDoc().createDocumentFragment();
				cf.appendChild(TQE.CE('div'));
				cf.firstChild.outerHTML=htm;
			}

			if('before'==position || -1==position){
				n= cf.lastChild;
				r.insertNode(cf);
				r.setEndAfter(n);
			}else if('after'==position || 1==position){
				r.collapse(false);
				r.insertNode(cf);
			}else{
				r.deleteContents();
				r.insertNode(cf);
			}
			r.collapse(false);
			sel	= $._getSelection();
			sel.removeAllRanges();
			sel.addRange(r);
		}else{ //ie6,7,8
			if('before'==position || -1==position){
				r.collapse(true);
			}else if('after'==position || 1==position){
				r.collapse(false);
			}
			r.pasteHTML(htm);
			r.collapse(false);
			r.select();
		}
		$.save();
	};
	//���ĵط�������ҳ����
	$.appendHtml=function(html){
		$.unselectAllCells();
		$.setContent($.content()+html);
		//if('code'==$.currentMode()) $._editorObj().value += html;
		//else $._getDoc().body.innerHTML += html;
		//return $;
	};
	$._editorObj= function(){
		return  TQE.find($.objId+'_Editor');
	};
	$.currentMode=function(){//ȡ�õ�ǰ�ı༭ģʽ, code ���� , visual ���ӻ�
		return ($._editorObj().tagName.toUpperCase()==='IFRAME')?'visual':'code';
	};
	$._getDoc=function(){
		var iframe =$._editorObj();
		return iframe.contentDocument || iframe.contentWindow.document;
	};
	$._getSelection=function(){
		var win =$._editorObj().contentWindow;
		return (win.getSelection) ? win.getSelection() : $._getDoc().selection;
	};
	$._getRange=function(){
		var selection=$._getSelection(), d=$._getDoc(),r;
		try{//��׼dom
			r = selection.rangeCount > 0 ? selection.getRangeAt(0) : (selection.createRange ? selection.createRange() : d.createRange());
		}catch (e){}
		if(!r) r = isIE ? d.body.createTextRange() : d.createRange();
		return r;
	};
	$.insertImage=function(image_url,width,height,border,alt,align){
		if(''==image_url) return;
		var s='<img src="'+image_url+'" ';
		if(width!==undefined && ''!=width && !isNaN(parseInt(width)) && parseInt(width)>0 ) s += ' width='+parseInt(width);
		if(height!==undefined && ''!=height && !isNaN(parseInt(height)) && parseInt(height)>0 ) s +=' height='+parseInt(height);
		if(border!==undefined && ''!=border && !isNaN(parseInt(border))) s +=' border='+parseInt(border);
		if(alt!==undefined && ''!=alt) s +=' alt="'+alt+'"';
		if(align!==undefined && ''!=align) s +=' align="'+align+'"';
		s +=' />';
		$.insertHtml(s); 
		return;
	};
	$._enableButton=function(btn,enable){
		if(enable){
			btn.disabled=false;
			$.setOpacity(btn,100);
			TQE.addListener(btn,'click',$._toolbarBottonClick);
		}else{
			btn.disabled=true;
			$.setOpacity(btn,30);
			TQE.delListener(btn,'click',$._toolbarBottonClick);
		}
	};
	$._updateToolbar=function(enable, buttons, exceptFlag){
		var tb=$._toolbarDiv,
			btns = TQE._tags(tb,'LI'),
			i;
		for(i=btns.length-1;i>=0;i--){
			if($.inArray(btns[i].getAttribute('name'), buttons) ^ exceptFlag  ) {
				$._enableButton(btns[i],enable);
			}
		}
	};
	//���ù������
	//$.setCaret=function(node, position){
	//	var d=$._getDoc(), sel = $._getSelection();
	//	if(isIE){
	//		tr=d.body.createTextRange();
	//		tr.moveToElementText(node)
	//		tr.collapse(true);
	//		tr.select();
	//	}else{	
	//		tr = d.createRange();
	//		tr.selectNode(node);
	//		tr.collapse(true);
	//		sel.removeAllRanges();
	//		sel.addRange(tr);
	//	}
	//};
	$._hook=function(){
		//if('visual'!=$.currentMode()) return; //!isIE || !$.config.enter2br || 
		var doc=$._getDoc();
		doc.onkeydown = function(e){
			TQE._hideContext();
			if(isIE) e = $._realEvent(e);
			var key = e.keyCode || e.which;
			if(e.ctrlKey &&(13==key || 10==key || 83==key)) $.save();//�� Ctrl+(Enter|s) ������������
			var r=$._call('keydown',e);
			if(undefined!==r && false==r) return false;//�û��Ѵ���

			if((8==key || 46==key) && $.selectedCells().length)  $.exeCmd('delete'); //���ѡ�е�Ԫ������

			if(key==13){
				var o = $.focusNode(), p,br,tr,d;
				if(!o) return;
				if(o.nodeName==='#text') o = o.parentNode;
				if(o.className == 'splitPage'){//��ҳ����
					p=o.parentNode;
					if(o.nextSibling && o.nextSibling.nodeName==='BR'){
						br = o.nextSibling;
					}else{
						br = document.createElement('BR');
						if(p.lastChild == o){
							p.appendChild(br);
						}else{
							p.insertBefore(br,o.nextSibling);
						}
					}
					//������
					d=$._getDoc();
					if(isIE){
						tr=d.body.createTextRange();
						tr.moveToElementText(br)
						tr.collapse(true);
						tr.select();
					}else{	
						tr = d.createRange();
						tr.selectNode(br);
						tr.collapse(true);
						var sel	= $._getSelection();
						sel.removeAllRanges();
						sel.addRange(tr);
					}
					return false;
				}
				//--
				//if(e.shiftKey){e.shiftKey=false;return true;}
				if($.config.enter2br){
					if(o.tagName != 'LI'){
						$.insertHtml('<br />');
						return false;
					}
				}else if(!isIE && !e.shiftKey){//��IE, �س�Ĭ����br��,��������
					var r=$._getRange(), n, nn;
					for(n=r.startContainer; n; n=n.parentNode){
						nn=n.nodeName;
						if(nn=='LI' || nn=='P'|| 'H1'==nn|| 'H2'==nn|| 'H3'==nn|| 'H4'==nn|| 'H5'==nn|| 'H6'==nn|| 'DIV'==nn) return true;
					}
					$.exeCmd('FormatBlock','p');
				}
			}
		};
		doc.onkeyup =function(e){
			e = $._realEvent(e);
			var key = e.keyCode || e.which, 
				node = this.body.firstChild;
			if((8===key || 46===key) && node&& node === doc.body.lastChild ){//delete | backspace
				var nn=node.nodeName,
					ns=node.innerHTML,
					reg=/^<(strong|em|b|i|u|font)[^>]*><\/\1>|&nbsp;$/i;
				if(nn==='P' || nn==='H1'|| nn==='H2'|| nn==='H3'|| nn==='H4'|| nn==='H5'|| nn==='H6' || nn==='BR'){
					if(''===ns || reg.test(ns))doc.body.innerHTML='';
				}
			}
			$._call('keyup',e);
		};

		doc.ondragstart=function(e){TQE._hideContext();$._showTableContext(0);};
		doc.onclick=function(e){
			var node,r,sel;
			TQE._hideContext();
			//$.focus();
			//$._popTableContext();
			if(!isIE){
				node=e.target;
				if('IMG'===node.nodeName){
					r = $._getDoc().createRange();
					sel = $._getSelection();
					r.selectNode(e.target);
					sel.removeAllRanges();
					sel.addRange(r);
					$._popImagePanel(node);
				}
			}else{
				e=$._realEvent(e);
				node= e.srcElement;
				if('IMG'===node.nodeName){
					$._popImagePanel(node);
				}
			}
		};
		if(isIE){
			doc.onmousedown=doc.onclick;
			doc.onclick=null;
		}
		if($.config.autoClean) doc.body.onpaste=function(){setTimeout($.cleanWord,200);};

		TQE.addListener(doc,'mouseup',TQE.hidePop);
		TQE.addListener(doc,'scroll',$._popTableContext);
		//if(isIE6 || isIE && (document.documentMode==7 || document.compatMode != "CSS1Compat"))
		TQE.addListener(doc.body,'scroll',$._popTableContext); //IE6�������ͼ
		
		TQE.addListener(doc,'keyup', $._popTableContext);
		TQE.addListener(doc,'mouseup',$._popTableContext);
		TQE.addListener(doc,'mousedown',$._startSelectTd);
		TQE.addListener(doc,'keypress',$.unselectAllCells);
		//TQE.addListener(doc,'blur',$.unselectAllCells);

		if(isIE6){
			TQE.addListener(doc,'mouseup',$.focus);
			TQE.addListener(doc,'mousedown',$.focus);
		}
	};
	$._popImagePanel=function(img){
		var o=TQE.find('__tqImagePanel');
		if(null==o){
			//var s='<form class="ePopForm">��ַ: <input type="text" id="url" size="50"><br>�ߴ�: <input  type="text" id="width" style="width:30px; overflow:visible;"> &times; <input  type="text" id="height" style="width:30px; overflow:visible;">px <a href="javascript:void(0)" id="resetWH" >�����С</a><label><input name="lock" type="checkbox" id="lock" value="1" checked>��������</label><br>�滻�ı�: <input type="text" id="alt" /> �߿�: <input type="text" id="border" style="width:30px; overflow:visible;" /><br>�Ű�: <a href="javascript:void(0)" id="alignDefault">Ĭ��</a> | <a href="javascript:void(0)" id="alignLeft">�󸡶�</a> | <a href="javascript:void(0)" id="alignC">����</a> | <a href="javascript:void(0)" id="alignRight">�Ҹ���</a> | <a href="javascript:void(0)" id="alignTop">�϶���</a> | <a href="javascript:void(0)" id="alignAbsMiddle">�ж���</a> | <a href="javascript:void(0)" id="alignBottom">�¶���</a><div>����: <input type="text" id="link" size="40"><select id="target"><option value="" >Ĭ��</option><option value="_blank">�´���</option><option value="_self">��ǰ����</option><option value="_parent">������</option><option value="_top">���㴰��</option></select></div></form>';
			o= TQE.CE('DIV','__tqImagePanel', resStrImageContextUI ,true);
			o.className='ePopPanel';
			var f= TQE._tags(o,'FORM')[0],
				es = f.elements;
			$._intEdit(es['width']);
			$._intEdit(es['height']);
			$._intEdit(es['border']);

			o.urlField=es['url'];
			o.widthField=es['width'];
			o.heightField=es['height'];
			o.lockField=es['lock'];
			o.borderField=es['border'];
			o.altField=es['alt'];
			//link attri
			o.linkField=es['link'];
			o.targetField=es['target'];
			o.linkField.onchange=function(){
				var o=this.parentNode.parentNode.parentNode, img=o.target, a=img.parentNode;
				if(a.nodeName!=='A') return;
				a.setAttribute('href', this.value);
			};
			o.targetField.onchange=function(){
				var o=this.parentNode.parentNode.parentNode, img=o.target, a=img.parentNode;
				if(a.nodeName!=='A') return ;
				a.setAttribute('target', this.options[this.selectedIndex].value);
			};

			o.urlField.onchange=function(){
				this.parentNode.parentNode.target.src=this.value;;
			};
			o.urlField.onfocus=function(){this.select();};
			var setwh=function(){
				var o=this.parentNode.parentNode, img=o.target, w=h=Math.max(1,this.value);
				if(o.lockField.checked){
					if(this.id=='width') h=o.heightField.value=Math.round(w/ o.rate);
					else w=o.widthField.value=Math.round(h * o.rate);
				}else{
					if(this.id=='width') h=o.heightField.value;
					else w=o.widthField.value;
				}
				img.width=w;
				img.height=h;
				img.style.width='';
				img.style.height='';
			};
			o.widthField.onchange=o.widthField.onblur=o.heightField.onchange=o.heightField.onblur=setwh;
			o.borderField.onchange=function(){
				var img= this.parentNode.parentNode.target;
				if(''===this.value) img.removeAttribute('border'); else img.border=this.value;
			};
			o.altField.onchange=function(){
				var img= this.parentNode.parentNode.target;
				if(''===this.value) img.removeAttribute('alt'); else img.alt=this.value;
			};
			o.lockField.onclick=function(){
				if(this.checked){
					var o=this.parentNode.parentNode.parentNode; 
					o.rate=o.widthField.value/o.heightField.value;
				}
			};

			var aClick=function(e){
				var op=this.parentNode.parentNode,
					o=op.target,
					id=this.id;
				if('resetWH'==id){
					o.removeAttribute('width');
					o.removeAttribute('height');
					o.style.width='';
					o.style.height='';
					op.widthField.value=o.width;
					op.heightField.value=o.height;
					op.rate=o.width/o.height;
				}else{ // if('align'==id.substring(0,5)){
					o.align='';
					o.style.display='';
					o.style.margin='';
					$.exeCmd('justifynone');
					if('alignC'==id){
						//o.style.display='block';
						//o.style.margin='5px auto';
						$.exeCmd('justifycenter');
					}else if('alignDefault'!=id){
						o.align=id.substring(5,20).toLowerCase();;
					}
				}
				return false;
			};
			var os = TQE._tags(o,'A'),i;
			for(i=os.length-1;i>=0;i--){
				os[i].onclick=aClick;
			}
			TQE._noSel(o);
		}
		if(img.parentNode && img.parentNode.nodeName==='A'){
			o.lastChild.lastChild.style.display='';
			o.linkField.value=img.parentNode.getAttribute('href');
			$.__selectOption(o.targetField, img.parentNode.getAttribute('target'));
		}else{
			o.lastChild.lastChild.style.display='none';
			o.linkField.value='';
			o.targetField.selectedIndex=0;
		}
		o.target=img;
		o.urlField.value=img.getAttribute('src');
		o.widthField.value=img.width;
		o.heightField.value=img.height;
		o.borderField.value=img.border;
		o.altField.value=img.alt;
		o.rate= (img.width>0 && img.height>0) ? img.width / img.height:1;
		$._popContext(o, img);
	};
	$._popContext=function(o,relateObj){
		if(TQE.activeContext != o )TQE._hideContext();
		var x,y,my
			p=TQE.pos($._editorObj()),
			p2=relateObj.getBoundingClientRect(),
			d=TQE._doc();

		o.style.display="block";
		x = Math.min( p.x+ p2.left , TQE._docScrollLeft() + d.clientWidth-o.offsetWidth-5);
		y =  10+p.y+ p2.bottom;
		my = TQE._docScrollTop() + d.clientHeight - o.offsetHeight-10; 
		if(y > my) y=p2.top+p.y-o.offsetHeight-10;
		if(y<TQE._docScrollTop())y=10+Math.max(p.y, TQE._docScrollTop());//	isIE? window.event.clientY : 
		o.style.left= x+'px';
		o.style.top= y+'px';
		TQE.activeContext=o;
		TQE.drag(o);
	};
	$._showTableContext=function(b){
		var i,btn, 
			ids=['TQEAddRowBefore','TQEAddRowAfter','TQEAddColBefore','TQEAddColAfter','TQEDeleteRow','TQEDeleteColumn','TQEDeleteTable'];
		for(i=ids.length;i>=0;i--){
			btn=TQE.find(ids[i])
			if(btn){
				btn.style.display = b ? '':'none'; 
			} 
		}
	};
	$._popTableContext=function(){
		var table,td=$.focusNode();
		while(td && td.nodeName !=='TD'){
			td = td.parentNode	 
		}
		if(!td){
			return $._showTableContext(0);
		}
		table=td.parentNode.parentNode.parentNode;
		var editorPos=TQE.pos($._editorObj()),
			tablePos=table.getBoundingClientRect(),
			tdPos=td.getBoundingClientRect(),
			btnTQEAddRowBefore=TQE.find('TQEAddRowBefore'),
			btnTQEAddRowAfter=TQE.find('TQEAddRowAfter'),
			btnTQEAddColBefore=TQE.find('TQEAddColBefore'),
			btnTQEAddColAfter=TQE.find('TQEAddColAfter'),
			btnTQEDeleteRow=TQE.find('TQEDeleteRow'),
			btnTQEDeleteColumn=TQE.find('TQEDeleteColumn'),
			btnTQEDeleteTable=TQE.find('TQEDeleteTable'),
			d=$._getDoc().documentElement,
			p1,p2;
		//��Ԫ���ڿɼ���Χ��
		if(tdPos.top+td.offsetHeight<0 || tdPos.left+td.offsetWidth<0 || tdPos.top>=d.clientHeight || tdPos.left>=d.clientWidth){
			return $._showTableContext(0);
		}
		if(!btnTQEAddRowBefore){
			btnTQEDeleteTable=TQE.CE('A','TQEDeleteTable','',true);
			btnTQEDeleteTable.title=resStrDeleteTable;

			btnTQEAddRowBefore=TQE.CE('A','TQEAddRowBefore','',true);
			btnTQEAddRowAfter=TQE.CE('A','TQEAddRowAfter','',true);
			btnTQEAddRowAfter.className=btnTQEAddRowBefore.className='eButtonAddRow';
			btnTQEAddRowBefore.title=btnTQEAddRowAfter.title=resStrInsertRow;

			btnTQEAddColBefore=TQE.CE('A','TQEAddColBefore','',true);
			btnTQEAddColAfter=TQE.CE('A','TQEAddColAfter','',true);
			btnTQEAddColAfter.className=btnTQEAddColBefore.className='eButtonAddCol';
			btnTQEAddColBefore.title=btnTQEAddColAfter.title=resStrInsertCol;

			btnTQEDeleteRow=TQE.CE('A','TQEDeleteRow','',true);
			btnTQEDeleteColumn=TQE.CE('A','TQEDeleteColumn','',true);
			btnTQEDeleteColumn.className=btnTQEDeleteRow.className=btnTQEDeleteTable.className='eButtonDeleteRC';
			btnTQEDeleteRow.title=resStrDeleteRow;
			btnTQEDeleteColumn.title=resStrDeleteCol;

			TQE._noSel(btnTQEAddRowBefore);
			TQE._noSel(btnTQEAddRowAfter);
			TQE._noSel(btnTQEAddColBefore);
			TQE._noSel(btnTQEAddColAfter);
			TQE._noSel(btnTQEDeleteRow);
			TQE._noSel(btnTQEDeleteColumn);
			TQE._noSel(btnTQEDeleteTable);

			btnTQEAddRowBefore.onclick=btnTQEAddRowAfter.onclick=btnTQEAddColBefore.onclick=btnTQEAddColAfter.onclick=btnTQEDeleteRow.onclick=btnTQEDeleteColumn.onclick=btnTQEDeleteTable.onclick=function(){
				if(!this.cell) return;
				var td=this.cell, 
					table=td.parentNode.parentNode.parentNode,
					$=this.editor;
				$.unselectAllCells();
				switch(this.id){
					case 'TQEAddRowBefore': 
						TQE.insertRow(td,1);
						break;
					case 'TQEAddRowAfter':
						TQE.insertRow(td,0);
						break;
					case 'TQEAddColBefore':
						TQE.insertCol(td,1);
						break;
					case 'TQEAddColAfter':
						TQE.insertCol(td,0);
						break;
					case 'TQEDeleteRow': 
						TQE.deleteRow(td);
						break;
					case 'TQEDeleteColumn':
						TQE.deleteCol(td);
						break;
					case 'TQEDeleteTable':
						table.parentNode.removeChild(table);
						break;
				}
				$.focus();
				$._popTableContext();
			};
		}else{
			$._showTableContext(1);
		}
		btnTQEAddRowBefore.cell=btnTQEAddRowAfter.cell=btnTQEAddColBefore.cell=btnTQEAddColAfter.cell=btnTQEDeleteRow.cell=btnTQEDeleteColumn.cell=btnTQEDeleteTable.cell=td;
		btnTQEAddRowBefore.editor=btnTQEAddRowAfter.editor=btnTQEAddColBefore.editor=btnTQEAddColAfter.editor=btnTQEDeleteRow.editor=btnTQEDeleteColumn.editor=btnTQEDeleteTable.editor=$;

		btnTQEAddRowAfter.style.left =  btnTQEAddRowBefore.style.left= btnTQEDeleteRow.style.left = (Math.min(editorPos.x+$._getDoc().body.offsetWidth-12,editorPos.x+tablePos.left+table.offsetWidth))+'px'; 
		btnTQEAddRowBefore.style.top=(editorPos.y+tdPos.top-5)+'px';
		btnTQEAddRowAfter.style.top=(editorPos.y+tdPos.top+td.offsetHeight-4)+'px';
		btnTQEDeleteRow.style.top=(editorPos.y+tdPos.top+ (td.offsetHeight >>> 1) -6)+'px';
		
		btnTQEAddColAfter.style.top = btnTQEAddColBefore.style.top= btnTQEDeleteColumn.style.top = (editorPos.y + Math.max(0,tablePos.top-12))+'px';
		btnTQEAddColBefore.style.left=(editorPos.x+tdPos.left-5)+'px';
		btnTQEAddColAfter.style.left=(editorPos.x+tdPos.left+td.offsetWidth-5)+'px';
		btnTQEDeleteColumn.style.left=(editorPos.x+tdPos.left+(td.offsetWidth>>1)-6)+'px';

		p1=TQE.pos(btnTQEAddRowBefore);
		p2=TQE.pos(btnTQEAddColAfter);
		btnTQEDeleteTable.style.top= Math.min(p2.y,p1.y-13) + 'px';
		btnTQEDeleteTable.style.left= Math.max(p2.x+10, p1.x) + 'px';

	};
	$.activeTable=null;

	$.unselectAllCells=function(){
		if(!$.activeTable) return;
		var table=$.activeTable,i,j;
		for(  i=table.rows.length-1; i>=0;i--){
			for(j=table.rows[i].cells.length-1;j>=0;j--){
				td=table.rows[i].cells[j];
				p=TQE.pos(td);
				$._selectCell(td, false );
			}
		}
		$.activeTable=null;
	};
	$.selectedCells=function(){
		if(!$.activeTable) return [];
		var table=$.activeTable,i,j,result=[];
		for(  i=table.rows.length-1; i>=0;i--){
			for(j=table.rows[i].cells.length-1;j>=0;j--){
				if(table.rows[i].cells[j].isSelected) result.push(table.rows[i].cells[j]);
			}
		}
		return result;
	};
	$._selectCell=function(cell,b){

		if(cell.isSelected==b) return;
		cell.className=b?'selectedTD':'';
		//if(b){
		//	var rang=document.createRange(), sel=$._getSelection();
		//	rang.selectNode(cell);
		//	sel.addRange(rang);
		//}
		cell.isSelected=b;
	};
	$._pos=function(e){
		var x=0,y=0,
			box = e.getBoundingClientRect(),
			d=$._getDoc(),
			dWin= $._editorObj().contentWindow,
			dRoot= d.compatMode != "CSS1Compat" ? d.body : d.documentElement;
		x = box.left - dRoot.clientLeft + (dWin.scrollX ? dWin.scrollX : dRoot.scrollLeft);
		y = box.top - dRoot.clientTop + (dWin.scrollY ? dWin.scrollY : dRoot.scrollTop);
		return {'x' : x, 'y' : y, left:x, top:y, width:e.offsetWidth, height:e.offsetHeight};
	};
	$._startSelectTd=function(e){
		var x,y,oldX,oldY,
			table, 
			d,dt,dl,
			tdP,td,
			startSelectCell=false,
			doc = $._getDoc(),
			docWin=doc.parentWindow,
			dRoot= doc.compatMode != "CSS1Compat" ? doc.body : doc.documentElement,
			docMM=doc.onmousemove,
			docMU=doc.onmouseup,
			docDS=doc.ondragstart,
			rootMU=document.onmouseup;
		e =$._realEvent(e);
		if(e.ctrlKey || e.altKey || e.button!= isIE ? 1:0) return;
		$.unselectAllCells();
		td =(isIE)? e.srcElement:e.target;
		while(td && td.nodeName!='BODY' && td.nodeName!='TD' ){
			td=td.parentNode;
		}
		if(!td || td.nodeName!=='TD') return;
		table=td.parentNode.parentNode.parentNode;
		tdP=$._pos(td);

		//d = (doc.compatMode != "CSS1Compat") ? doc.body : doc.documentElement;
		//dl=0;// (isIE || isOpera)? d.scrollLeft : docWin.scrollX;
		//dt=0;// (isIE || isOpera)? d.scrollTop : docWin.scrollY;
		oldX=isIE ? e.clientX + (docWin.scrollX ? docWin.scrollX : dRoot.scrollLeft) : e.pageX;
		oldY=isIE ? e.clientY + (docWin.scrollY ? docWin.scrollY : dRoot.scrollTop) : e.pageY;

		doc.onmousemove=function(docEvent){
			var x,y,dl,dt,i,j,x1,x2,y1,y2,p,cell,sel;
			//dl= (isIE || isOpera)? d.scrollLeft : docWin.scrollX;
			//dt= (isIE || isOpera)? d.scrollTop : docWin.scrollY;
			if(isIE){
				docEvent=this.parentWindow.event;
				x=docEvent.clientX + (docWin.scrollX ? docWin.scrollX : dRoot.scrollLeft);
				y=docEvent.clientY + (docWin.scrollY ? docWin.scrollY : dRoot.scrollTop);
			}else{
				x=docEvent.pageX;
				y=docEvent.pageY;
			}

			if(!startSelectCell && (x<tdP.x || x>tdP.x+tdP.width || y<tdP.y || y>tdP.y+tdP.height)){
				startSelectCell=true;
				$.activeTable=table;
			}
			if(! startSelectCell) return;
			
			sel= $._getSelection()
			if(sel.empty) sel.empty();	//ie6,7,8
			else if(sel.removeAllRanges) sel.removeAllRanges();

			x1=Math.min(oldX,x);
			x2=Math.max(oldX,x);
			y1=Math.min(oldY,y);
			y2=Math.max(oldY,y);
			for(i=table.rows.length-1; i>=0;i--){
				for(j=table.rows[i].cells.length-1;j>=0;j--){
					cell=table.rows[i].cells[j];
					p=$._pos(cell);
					$._selectCell(cell, p.x<=x2 && p.y<=y2 && p.x+p.width>=x1 && p.y+p.height>y1 );
				}
			}
		};

		doc.ondragstart=doc.onmouseup=document.onmouseup=function(){
			doc.onmousemove=docMM;
			doc.ondragstart=docDS,
			document.onmouseup=rootMU;
			startSelectCell=false;
			//return;
			//var i,cells = $.selectedCells(),
			//	rang,
			//	sel=$._getSelection();
			//if(startSelectCell){
			//	if(sel.removeAllRanges) sel.removeAllRanges();
			//	else if(sel.clear) sel.clear();	//ie6,7,8
			//}
			//for(i=cells.length-1;i>=0;i--){
			//	rang=document.createRange();
			//	rang.selectNode(cells[i]);
			//	sel.addRange(rang);
			//}
			//sel.selectAllChildren();

		};
	};
	//�ж��Ƿ���Ժϲ�ѡ�еĵ�Ԫ��
	$.canMergeSelectedCells=function(){
		if(!$.activeTable) return false;
		var table=$.activeTable, i,j, l,t,r,b, p,cell,count=0,cells = $.selectedCells();
		if(cells.length<2) return false;
		TQE.updateCellsStatus(table);
		l=cells[0].colStart;
		r=cells[0].colEnd;
		t=cells[0].rowStart;
		b=cells[0].rowEnd;
		for(i=cells.length-1;i>0;i--){
			l=Math.min(l,cells[i].colStart);
			t=Math.min(t,cells[i].rowStart);
			r=Math.max(r,cells[i].colEnd);
			b=Math.max(b,cells[i].rowEnd);
		}

		for(i=Math.min(b,table.rows.length-1); i>=0;i--){
			for(j=table.rows[i].cells.length-1;j>=0;j--){
				cell=table.rows[i].cells[j];
				if(!cell.isSelected && cell.colStart<=r && cell.colEnd>=l && cell.rowStart<=b && cell.rowEnd>=t){
					return false;
				}
			}
		}

		return true;
	
	};
	//�ϲ�ѡ�еĵ�Ԫ��
	$.mergeSelectedCells=function(){
		if(! $.canMergeSelectedCells() ) return false;

		var table=$.activeTable,i,j,l,t,p, s='', firstCell=null, rowSpan=1,colSpan=1,cell;
		for(  i=0; i<table.rows.length;i++){
			for(j=0;j<table.rows[i].cells.length;j++){
				if(table.rows[i].cells[j].isSelected){
					cell=table.rows[i].cells[j];
					s += cell.innerHTML;
					p=TQE.pos(cell);
					if(firstCell===null){
						firstCell=cell;						
						l=p.x;
						t=p.y;
						rowSpan=firstCell.rowSpan;
						colSpan=firstCell.colSpan;
					}else{
						if(p.y===t) colSpan+= cell.colSpan;
						else if(p.x===l) rowSpan+=cell.rowSpan;
						cell.innerHTML='';
						//cell.parentNode.removeChild(cell);
					}
				}
			}
		}
		firstCell.innerHTML=s.replace(/(&nbsp;|<br\s*\/?>)+$/,'$1');
		firstCell.rowSpan=rowSpan;
		firstCell.colSpan=colSpan;
		for(i=table.rows.length-1; i>=0;i--){
			p=table.rows[i];
			for(j=p.cells.length-1;j>=0;j--){
				if(p.cells[j].isSelected){
					cell= p.cells[j];
					if(cell!=firstCell){
						p.removeChild(cell);
					}
				}
			}
			if(p.cells.length<1)table.deleteRow(i);
		}
		TQE.updateCellsStatus(table);
	};
	$.splitCell=function(cell){
		if(!cell || cell.rowSpan<2 && cell.colSpan<2) return true;
		var table=cell.parentNode.parentNode.parentNode, i,j,k, row, td;
		TQE.updateCellsStatus(table);
		for(i=cell.rowStart;i<=cell.rowEnd;i++){
			row=table.rows[i];
			for(j=0;j<row.cells.length;j++){
				if(row.cells[j].colStart>cell.colEnd){
					break;
				}
			}
			k = cell.colSpan;
			if(i==cell.rowStart) k--;
			while( 0 < k-- ){
				td=row.insertCell(j);
				$._selectCell(td,cell.isSelected);
				td.style.backgroundColor=cell.style.backgroundColor;
			}
		}
		cell.rowSpan=1;
		cell.colSpan=1;

	};
	$._changed=function(){
		$._autoSave(1);
		$._showTableContext(false);
		var o;
		if('visual'===$.currentMode()){
			$._keepBookmark();
			$._hook();

			//if($.config.autoClean) doc.body.onpaste=function(){setTimeout($.cleanWord,200);};
			o=$._editorObj().contentWindow;
			TQE.addListener(o,'resize',$._popTableContext);
			if(isIE6){
				o=o.document.body;//��,ie6��iframe ���ܴ��� onfocus �� onblur �¼�
			}
		}else{
			o= $._editorObj();
			TQE.addListener(o,'mouseup',TQE.hidePop);
			o.onkeydown=function(e){return  $._call('keydown',e);}
			o.onkeyup=function(e){return $._call('keyup',e);}
		}

		//���¼�
		o.onfocus=function(e){
			$.hasFocus=true;
			if($.obj.onfocus) $.obj.onfocus();
			$._call('focus',e);
		};
		o.onblur=function(e){
			$.hasFocus=false; 
			if($.obj.onblur) $.obj.onblur();
			$._call('blur',e);
		};
	};
	$._realEvent=function(e){
		if(isIE) e = 'code'===$.currentMode() ? window.event :  $._editorObj().contentWindow.event; 
		return e;
	};
	$._call=function(eventName,e){
		var handler = $.config['on'+eventName];
		if('function'===typeof handler) return handler.call($,$._realEvent(e)); 
	};
	//mode ��ѡֵ[code|visual]
	$.changeMode=function(mode){//���ı༭��ʽ,��ֵʱ�Զ�ת��һ�ַ�ʽ
		var old=$.currentMode();
		if(old===mode) return;
		if(!mode){
			 mode = ('code'!=old) ? 'code' : 'visual';
		}
		if(mode==old) return;
		$.hasFocus=false;
		var eo=$._editorObj(), btn=$._button('mode'),h=eo.offsetHeight, editorHtm, visualMode='visual'===mode;
		if('code'===mode){
			$.unselectAllCells();
			h-=4;
			editorHtm='<table width=100% border=0 cellpadding=0 cellspacing=0><tr><td width=2></td><td><textarea id="'+$.objId+'_Editor" wrap="hard" style="height:'+h+'px"></textarea></td></tr></table>';
		}
		else if(visualMode){
			editorHtm='<IFRAME id="'+$.objId+'_Editor" frameborder=0 marginheight=0 marginwidth=0 style="height:'+h+'px" ></IFRAME>';
		}else{
			return;
		}
		$._autoSave(0);
		TQE._hideContext();

		var code = $.content();
		$._clientDiv.innerHTML=editorHtm;
		$._initContent(code);
		//$._setHeight(visualMode? h: h-4);

		if(btn) btn.id = visualMode?'eBTNmode':'eBTNmode-reset';
		$._updateToolbar(visualMode,['mode','fullscreen','help'], true);
	};
	//�Զ��Ű�
	$.typeset=function(){
		if('visual'!=$.currentMode()) return;
		$._showTableContext(0);
		var d = $._getDoc(),e,es,s,div,ep,fw;
		//ɾ������Ҫ�Ŀո�,��DIVתΪP
		s=d.body.innerHTML;
		s=s.replace(/&nbsp;/gi,' ').replace(/>\s+</gm,'> <').replace(/ +/gi,' ').replace(/<div/gi,'<p').replace(/<\/div/gi,'</p');
		d.body.innerHTML=s;
		//ɾ��CENTER
		es=TQE._tags(d,'CENTER');
		for(i=es.length-1;i>=0;i--){
			TQE.removeNode(es[i],false);
		}
		//ɾ������Ҫ��SPAN
		es=TQE._tags(d,'SPAN');
		for(i=es.length-1;i>=0;i--){
			e=es[i];
			fw=	e.style.fontWeight,c=e.style.color;
			if(''===e.innerHTML.replace(/<[a-z\/][^>]*>/gi,'') || ( fw=='' && c=='')){
				TQE.removeNode(e,false);
				continue;
			}
			e.removeAttribute('style');
			e.className=''
			e.removeAttribute('class');
			e.style.fontWeight=fw;
			e.style.color=c;
		}
		//�����ҳ��
		es=TQE._tags(d,'H3');
		for(i=es.length-1;i>=0;i--){
			if(es[i].className === 'splitPage') es[i].innerHTML=es[i].innerHTML.replace(/<\/?[a-zA-z][^>]*>/g,'');
		}
		//�������
		es=TQE._tags(d,'P');
		var emtpyReg=/^(&nbsp;| ){3,}/
		for(i=es.length-1;i>=0;i--){
			e=es[i];
			if(''==e.innerHTML.replace(/<[a-z\/][^>]*>/gi,'').replace( /( |\n|\r|\t|&nbsp;)/gi, '')){//ɾ���հ׵Ķ���
				TQE.removeNode(e,false);
				continue;
			} 
			e.removeAttribute('style');
			e.className=''
			e.removeAttribute('class');
			if(emtpyReg.test(e.innerHTML)){
				e.align='center';
			}else{
				e.removeAttribute('align');
				e.style.textIndent=$.config.typesetIndentParagraph ? '2em':'';
			}
		}
		//��������
		es=TQE._tags(d,'A');
		for(i=es.length-1;i>=0;i--){
			e=es[i];
			if(''==e.innerHTML.replace(/<[a-z\/][^>]*>/gi,'').replace( /( |\n|\r|\t|&nbsp;)/gi, '')){//ɾ���հ׵�����
				TQE.removeNode(e,false);
				continue;
			} 
			e.removeAttribute('style');
			e.className=''
			e.removeAttribute('class');
		}
		var isEmptyNode=function(node){
			var nn=node.nodeName, emptyReg=/^( |\n|\r|\t|&nbsp;)*$/gi;
			if(nn==='#text') return emptyReg.test(node.nodeValue)
			return 'BR'===nn || 'IMG'!==nn && ''==node.innerHTML.replace(/<[a-z\/][^>]*>/gi,'').replace(/(\s|&nbsp;)/gmi,'');
		},
		//�Ű�ͼƬ�·�������
		centerImageText=function(node){
			var ne,s;
			while(node && node.tagName!='BODY'){
				ne = node;
				while(ne = ne.nextSibling){
					if(isEmptyNode(ne)) continue;
					if(ne.nodeName==='#text'){
						div = TQE.CE('CENTER','',ne.nodeValue);
						//div.style.textAlign='center';
						node.parentNode.insertBefore(div,ne);
						TQE.removeNode(ne,true);
						return ;
					}else{
						if(ne.tagName=='CENTER') return;
						if(ne.tagName=='P') ne.align='center';
						else ne.style.textAlign='center';
						return ;
					}
				}
				node=node.parentNode;
			}
		};
		//�Ű�ͼƬ
		es=TQE._tags(d,'IMG');
		for(i=es.length-1;i>=0;i--){
			e=es[i];
			if(e.getAttribute('emot')) continue;//����
			if($.config.typesetImageTextCenter)centerImageText(e);
			e.removeAttribute('style');
			e.border= $.config.typesetImageBorder ? '1':'0';
			e.style.margin='5px';
			//e.style.margin='5px auto';
			//e.style.display='block';

			ep=e.parentNode;
			if(ep.tagName=='A' && ep.firstChild==ep.lastChild){
				e=ep;
				ep=ep.parentNode;
			}
			if(ep.tagName=='P' && ep.align=='center' || ep.tagName=='CENTER' || ep.style.textAlign=='center') continue;
			div = TQE.CE('CENTER');
			ep.insertBefore(div, e);
			div.appendChild(e);
		}
		//�Ű���Ƶ,flash��
		es=TQE._tags(d,'OBJECT');
		for(i=es.length-1;i>=0;i--){
			e=es[i];
			if($.config.typesetImageTextCenter)centerImageText(e);

			ep=e.parentNode;
			if(ep.tagName=='P' && ep.align=='center' || ep.tagName=='CENTER' || ep.style.textAlign=='center') continue;
			div = TQE.CE('CENTER');
			ep.insertBefore(div, e);
			div.appendChild(e);
		}
		//�Ű���Ƶ,flash��
		es=TQE._tags(d,'EMBED');
		for(i=es.length-1;i>=0;i--){
			e=es[i];
			if('OBJECT'===e.parentNode.tagName) continue;
			if($.config.typesetImageTextCenter) centerImageText(e);

			ep=e.parentNode;
			if(ep.tagName=='P' && ep.align=='center' || ep.tagName=='CENTER' || ep.style.textAlign=='center') continue;
			div = TQE.CE('CENTER');
			ep.insertBefore(div, e);
			div.appendChild(e);
		}
		$.cleanWord(true);
	};
	//force	����ֵ
	$.cleanWord=function(force){
		if('visual'!=$.currentMode()) return;
		var d = $._getDoc();
		if(!force){
			if(d.body.innerHTML.indexOf('mso') <0) return;
		}
		var cleanTag=function(tags, removeOnEmptyStyle){
			for(var e,es,s,i=tags.length-1; i>=0; i--){
				es=TQE._tags(d,tags[i]);
				for(i=es.length-1;i>=0;i--){
					e=es[i];
					e.removeAttribute('lang');
					s = e.getAttribute('style');
					if(null==s && removeOnEmptyStyle){
						TQE.removeNode(e,false);
						continue;
					}
					if(typeof s =='string' && s.indexOf('mso-')>=0){
						s+=';';
						s=s.replace(/mso\-[^;]+?;/gi,'').replace(/FONT\-FAMILY[^;]+?;/gi,'').replace(/\s+/g,' ').replace(/MARGIN\: 0cm 0cm 0pt/gi,'').replace(/<(\w[^>]*) style="([^"]*)"([^>]*)/gi,'');
						if(''==s || ' '==s){
							if(removeOnEmptyStyle){
								TQE.removeNode(e,false);
								continue;
							}
							e.removeAttribute('style');
						}else{
							e.setAttribute('style',s);
						}
					}
					s = e.getAttribute('class');
					if(null!=s){
						s=s.replace(/mso[a-z\-]+\s*/gi,'');
						if(''==s){
							e.removeAttribute('class');
						}else{
							e.setAttribute('class',s);
						}
					}
				}
			}
		};
		cleanTag(['SPAN'],true);
		cleanTag(['P','B','I','U','TD','TR','TABLE','DIV'],true);

		d.body.innerHTML=d.body.innerHTML.replace(/<o:[a-z0-9_]+[^>]*>/gi,'').replace(/<\/o:[a-z0-9_]+>/gi,'').replace(/<\?xml.+?\/>/gi,'');
		$.save();
	};
	//$._fixWidth=function(){
		//if('code'===$.currentMode())$._editorObj().style.width= $._containerDiv.offsetWidth-6+ 'px';
	//}
	// ���ñ༭���߶�,ֵ����1ʱ,��λ��, С��1ʱ�� �ٷֱ�, �� 0.5��ʾ50%
	// ����������0-1֮���С�������1������
	$.setWidth=function(value){
		//if(!isIE) TQE.delListener(window,'resize',$._fixWidth);
		if(value>1){
			$.config.width=value+'px';
			//$._containerDiv.onresize=null;
		}else{
			$.config.width=(value*100)+'%';
			//if(!isIE) TQE.addListener(window,'resize',$._fixWidth);
			
		}
		$._containerDiv.style.width= $.config.width;
		//$._fixWidth();
	};
	//���ñ༭���߶�,��λ����
	$.setHeight=function(value){
		if(value<9) value=9;
		$.config.height=value+'px';
		$._setHeight(value);
	};
	$._setHeight=function(value){
		if(value<2)value=2;
		$._editorObj().style.height= value+'px';
		if('visual'==$.currentMode()){
			value-=6;
			if(value<2) value=2;
			$._getDoc().body.style.minHeight=value+'px';
			if(isIE6)$._getDoc().body.style.height=value+'px';
		}
	};
	$._fixFull=function(){
		var d=TQE._doc(),e=$._containerDiv, h;;
		e.style.zIndex='10000';
		e.style.width=(d.clientWidth-2) + 'px';
		e.style.top = TQE._docScrollTop()+'px';
		e.style.left = TQE._docScrollLeft()+'px';

		h = d.clientHeight-$._toolbarDiv.offsetHeight
		$._setHeight(h- ('visual'==$.currentMode()?2:6));
		//return;
		//e=$._editorObj();
		//if('visual'==$.currentMode()){
		//	$.setHeight(h-2);
		//	e.style.height= h-2 + 'px';
		//	d = $._getDoc();
		//	if(d.body){
		//		h-=6;
		//		if(h<2)h=2;
		//		d.body.style.minHeight=h+'px';
		//		if(isIE6)d.body.style.height=h+'px'; 
		//	}
		//	//e.style.width='';
		//}else{
		//	//e.style.width=(d.clientWidth-6)+'px';
		//	e.style.height= h-7 + 'px';
		//}
	};
	$._button=function(name){
		var btns = TQE._tags($._toolbarDiv,'LI'),i;
		for(i=btns.length-1;i>=0;i--){
			if(name==btns[i].getAttribute('name') ) {//	
				return btns[i];
			}
		}
		return null;
	};
	$.fullScreen=function(b){
		var e=$._containerDiv,isVisualMode = 'visual'===$.currentMode(),o,btn;
		if(b===undefined || null===b){
			b = (e.style.position!='absolute');
		}else{
			if(b == (ec.style.position=='absolute'))return;
		}
		//�ƶ�IFrame�ᵼ�����ݶ�ʧ, ����ҪתΪ����ģʽ
		if(isVisualMode){
			$.changeMode('code');
		}
		if(b){
			document.body.parentNode.style.overflow = 'hidden';
			//e.parentNode.removeChild(e);
			document.body.appendChild(e);
			if(isVisualMode){
				if(isGecko){//����Firefox��bug��? 
					window.setTimeout(function(){$.changeMode('visual');},1);
				}else{
					$.changeMode('visual');
				}
			}
			e.style.position='absolute';
			e.style.top = '0px';
			$._fixFull();
			TQE.addListener(window,'resize',$._fixFull);

		}else{
			document.body.parentNode.style.overflow = 'auto';
			TQE.delListener(window,'resize',$._fixFull);
			e.style.position='';
			e.style.width=$.config.width;
			
			o=$._editorObj();
			//if(!isVisualMode)o.style.width=e.offsetWidth-6+'px';
			o.style.height=$.config.height;
			//e.parentNode.removeChild(e);
			$.obj.parentNode.insertBefore(e,$.obj);
			if(isVisualMode){
				$.changeMode('visual');
			}
			e.scrollIntoView(false);
		}

		//����ȫ����ťͼ��
		btn=$._button('fullscreen');
		if(btn){
			btn.id = b ? 'eBTNfullscreen-reset' : 'eBTNfullscreen';
		}
	};
	$.setContent=function(code){
		//code = ('code'==$.currentMode())? changeAsUbb(code):	changeAsHtml(code);
		if('code'===$.currentMode()){
			return $._editorObj().value=code;
		}
		if('function'===typeof $.config.onsetcontent) code=$.config.onsetcontent.call($,code); 
		$._getDoc().body.innerHTML=code;
	};
	$._initContent=function(code){
		if('code'!==$.currentMode()){
			if((isGecko || isOpera) && '<br>'==code.substring(0,4)){
				code=code.substr(4); // �ֵֹ�Firefox
			}
			if(isGecko) code = code.replace(/<td([^>]*)>\s*<\/td/ig, '<td$1 ><br></td');
			var s= '',//isIE6 ?'':'<!DOCTYPE html>',
				d = $._getDoc(),
				h;
			s+='<html><head><style>\nbody{border:none;height:100%;margin:0;padding:2px;overflow:auto;font-size:12px;';//Ĭ�������С
			if($.config.bgColor || $.config.textColor){
				if($.config.bgColor) s+='background-color:'+$.config.bgColor+';';
				if($.config.textColor) s+='color:'+$.config.textColor+';';
			}
			s +='}\nimg{cursor:default;}\n.selectedTD{background:#3399ff !important;}\n.splitPage{display:block;width:98%;!important;border-top:#999 1px solid;border-bottom:#999 1px solid;background:#FFFFFF; color:#000000;font-weight:bold;text-align:center;margin:30px auto 5;padding-top:30px;clear:both;}\n.splitPage:before{content:"��ҳ����:";}</style>\n';
			if($.config.baseHref!=''){
				s +='<base href="'+$.config.baseHref+'" />';
			}
			if ($.config.css) {
				s += '<link href="' +$.config.css+ '" rel="stylesheet" type="text/css">\n';
			}
			s+='</head><body contenteditable="true" >'+ code+'</body></html>';// 
			d.open();
			d.write(s);
			d.close();
			h=Math.max(20,$._editorObj().offsetHeight-6);
			if(d.body){
				d.body.style.minHeight=h+'px';
				if(isIE6)d.body.style.height=h+'px';
			}
		}
		$.setContent(code);
		$._changed();
	};
	$.setOpacity=function(e, opacity) {
		if (isIE) {// && !isIE9   
			e.style.filter = (opacity == 100) ? '' : 'gray() alpha(opacity=' + opacity + ')';
		} else {
			e.style.opacity = (opacity == 100) ? '' : opacity/100;
		}
    };
	$.inArray=function(e,arr){
		for(var i=arr.length-1;i>=0;i--) if(e==arr[i]) return true;
		return false;
	};
	$.content=function(){
		var m=$.currentMode(), cells,i,code;
		if('visual'===m){
			cells = $.selectedCells();
			for(i=cells.length-1;i>=0;i--){
				cells[i].className='';
			}
			code= $._getDoc().body.innerHTML;
			for(i=cells.length-1;i>=0;i--){
				cells[i].className='selectedTD';
			}
			if((isGecko || isOpera) && '<br>'==code.substring(0,4)) code=code.substr(4);//�ֵֹ�Firefox
			if(''==code.replace(/<\/?(div|p|font|b|s|u|i|strong|em|strike)(?![a-z0-9])[^>]*>|&nbsp;|\s/ig , '')) code= '';
			if('function'===typeof $.config.ongetcontent) code=$.config.ongetcontent.call($,code); 
			else code=code.replace(/\s*<(p|table|tr|hr|div|object|ol|li|ul|h[1-6])(?![a-z])/ig,'\n<$1')
						.replace(/<\/(p|table|div|object|ol|ul|h[1-6])>\s*/ig,'</$1>\n')
						.replace(/(<(?:br|hr)(?![a-z0-9])[^>]*>)\s*/ig,'$1\n')
						.replace(/^\s*/,'').replace(/\s*$/,'');
		}else{
			code= $._editorObj().value;
		}
		return code;
	};
	$._autoSave=function(b){
		var e = $._editorObj(), func = b? TQE.addListener : TQE.delListener;
		if (isIE ){
			func(e,'beforedeactivate',$.save);
		}else if('code'==$.currentMode()){ 
			func(e,'blur',$.save);
		}else{
			func(e.contentWindow,'blur',$.save);
		}
	};
	$.save=function(){
		$.obj.value=$.content();
	};
	$._intEdit=function(obj){
		if(obj.isIniEdit)return;
		obj.isIniEdit=true;
		obj.style.imeMode='disabled';
		obj.onkeypress=function(){
			if(45==event.keyCode){
				this.value = ('-'==this.value.charAt(0))? this.value.substr(1) : '-'+this.value;
				return false; 
			}
			return event.keyCode>=48&&event.keyCode<=57;
		};
		obj.onpaste=function(){return !clipboardData.getData('text').match(/\D/);};
		obj.ondragenter=function(){return ! /\D/.test(event.dataTransfer.getData('text'));};
		obj.oldBlur=obj.onblur;
		obj.onblur=function(){
			if(''==this.value) return; 
			this.value=parseInt(this.value,10);
			var max = this.getAttribute('max'), 
				min = this.getAttribute('min');
			if(max!==undefined && parseInt(this.value,10)>parseInt(max,10))this.value=parseInt(max,10);
			if(min!=undefined && parseInt(this.value,10)<parseInt(min,10))this.value=parseInt(min,10); 
			if(this.oldBlur)this.oldBlur();
		};
	};
	//------------ִ�г�ʼ��
	if(isIE6 || isIE && (document.documentMode==7 || document.compatMode != "CSS1Compat")){//IE6�������ͼ�����ģʽ
		var _topTableNode=function(node){
			var  reault=null;
			while(node){
				 if(node.tagName=='TABLE'){
					reault=node;
				 }
				 node = node.parentNode;
			}
			return reault;
		},
		t = _topTableNode(document.scripts(document.scripts.length-1));
		if(t && _topTableNode(TQE.find(objId))===t ){
			window.setTimeout($.init,20);
		}else{
			$.init();
		}
	}else{
		$.init();
	}
};
TQE.url='./';//�༭�������Դ��ַ
TQE.version = version; //�༭���汾
//-- init TQE resources
if(TQE.url==='./'){
	var scripts = document.scripts || document.getElementsByTagName('SCRIPT'), //firefox
		tqJSUrl = scripts[scripts.length-1].getAttribute('src'), 
		skinFind= /[\?&]skin=([a-z0-9_A-Z]+)/.exec(tqJSUrl),
		l=window.location,
		urlHost=l.protocol+'//'+l.host + (l.port?':'+l.port:'')+'/',
		skinUrl='';
	TQE.url = tqJSUrl.substring(0,tqJSUrl.toLowerCase().indexOf('tqeditor.'));
	if(TQE.url.indexOf(urlHost)===0) TQE.url=TQE.url.substring(urlHost.length-1, TQE.url.length);
	document.write('<link href="'+TQE.url+(skinFind ? 'skin/'+skinFind[1]+'/':'')+'TQEditor.css" id="TQEditorSkinCSS" rel="stylesheet" type="text/css" onload="TQE.checkSkin(this)" />');
}
//����
TQE.config={
	"paragraphs" : [
		["����",	"p"],
		["����1",	"h1"],
		["����2",	"h2"],
		["����3",	"h3"],
		["����4",	"h4"],
		["����5",	"h5"],
		["����6",	"h6"],
		["Ԥ����",	"pre"]
		],
	"fontnames" : [
		["����",            "����"],
		["������",          "������"],
		["����_GB2312",     "����_GB2312"],
		["����",            "����"],
		["΢���ź�",        "΢���ź�"],
		["����_GB2312",     "����_GB2312"],
		["Arial",           "arial, helvetica, sans-serif"],
		["Courier New",     "courier new, courier, mono"],
		["Georgia",         "Georgia, Times New Roman, Times, Serif"],
		["Tahoma",          "Tahoma, Arial, Helvetica, sans-serif"],
		["Times New Roman", "times new roman, times, serif"]
		],
	"fontsizes": [
		//[text,preview,	value]
		["С",	"xx-small",	1],
		["��",	"x-small",	2],
		["��",	"small",	3],
		["�ش�","medium",	5]
		],
	"toolbar":'default',
	"toolbarRight":[],
	'showToolbarGroupHandle':true, //��ʾ��������ǰ��ͼ
	
	//ģ��
	//����ý��***ʱ��ģ�����,
	//flash	�� {$url} {$width} {$height}
	'tplFlash':'<embed  src="{$url}" width="{$width}" height="{$height}" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" wmode="transparent" allowscriptaccess="always"></embed>',
	//flv��Ƶ �� {$url} {$width} {$height} {$auto_start}
	'tplFlv':'<embed src="'+TQE.url+'flvPlayer.swf" flashvars="vcastr_file={$url}&IsAutoPlay={$auto_start}" width="{$width}" height="{$height}" quality="high" bgcolor="#0E0E0E" name="play" align="middle" allowscriptaccess="sameDomain" allowfullscreen="true" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" wmode="opaque" />',
	//rm��Ƶ �� {$url} {$width} {$height} {$auto_start}
	'tplRm':'<embed  src="{$url}" width={$width} height={$height} autostart="{$auto_start}"type="audio/x-pn-realaudio-plugin" console="Clip1" controls="ImageWindow" ></embed>',
	//windows��ý���ļ� �� {$url} {$width} {$height} {$auto_start}
	'tplVideo':'<object classid="clsid:6BF52A52-394A-11d3-B153-00C04F79FAA6"  type="application/x-ms-wmp" width="{$width}" height="{$height}"><PARAM NAME="URL" VALUE="{$url}" /><PARAM NAME="autoStart" VALUE="{$auto_start}" /><PARAM NAME="invokeURLs" VALUE="false"><PARAM NAME="playCount" VALUE="100"><PARAM NAME="Volume" VALUE="100"><PARAM NAME="defaultFrame" VALUE="datawindow"></object>',
	//���ֲ��� �� {$url} {$auto_start} 
	'tplMusic':'<object id="tqeMP"  classid="clsid:6BF52A52-394A-11d3-B153-00C04F79FAA6"  type="application/x-ms-wmp" width="230" height="64"><PARAM NAME="URL" VALUE="{$url}" /><PARAM NAME="autoStart" VALUE="{$auto_start}" /><PARAM NAME="invokeURLs" VALUE="false"><PARAM NAME="playCount" VALUE="100"><PARAM NAME="Volume" VALUE="100"><PARAM NAME="defaultFrame" VALUE="datawindow"></object>',
	//��ҳ������
	'tplSplitPage':'<h3 class="splitPage" name="split_page"></h3>',

	'uploadFieldName':'file',//�����ϴ�ʱ���ϴ�����

	'imageUploadUrl':'',//�����ϴ�ͼƬ��ַ, ���ձ�ʾ�������ϴ�, ��#��ͷ, ��ʾ����ϴ���Ť��Ľ�����ַ
	'imageFileTypes':'*.jpg;*.gif;*.png;*.jpeg',//�����ϴ���ͼƬ��׺
	'imageFileSize':'2MB',//�����ϴ��ļ���С
	'directInsertUploadImage':false,//ֱ�ӽ��ϴ���ͼƬ����༭����
	
	'linkUploadUrl':'',//�����ϴ���ַ, ���ձ�ʾ�������ϴ�, ��#��ͷ, ��ʾ����ϴ���Ť��Ľ�����ַ
	'linkFileTypes':'*.*',//�����ϴ��ĺ�׺
	'linkFileSize':'2MB',//�����ϴ��ļ���С

	'flashUploadUrl':'',//�����ϴ�Flash��ַ, ���ձ�ʾ�������ϴ�, ��#��ͷ, ��ʾ����ϴ���Ť��Ľ�����ַ
	'flashFileTypes':'*.swf',//�����ϴ���ͼƬ��׺
	'flashFileSize':'2MB',//�����ϴ��ļ���С

	'musicUploadUrl':'',//�����ϴ���Ƶ�ļ���ַ, ���ձ�ʾ�������ϴ�, ��#��ͷ, ��ʾ����ϴ���Ť��Ľ�����ַ
	'musicFileTypes':'*.mp3;*.wma',//�����ϴ���Ƶ�ļ���׺
	'musicFileSize':'2MB',//�����ϴ��ļ���С

	'videoUploadUrl':'',//�����ϴ���Ƶ�ļ���ַ, ���ձ�ʾ�������ϴ�, ��#��ͷ, ��ʾ����ϴ���Ť��Ľ�����ַ
	'videoFileTypes':'*.flv;.mp4;*.mkv;*.wmv;*.asf;*.avi;*.rmvb;*.rm',//�����ϴ����ļ���׺
	'videoFileSize':'2MB',//�����ϴ��ļ���С

	'enter2br':false,
	//'formatHTML':true,//��HTMLԴ���������(����������)
	//'height_delta':20,
	'width':'auto',
	'height':'auto',
	'autoClean':true,//��ִ��ճ��ʱ,�Զ�����word��ص���������
	'faceCount':55,//��������, ��faceĿ¼��
	'css':'',//��Ƹ����������ʽ��ַ
	'bgColor':'',  //���ӻ��༭ʱ, �༭������ɫ
	'textColor':'',	//���ӻ��༭ʱ, ����ɫ

	'typesetImageTextCenter':true,//�Զ��Ű�ͼƬ�·������־���
	'typesetIndentParagraph':true,//�Զ��Ű������������2��
	'typesetImageBorder':true,//�Զ��Ű�ͼƬ��߿�

	'baseHref':'',//���������ַ�Ļ���ַ
	'resize':true,//�����޸ı༭���ߴ�
	
	'onfocus':null,//��ȡ����ʱ�Ļص�, function(event){}
	'onblur':null,//ʧȥ����ʱ�Ļص�, function(event){}
	'onkeydown':null,//���û����¼��̰���ʱ����, function(event){}
	'onkeyup':null,//���û��ͷż��̰���ʱ����, function(event){}
	
	'oninit':null,//����ִ�г�ʼ��ʱ�Ļص�����, function(){}
	'ongetcontent':null,//ȡ����ʱ���¼��ص�����, function(code){}
	'onsetcontent':null,//���ô���ʱ���¼��ص�����, function(code){}
	//'onuploaded':null,//�ϴ��ɹ���Ļص�����
	'debug':false
};

TQE.find=function(id){return document.getElementById(id);};
TQE.CE=function(tag,id, htm, append){
	var o=document.createElement(tag);
	if(id)o.id=id;
	if(htm)o.innerHTML=htm;
	if(append) document.body.appendChild(o);
	return o;
};
TQE.loadSkin=function(skin){
	var e = TQE.find('TQEditorSkinCSS');
	e.href= TQE.url+(skin &&'default'!=skin ? 'skin/'+skin+'/':'')+'TQEditor.css';
};
TQE.checkSkin=function(linkObj){
	var i,objs=document.styleSheets;
	for(i=objs.length-1;i>=0;i--){
		if(objs[i].href==linkObj.href){
			if(isOpera && objs[i].cssRules.length<1 || !isOpera&&objs[i].rules.length<1) linkObj.href= TQE.url+'TQEditor.css';
			break;
		}
	}
};
TQE._doc=function(){return (document.compatMode != "CSS1Compat") ? document.body : document.documentElement;};
TQE._docScrollTop=function(){
	return window.scrollY ? window.scrollY : TQE._doc().scrollTop ;
};
TQE._docScrollLeft=function(){
	return window.scrollX ? window.scrollX : TQE._doc().scrollLeft ;
};
TQE.pos=function(e){
	var x = 0, y = 0;
	//if (e.getBoundingClientRect) {
		var box = e.getBoundingClientRect(),
			d=TQE._doc();
		x = box.left + TQE._docScrollLeft() - d.clientLeft;
		y = box.top + TQE._docScrollTop() - d.clientTop;
	//} else {
	//	x = e.offsetLeft;
	//	y = e.offsetTop;
	//	var parent = e.offsetParent;
	//	while (parent) {
	//		x += parent.offsetLeft;
	//		y += parent.offsetTop;
	//		parent = parent.offsetParent;
	//	}
	//}
	return {'x' : x, 'y' : y, left:x, top:y, width:e.offsetWidth, height:e.offsetHeight};
};
TQE._hideContext=function(){
	if(TQE.activeContext){
		TQE.activeContext.style.display="none";
		TQE.activeContext=null;
	}
};

TQE.pop=function(obj,relateObj){
	if(TQE.activePop==obj){
		TQE.hidePop();
		return false;
	}
	TQE._hideContext();
	TQE.hidePop();
	TQE.activePop=obj;
	obj.relateObj=relateObj;
	relateObj.className='eButtonDropDown';

	obj.style.display='block';
	var p=TQE.pos(relateObj),
		l=p.x,
		t= 1+relateObj.clientHeight +p.y,
		dl = TQE._docScrollLeft(),
		w =  TQE._doc().clientWidth;
	if(l+obj.offsetWidth>dl+w)l-=obj.offsetWidth-relateObj.offsetWidth;
	obj.style.left= l + 'px';
	obj.style.top = t+ 'px';
	return true;
};
TQE.hidePop=function(){
	if(!TQE.activePop) return;
	var o=TQE.activePop;
	TQE.activePop=null;
	o.style.display='none';
	if(o.relateObj && o.relateObj.className=='eButtonDropDown') o.relateObj.className='eButton';
	
	o=TQE.find('ColorPicker');
	if(o)o.style.display='none';
};
TQE.delListener=function(e,eventName,func){
	if(e.removeEventListener) {
		e.removeEventListener(eventName, func, false );
	}else{
		e.detachEvent('on'+eventName, e['e'+eventName+func]);
	}
};
TQE.addListener=function(e,eventName,func){
	if(e.addEventListener) {
		e.addEventListener( eventName, func, false );
	}else{
		e['fn'+func]=func;
		e['e'+eventName+func]=function(et){e['fn'+func](et)};
		e.attachEvent('on'+eventName, e['e'+eventName+func]);//
	}
};
TQE._tags=function(domNode, tagName){
	return domNode.getElementsByTagName(tagName); 
};
	
TQE.listenTags=function(node,tag,eventname,func){
	var ds = TQE._tags(node,tag),i;
	for(i=ds.length-1;i>=0;i--){
		TQE.addListener(ds[i],eventname,func);
	}
};
TQE._noSel=function(o, tagsName){
	var noSel=function(){
		if(isIE){
			var tn=event.srcElement.tagName;
			return 'INPUT'==tn || 'TEXTAREA'==tn;
		}
		return false;
	};
	if(tagsName){
		var a=tagsName.split(','),nodes,i,j;
		for(j=a.length-1;j>=0;j--){
			nodes = TQE._tags(o,a[j]);
			for(i=nodes.length-1; i>=0;i--){
				nodes[i].unselectable='on';
				nodes[i].onselectstart = noSel
			}
		}
	}else{
		o.unselectable='on';
		o.onselectstart = noSel;
	}
};
TQE._colorPickerUI=function(){
	var ColorHex=['00','33','66','99','CC','FF'],
		SpColorHex=['FF0000','00FF00','0000FF','FFFF00','00FFFF','FF00FF'],
		w=191,//1+19*10
		colorTable='<table width="'+w+'" cellpadding=0 ',i,j,k;
	colorTable += isIE ? 'cellspacing=1 border=0 bgcolor="#000000;">':'cellspacing=0 border=1 style="border-collapse:collapse;" >';
	for (i=0;i<2;i++){
		for (j=0;j<6;j++){
			colorTable += '<tr height=10>';
			//colorTable += '<td bgcolor="#000" name="'+name+'" ></td>';

			if (i>0) colorTable += '<td bgcolor="#'+SpColorHex[j]+'" name="'+name+'"></td>';
			else colorTable += '<td bgcolor="#'+ColorHex[j]+ColorHex[j]+ColorHex[j]+'" ></td>';

			//colorTable += '<td bgcolor="#000"></td>';
			for (k=0;k<3;k++){
				for (l=0;l<6;l++){
					//����ɫ
					colorTable += '<td bgcolor="#'+ColorHex[k+i*3]+ColorHex[l]+ColorHex[j]+'" ></td>';
					//����ɫ
					//var c1 = 4+i-k*2;
					//var c2 = i*j + (1-i)*(5-j);// i?j:5-j
					//var c3 = k!=1? 5-l:l;      //k!=1? 5-l:l;
					//colorTable += '<td bgcolor="#'+ColorHex[c1]+ColorHex[c2]+ColorHex[c3]+'"></td>';
				}
			}
			colorTable += '</tr>';
		}
	}
	colorTable +='</table>'+
			'<div style="padding:2px 2px 1px 2px;line-height:25px;"><div style="height:25px;width:60px;display:block;float:left; margin-right:5px;" name=colorPreview></div><span name="colorValue" style="width:50px; display:inline-block;"></span><input type="button" value="'+resStrNoColor+'" /></div>';
	return colorTable;
};
TQE.pickColor=function(callbackFunc,x,y){
	var o=TQE.find('ColorPicker');
	if(null==o){
		o = TQE.CE('DIV','ColorPicker', TQE._colorPickerUI() ,true);
		o.className='ePopPanel';
		o.style.zIndex=99999999;
		o.onmouseover=function(){this.active=true;};
		o.onmouseout=function(){this.active=false;};

		TQE._noSel(o);//,'TABLE,TR,TD,DIV');
		TQE.listenTags(o,'TD','mouseover',function(){
			var self=this, 
				p=this.parentNode.parentNode.parentNode.parentNode.lastChild,
				iid = window.setInterval(function(){self.style.backgroundColor=(self.style.backgroundColor!='')?'':'#FFFFFF'},500);
			p.firstChild.style.backgroundColor=this.bgColor;
			p.firstChild.nextSibling.innerHTML=this.bgColor;

			self.style.backgroundColor='#FFFFFF';
			self.onmouseout=function(){window.clearInterval(iid);this.style.backgroundColor='';}
		});
		o.lastChild.lastChild.onclick=function(){this.parentNode.parentNode.style.display='none';};
		TQE.listenTags(o,'TD','mouseup',function(){this.parentNode.parentNode.parentNode.parentNode.style.display='none';});
		TQE.addListener(document,'mousedown',function(){var o=TQE.find('ColorPicker');if(!o.active)o.style.display='none';});
	}
	var ds = TQE._tags(o,'TD');
	for(var i=ds.length-1;i>=0;i--){
		ds[i].onmousedown=function(e){callbackFunc(this.bgColor); e = e || event; e.cancelBubble=true; }
	}
	o.lastChild.lastChild.onclick=function(e){callbackFunc(''); e = e || event; e.cancelBubble=true;};

	o.style.left=x+'px';
	o.style.top=y+'px';
	o.style.display='block';
};
TQE.drag=function(o){
	o.style.position='absolute';
	o.onmousedown=function(a){
		var x,y,obj,n,
			p=TQE.pos(this),
			scrollLeft = TQE._docScrollLeft(), 
			scrollTop =TQE._docScrollTop(), 
			documentMM=document.onmousemove,
			documentMU=document.onmouseup;
		if(isIE){
			a=window.event;
			obj=a.srcElement;
			x=scrollLeft + a.clientX-p.x;
			y=scrollTop + a.clientY-p.y;
		}else{
			obj=a.target;
			x=a.pageX-p.x;
			y=a.pageY-p.y;
		}
		n=obj.tagName;
		if(obj!=this && ( n=='INPUT' || n=='TEXTAREA'|| n=='BUTTON'|| n=='SELECT')) return;

		if(o.setCapture)
			o.setCapture();
		else if(window.captureEvents)
			window.captureEvents(Event.MOUSEMOVE|Event.MOUSEUP);

		document.onmousemove=function(a){
			if(isIE){
				a=window.event;
				a.pageX=scrollLeft + a.clientX;
				a.pageY=scrollTop + a.clientY;
			}
			var tx=a.pageX-x, ty=a.pageY-y;
			o.style.left=tx +'px';
			o.style.top=ty +'px';
			o.style.cursor='move';
		};

		document.onmouseup=function(){
			if(o.releaseCapture)
				o.releaseCapture();
			else if(window.releaseEvents)
				window.releaseEvents(Event.MOUSEMOVE|Event.MOUSEUP);
			document.onmousemove=documentMM;
			document.onmouseup=documentMU;
			o.style.cursor='';
		};
	};
};
TQE.clone=function(srcObj){   
    if(typeof(srcObj) != 'object') return srcObj;   
    if(srcObj == null) return srcObj;   
    var i, newObj = '[object Array]'==Object.prototype.toString.call(srcObj)? [] : {};    
    for(i in srcObj) newObj[i] = TQE.clone(srcObj[i]);    
    return newObj;   
}; 
TQE.removeNode=function(node, delChild){
	if(node.removeNode) return node.removeNode(delChild); //ie
	if(delChild) return node.parentNode.removeChild(node);
	while(node.firstChild) node.parentNode.insertBefore(node.firstChild,node);
	node.parentNode.removeChild(node);

};
//table functions
TQE.updateCellsStatus=function(table){
	var rows=table.rows,cells,cells2,td,td2, i,j,k,l;
	for(i=0;i<rows.length;i++){
		k=0;
		for(j=0;j<rows[i].cells.length;j++){
			td= rows[i].cells[j];
			td.colStart=k;
			k += td.colSpan;
			td.colEnd= k-1;

			td.rowStart=i;
			td.rowEnd= i+td.rowSpan-1;
		}
	}
	//fix colIndex
	for(i=0;i<rows.length;i++){
		cells= rows[i].cells;
		for(j=0;j<cells.length;j++){
			if(cells[j].rowSpan<2) continue;
			for(k=1;k<cells[j].rowSpan; k++){
				if(!rows[i+k]) break;
				cells2=	rows[i+k].cells;
				td= cells[j];
				for(l=0;l<cells2.length;l++){
					td2= cells2[l];
					if(td2.colStart >= td.colStart){
						td2.colStart += td.colSpan;
						td2.colEnd += td.colSpan;
					}
				}
			}
		}
	}
	//fix rowIndex
	for(i=rows.length-1;i>=0;i--){
		cells= rows[i].cells;
		var minrs=cells[0].rowSpan;
		for(j=cells.length-1;j>0;j--){
			minrs = Math.min(minrs,cells[j].rowSpan);
			if(minrs<2)break;
		}
		if(--minrs > 0){
			for(j=cells.length-1;j>=0;j--){
				cells[j].rowSpan-=minrs;
				cells[j].rowEnd -=minrs;
			}
			for(k=i-1;k>=0;k--){
				for(l=rows[k].cells.length-1;l>=0;l--){
					if(rows[k].cells[l].rowEnd>cells[0].rowStart){
						rows[k].cells[l].rowSpan-=minrs;
						rows[k].cells[l].rowEnd-=minrs;
					}
				}
			}

		}
	}
};
TQE.insertCol=function(cell, bBegin){
	var table=cell.parentNode.parentNode.parentNode,cells,td,td2, i,j, colIndex;
	TQE.updateCellsStatus(table);
	colIndex= bBegin?cell.colStart : cell.colEnd+1;
	for(i=table.rows.length-1;i>=0;i--){
		cells= table.rows[i].cells;
		td=null;
		for(j=Math.min(colIndex, cells.length-1);j>=0;j--){
			if(cells[j].colStart<=colIndex){
				if(cells[j].colStart==colIndex){
					td=table.rows[i].insertCell(j);
				}else if(cells[j].colEnd==colIndex-1){
					td=table.rows[i].insertCell(j+1);
				}else if(cells[j].colSpan>1 && cells[j].colEnd>=colIndex){
					cells[j].colSpan++;
				}else if(j==cells.length-1 && cells[j].colEnd<colIndex){
					 td=table.rows[i].insertCell(j+1);
				}
				if(td){
					if(bBegin){
						td2=td.nextSibling;
						if(td2 && td2.rowSpan<2 && td2.colStart==colIndex){
							td.style.backgroundColor = td2.style.backgroundColor;
						}
					}else{
						td2 = td.previousSibling;
						if(td2 && td2.rowSpan<2 && td2.colEnd==colIndex-1){
							td.style.backgroundColor = td2.style.backgroundColor;
						}
					}
				}
				break;
			}
		}
	}
};
TQE.insertRow=function(cell, bBegin){
	var table=cell.parentNode.parentNode.parentNode,cells, row, cells2, i,j,k=0, rowIndex;
	TQE.updateCellsStatus(table);
	rowIndex= bBegin?cell.rowStart : cell.rowEnd+1;
	for(i=Math.min(table.rows.length-1, rowIndex-1);i>=0;i--){
		cells= table.rows[i].cells;
		for(j=cells.length-1;j>=0;j--){
			if(cells[j].rowEnd>=rowIndex){
				cells[j].rowSpan++;
				k+=cells[j].colSpan; 
			}
		}
	}
	cells= table.rows[0].cells;
	for(j=cells.length-1;j>=0;j--){
		k -= cells[j].colSpan; 
	}
	row = table.insertRow(rowIndex);
	while(k++ <0){
		row.insertCell(0).innerHTML='&nbsp;';
	}

	TQE.updateCellsStatus(table);
	cells = row.cells;
	cells2 = table.rows[ bBegin ? rowIndex+1 : rowIndex-1].cells;
	j=cells2.length-1;
	for(i=cells.length-1;i>=0;i--){
		while(j>=0 && cells2[j].colStart > cells[i].colStart ) j--;
		if(j>=0 && cells2[j].colStart == cells[i].colStart && cells2[j].colEnd == cells[i].colEnd ){
			cells[i].style.backgroundColor = cells2[j].style.backgroundColor; 
		}
	}
};

TQE.deleteCol=function(cell){
	var table=cell.parentNode.parentNode.parentNode,rows=table.rows,td, i,j, start,end;
	TQE.updateCellsStatus(table);
	start=cell.colStart;
	end=cell.colEnd;
	for(i=rows.length-1;i>=0;i--){
		for(j=Math.min(end,rows[i].cells.length-1);j>=0;j--){
			td=rows[i].cells[j];
			if(td.colEnd>=start && td.colStart<=end){
				if(td.colStart>=start && td.colEnd<=end) rows[i].removeChild(td);
				else if(td.colStart<start && td.colEnd>end) td.colSpan -= 1+end-start;
				else if(td.colStart<start && td.colEnd>=start) td.colSpan -= td.colEnd-start+1;
				else if(td.colStart<=end && td.colEnd>end){ td.colSpan -= 1+end-td.colStart; td.innerHTML='';}
			}
		}
		if(rows[i].cells.length<1) rows[i].parentNode.removeChild(rows[i]);
	}
	if(rows.length<1) table.parentNode.removeChild(table);
};
TQE.deleteRow=function(cell){
	var table=cell.parentNode.parentNode.parentNode,rows=table.rows,nextRow,td,td2, i,j,k, start,end;
	TQE.updateCellsStatus(table);
	start=cell.rowStart;
	end=cell.rowEnd;
	nextRow=rows[end+1];
	for(i=Math.min(end,rows.length-1);i>=0;i--){
		for(j=rows[i].cells.length-1;j>=0;j--){
			td=rows[i].cells[j];
			if(td.rowEnd>=start && td.rowStart<=end) {
				if(td.rowStart>=start && td.rowEnd<=end) rows[i].removeChild(td);
				else if(td.rowStart<start && td.rowEnd>end) td.rowSpan -= 1+end-start;
				else if(td.rowStart<start && td.rowEnd>=start) td.rowSpan -= td.rowEnd-start+1;
				else if(td.rowStart<=end && td.rowEnd>end){
					if(nextRow){ 
						for(k=0;k<nextRow.cells.length;k++){
							if(nextRow.cells[k].colStart>td.colEnd){
								break;
							}
						}
						td2=nextRow.insertCell(k);
						td2.rowSpan=1+end-td.rowStart;
						td2.colSpan=td.colSpan;
						td2.style.backgroundColor=td.style.backgroundColor;
					}
					rows[i].removeChild(td);
				}
			}
		}
		if(rows[i].cells.length<1) rows[i].parentNode.removeChild(rows[i]);
	}
	if(rows.length<1) table.parentNode.removeChild(table);
};

//���ù�����ģʽ
TQE.toolbarMode={
	admin:{left:['paragraph','fontsize','|','bold','italic','underline','forecolor','backcolor','removeformat','|','justifyleft','justifycenter','orderedlist',['justifyright','unorderedlist','|','indent','outdent','|','superscript','subscript'],
		'||','createlink','insertimage','inserttable',['splitpage','inserthorizontalrule'],'|','insertmusic','insertflash','insertvideo',
		'||','cleanword','typeset'],
		right:['fullscreen','mode','|','help']},
	bbs:{left:['fontsize','|','bold','italic','underline','forecolor','backcolor','removeformat','|','justifyleft','justifycenter','orderedlist',['justifyright','unorderedlist','|','indent','outdent'],
		'||','createlink','insertimage','|','insertface','insertmusic','insertflash','insertvideo','|','typeset'],
		right:['fullscreen','mode','|','help']},
	user:{left:['fontsize','|','bold','italic','underline','forecolor','backcolor','removeformat','|','justifyleft','justifycenter','orderedlist',['justifyright','unorderedlist','|','indent','outdent','|','inserthorizontalrule'],
		'||','createlink','insertimage','insertvideo','|','typeset'],
		right:['fullscreen','mode','|','help']},
	mini:{left:['fontsize','|','bold','italic','underline','forecolor','backcolor','removeformat','|','align','|','createlink','insertimage'],
		right:['help']},
	message:{left:['insertface','|','fontsize','|','bold','italic','underline','|','forecolor','backcolor','|','removeformat'],
		right:['help']},
	full:{left:['paragraph','fontname','fontsize','|','bold','italic','underline','forecolor','backcolor','removeformat',
		'||','align','list','iodent',['superscript','subscript'],
		'||','createlink','unlink','insertimage','inserttable','inserthorizontalrule','splitpage','insertface','|','insertmusic','insertflash','insertvideo',
		'||','cleanword','typeset'],
		right:['fullscreen','mode','|','help']},
	crm:{left:['paragraph','fontname','fontsize','|','bold','italic','underline','forecolor','backcolor','removeformat',
		'||','align','superscript','subscript',
		'||','createlink','unlink','insertimage','inserttable','inserthorizontalrule','insertface',
		'||','typeset'],
		right:['fullscreen','mode']},
	normal:{left:['paragraph','fontsize','|','bold','italic','underline','forecolor','backcolor','removeformat',
		'||','justifyleft','justifycenter','orderedlist',['justifyright','unorderedlist','|','indent','outdent','|','superscript','subscript'],
		'||','createlink','insertimage','inserttable','|','insertvideo','|','typeset'],
		right:['fullscreen','mode','|','help']}
};
//���õİ�ť
TQE.toolbarBottons={
	paragraph:{
		text:'����', 
		click:function($,sender){
			var o=TQE.find($.objId+'Paragraph'),i;
			if(null==o){
				var s='',fs=$.config.paragraphs;
				for(i=0;i < fs.length; i++){
					s +='<div name="FormatBlock" class="eMenuItem" unselectable=on onselectstart="return false" paraValue="'+fs[i][1]+'">'+fs[i][0]+'</div>';
				}
				o = TQE.CE('DIV',$.objId+'Paragraph',s,true);
				o.className='ePopPanel';
				TQE._noSel(o);//,'DIV'
				TQE.listenTags(o,'DIV','click',function(){
					TQE.hidePop();
					var pv=this.getAttribute('paraValue');
					if(isIE)pv ='<'+pv+'>';
					$.exeCmd('FormatBlock', pv);
				});
				if(isIE6){ 
					var es = TQE._tags(o,'DIV');
					for(i=es.length-1;i>=0;i--){
						$._ieMouseStyle(es[i],'eMenuItem','eMenuItemUp','eMenuItem');
					}
				}
			}
			TQE.pop(o,sender);
		}
	},
	
	fontsize:{
		text:'���ִ�С', 
		click:function($,sender){
			var o=TQE.find($.objId+'FontSize');
			if(null==o){
				var s='',i,es, 
				fs=$.config.fontsizes;
				for(i=0;i < fs.length; i++){
					s +='<div name="FontSize" class="eMenuItem" unselectable=on onselectstart="return false" paraValue="'+fs[i][2]+'" style="font-size:'+fs[i][1]+';" >'+fs[i][0]+'</div>';
				}
				o = TQE.CE('DIV',$.objId+'FontSize',s,true);
				o.className='ePopPanel';
				TQE._noSel(o);//,'DIV'
				TQE.listenTags(o,'DIV','click',function(){TQE.hidePop();$.exeCmd('FontSize', this.getAttribute('paraValue'))});
				if(isIE6){ 
					es = TQE._tags(o,'DIV');
					for(i=es.length-1;i>=0;i--){
						$._ieMouseStyle(es[i],'eMenuItem','eMenuItemUp','eMenuItem');
					}
				}
			}
			TQE.pop(o,sender);
		}
	},
	fontname:{
		text:'����', 
		click:function($,sender){
			var o=TQE.find($.objId+'FontName');
			if(null==o){
				var es,i,s='',fs=$.config.fontnames;
				for(i=0;i < fs.length; i++){
					s +='<div name="FontName" class="eMenuItem" unselectable=on onselectstart="return false"  paraValue="'+fs[i][1]+'"  style="font-family:'+fs[i][1]+';" >'+fs[i][0]+'</div>';
				}
				o = TQE.CE('DIV',$.objId+'FontName',s,true);
				o.className='ePopPanel';
				TQE._noSel(o);//,'DIV'
				TQE.listenTags(o,'DIV','click',function(){TQE.hidePop();$.exeCmd('FontName', this.getAttribute('paraValue'))});
				if(isIE6){ 
					es = TQE._tags(o,'DIV');
					for(i=es.length-1;i>=0;i--){
						$._ieMouseStyle(es[i],'eMenuItem','eMenuItemUp','eMenuItem');
					}
				}
			}
			TQE.pop(o,sender);
		}
	},
	forecolor:{
		text:'�ı���ɫ', 
		click:function($,sender){
			$._showColorSelectUI(sender,'forecolor'); 
		}
	},
	backcolor:{
		text:'����ɫ', 
		click:function($,sender){
			$._showColorSelectUI(sender, isIE ? 'backcolor':'hilitecolor');
		}
	},
	bold:{
		text:'����', 
		click:null
	},
	italic:{
		text:'б��', 
		click:null
	},
	underline:{
		text:'�»���', 
		click:null
	},
	strikethrough:{
		text:'ɾ����', 
		click:null
	},
	removeformat:{
		text:'������ʽ', 
		click:function($,sender){
			TQE.hidePop();
			$.exeCmd('removeformat'); 
			if(isGecko || isOpera) return;//ff,opera�����������������ʽ, �������������Ҫ�Ӵ��봦��word�Ĵ���
			var rang=$._getRange(),
				clear=function(s){return s.replace(/mso\-[^;]+?;/gi,'').replace(/FONT\-FAMILY[^;]+?;/gi,'').replace(/\s+/g,' ').replace(/MARGIN\: 0cm 0cm 0pt;/gi,'').replace(/tab\-stops\:[^;]+?;/gi,'');},
				htm='',
				str=rang.htmlText,
				nodes,
				i;
			if(str){//ie6,7,8
				$.insertHtml(clear(s));
			}else{//ie9,chrome
				nodes=rang.cloneContents().childNodes;
				for(i=0;i<nodes.length;i++){
					if('#comment'!=nodes[i].nodeName){
						str =  nodes[i].outerHTML || nodes[i].nodeValue ;
						if(str) htm += str;
					}
				}
				$.insertHtml(clear(htm));
			}
		}
	},
	submenu:{
		text:'', 
		click:function($,sender){
			$._popAsMenu(sender,sender.getAttribute('submenu').split(','));
		}
	},
	//ʹ��submenu������˹���
	//moreformat:{
	//	text:'', 
	//	click:function($,sender){
	//		$._popAsMenu(sender,['orderedlist','unorderedlist','|','indent','outdent','|','superscript','subscript']);
	//	}
	//},
	subscript:{
		text:'�±�', 
		click:null
	},
	superscript:{
		text:'�ϱ�', 
		click:null
	},
	align:{
		text:'����', 
		click:function($,sender){
			$._popAsMenu(sender,['justifyleft','justifycenter','justifyright']);
		}
	},
	justifyleft:{
		text:'�����', 
		click:null
	},
	justifycenter:{
		text:'���ж���', 
		click:null
	},
	justifyright:{
		text:'�Ҷ���', 
		click:null
	},
	list:{
		text:'�б�', 
		click:function($,sender){ 
			$._popAsMenu(sender,['orderedlist','unorderedlist']); 
		}
	},
	unorderedlist:{
		text:'��Ŀ�б�', 
		click:function($,sender){
			TQE.hidePop();
			$.exeCmd('InsertUnorderedList');
		}
	},
	orderedlist:{
		text:'��Ŀ�б�', 
		click:function($,sender){
			TQE.hidePop();
			$.exeCmd('InsertOrderedList');
		}
	},
	iodent:{
		text:'��������', 
		click:function($,sender){
			$._popAsMenu(sender,['indent','outdent'])
		}
	},
	outdent:{
		text:'��������', 
		click:null
	},
	indent:{
		text:'��������', 
		click:null
	},
	inserthorizontalrule:{
		text:'����ˮƽ��', 
		click:null
	},
	createlink:{
		text:'�������', 
		click:function($,sender){
			var o=TQE.find($.objId+'CreateLink');
			if(null==o){
				//var s='��ַ:<input type="text" id="url" value="" size=45 /><br />��:<select id="target"><option value="" >Ĭ��</option><option value="_blank">�´���</option><option value="_self">��ǰ����</option><option value="_parent">������</option><option value="_top">���㴰��</option></select>';
				var okClick= function(){
					TQE.hidePop();
					var es=this.form.elements,
						href=es['url'].value,
						target=es['target'].value,
						title=es['title'].value,
						titleTxt='';
					if('http://'==href || ''==href){
						$.exeCmd('Unlink');
						return false;
					}
					var r= $._getRange();
					if(isWebkit && r.collapsed || !$.exeCmd('CreateLink', href)){
						if(target!=='')  target = ' target="'+target+'" ';
						if(title!=='')  titleTxt = ' title="'+title+'" ';
						$.insertHtml('<a href="'+href+'" '+titleTxt+target+'>'+(title?title:href)+'</a>');
						return false;
					}
					var n=$.focusNode();
					do{
						if(n.tagName=='A'){
							if(''==target) n.removeAttribute('target') ;
							else n.setAttribute('target',target);
							if(''==title) n.removeAttribute('title');
							else n.setAttribute('title',title);
							break;
						}
					}while(n = n.parentNode);
					return false;
				};
				o = $._popFormUI($.objId+'CreateLink',resStrCreateLink, resStrInsertLinkUI, okClick,null, $.config.linkUploadUrl,$.config.linkFileTypes,$.config.linkFileSize,'link');
				TQE.drag(o);
			}
			if(!TQE.pop(o,sender)) return;

			var f = TQE._tags(o,'FORM')[0],
				es = f.elements,
				n = $.focusNode();
			f.reset();
			do{
				if(n.tagName=='A'){
					es['url'].value=n.getAttribute('href');
					es['title'].value=n.getAttribute('title');
					$.__selectOption(es['target'],n.getAttribute('target'));
					break;
				}
			}while(n = n.parentNode);
			es['url'].focus();
		}
	},
	unlink:{
		text:'ɾ������', 
		click:null
	},
	inserttable:{
		text:'������', 
		click:function($,sender){
			var o=TQE.find($.objId+'InsertTable');
			if(null==o){
				var s=resStrTableContextUI + '<div><table border=1 cellspacing=0 cellpadding=0 bgcolor="#FFFFFF"  bordercolor="#CCCCCC"  style="border:none;border-collapse:collapse;"><tbody>',i=0,j=0;
				for(i=0;i<5;i++){
					s+='<tr>';
					for(j=0;j<8;j++){
						s+='<td width="20" height="20"></td>';
					}
					s+='</tr>';
				}
				s+='</tbody></table><div style="padding:2px 5px;">1 &times 1</div></div>';
				o = TQE.CE('DIV',$.objId+'InsertTable', s ,true);
				o.className='ePopPanel';
				TQE._noSel(o);//,'TABLE,TR,TD,DIV');
				TQE.listenTags(o.lastChild,'TD','click',function(){
						TQE.hidePop();
						var i,j,s='<table style="border-collapse:collapse;" width=96% align=center border=1 cellspace=0 cellpadding=3 ><tbody>';
						for(i=this.parentNode.rowIndex;i>=0;i--){
							s+='<tr>';
							for(j=this.cellIndex;j>=0;j--){
								s+= isIE ? '<td valign="top"></td>':'<td valign="top"><br></td>';
							}
							s+='</tr>';
						}
						s+='</tbody></table>';
						//s = '<p>'+s+'</p>';
						//$.exeCmd('InsertParagraph');
						//$.exeCmd('FormatBlock','p');
						$.insertHtml(s);
					}
				);
				TQE.listenTags(o.lastChild,'TD','mouseover',function(){
					var i,j,ci=this.cellIndex, row=this.parentNode,ri=row.rowIndex,table=row.parentNode.parentNode, rc=table.rows.length,cc=row.cells.length;
					for(i=0;i<rc;i++){
						row=table.rows[i];
						for(j=0;j<cc;j++)
							row.cells[j].bgColor= (i<=ri && j<=ci)?'#3399ff':'';
					}
					table.parentNode.lastChild.innerHTML= (ri+1) +' &times; '+(ci+1);
				});
			}
			var table, td, 
				selectedCells=$.selectedCells(),
				propertyPanel=o.firstChild,
				miMerge=propertyPanel.firstChild,
				miSplit=miMerge.nextSibling,
				insertPanel=o.lastChild,

				row=o.lastChild.firstChild.firstChild.firstChild,
				cc=row.cells.length,
				showObj=function(obj,b){obj.style.display=b?'':'none';};
			while(row){
				for(var i=0;i<cc;i++) row.cells[i].bgColor='';
				row=row.nextSibling;
			}
			insertPanel.lastChild.innerHTML='1 &times; 1';
			//--
			if(selectedCells.length){//ѡ�и�
				showObj(insertPanel, 0);
				td=selectedCells[0];
			}else{
				showObj(insertPanel, 1);
				td=$.focusNode();
				while(td && td.nodeName !=='TD'){
					td = td.parentNode	 
				}
			}
			if(td){
				table=td.parentNode.parentNode.parentNode;
				showObj(propertyPanel, 1);
				showObj(miMerge, $.canMergeSelectedCells());
				showObj(miSplit, selectedCells.length<2 && (td.rowSpan>1 || td.colSpan>1));
				showObj(miSplit.nextSibling, miMerge.style.display!='none' || miSplit.style.display!='none');

			}else{
				showObj(propertyPanel, 0);
			}
			showObj(insertPanel.previousSibling, propertyPanel.style.display!='none' && insertPanel.style.display!='none');
			
			var objs=o.firstChild.getElementsByTagName('A'), 
				alignTD=function(align,vAlign){
					if(selectedCells.length>0){
						for(var i=selectedCells.length-1;i>=0;i--){
							selectedCells[i].align=align;
							selectedCells[i].vAlign=vAlign;
						}
					}else if(td){
						td.align=align;
						td.vAlign=vAlign;
					}
				};
			for(i=objs.length-1;i>=0;i--){
				objs[i].href='#';
				objs[i].onclick=function(){
					switch(this.id){
						case 'eTableAlignDefault':table.align=''; break;
						case 'eTableAlignLeft':table.align='left'; break;
						case 'eTableAlignCenter':table.align='center'; break;
						case 'eTableAlignRight':table.align='right'; break;
						case 'eCellAlignLT':alignTD('left','top');break;
						case 'eCellAlignCT':alignTD('center','top');break;
						case 'eCellAlignRT':alignTD('right','top');break;
						case 'eCellAlignLC':alignTD('left','middle');break;
						case 'eCellAlignCC':alignTD('center','middle');break;
						case 'eCellAlignRC':alignTD('right','middle');break;
						case 'eCellAlignLB':alignTD('left','bottom');break;
						case 'eCellAlignCB':alignTD('center','bottom');break;
						case 'eCellAlignRB':alignTD('right','bottom');break;
					}
					$.focus();
					return false;
				};
			}
			objs=o.firstChild.getElementsByTagName('INPUT');
			for(i=objs.length-1;i>=0;i--){
				if(table) objs[i].value= undefined==table[objs[i].id] ? '' : table[objs[i].id];
				objs[i].onblur=function(){
					if(''==this.value ) table.removeAttribute(this.id);
					else table[this.id]=this.value;
					table.style.borderCollapse = table.cellSpacing<1 && table.border==1 ? 'collapse':'';
				};
			}

			miMerge.onclick=function(){TQE.hidePop();$.mergeSelectedCells();$.focus(); return false;}
			miSplit.onclick=function(){TQE.hidePop();$.splitCell(td);$.focus();return false;}

			//--
			if(!TQE.pop(o,sender)) return;
		}
	},
	insertimage:{
		text:'����ͼƬ', 
		click:function($,sender){
			var o=TQE.find($.objId+'InsertImage');
			if(null==o){
				//var s='ͼƬ��ַ:<input type="text" id="url" value="" size=40 /><br />�滻�ı�:<input name="alt" type="text" id="alt" size="20" /><br />�ߴ�:<input name="width" type="text" id="width" size="4" /> &times; <input name="height" type="text" id="height" size="4" /> px &nbsp;<br />����:<select name="align" id="align"><option value="" >Ĭ��</option><option value="top">����</option><option value="textTop">�ı��Ϸ�</option><option value="middle">����</option><option value="absMiddle">���Ծ���</option><option value="baseline">����</option><option value="bottom">�ײ�</option><option value="absBottom">���Եײ�</option><option value="left">��</option><option value="right">��</option></select> &nbsp; &nbsp; �߿�:<input name="border" type="text" id="border" size="5" />';
				var okClick=function(){
					TQE.hidePop();
					var es=this.form.elements;
					$.insertImage(es['url'].value,es['width'].value,es['height'].value,es['border'].value, es['alt'].value,es['align'].value);
					return false;
					};
				o = $._popFormUI($.objId+'InsertImage',resStrInsertImage, resStrInsertImageUI, okClick ,['width','height','border'], $.config.imageUploadUrl,$.config.imageFileTypes, $.config.imageFileSize, 'image');
				TQE.drag(o);
			}
			if(!TQE.pop(o,sender)) return;
			var f = TQE._tags(o,'FORM')[0],
				es = f.elements,
				r=$._getRange(),
				n=r.commonAncestorContainer;
			f.reset();
			if(n){ //��׼dom
				if(!r.collapsed && r.startContainer == r.endContainer && r.startOffset - r.endOffset < 2 && r.startContainer.hasChildNodes()){
					n=r.startContainer.childNodes[r.startOffset];
				}
				if(n.tagName!='IMG' && isIE9 && ''===r.toString()){
					if(undefined===r.startContainer.length && r.startContainer.firstChild===r.endContainer.previousSibling)n=r.startContainer.firstChild;
					else if(r.endContainer.lastChild===r.startContainer.nextSibling)n=r.endContainer.lastChild;
					else if(r.startContainer.length == r.startOffset && r.endOffset==0 && r.startContainer.nextSibling==r.endContainer.previousSibling){
						n=r.startContainer.nextSibling;
					}
				} 
				if(!n || n.tagName!='IMG') n=null; 
			}else if(r.length && r.item(0).tagName=='IMG'){	//isIE, ie6,7,8
				n=r.item(0);
			}
			if(n && 'IMG'==n.nodeName){
				 es['url'].value=n.getAttribute('src');
				 es['width'].value=n.width;
				 es['height'].value=n.height;
				 es['border'].value=n.border;
				 es['alt'].value=n.alt;
				 $.__selectOption(es['align'],n.align);
			}
			es['url'].focus();
		}
	},
	insertface:{
		text:'����', 
		click:function($,sender){
			var o=TQE.find($.objId+'InsertFace');
			if(null==o){
				var i,es,s='<div class="eFace">';
				for( i=1;i<=$.config.faceCount;i++){
					s+='<img src="'+TQE.url+'face/'+i+'.gif" emot="'+i+'" />';
				}
				s+='</div>';
				o = TQE.CE('DIV',$.objId+'InsertFace', s ,true);
				o.className='ePopPanel';
				TQE._noSel(o);//,'IMG,DIV'); 
				TQE.listenTags(o,'IMG','click',function(){TQE.hidePop();$.insertHtml('<img align=absMiddle src="'+this.getAttribute('src')+'" emot="'+this.getAttribute('emot')+'" />');});
				if(isIE6){ 
					es = TQE._tags(o,'IMG');
					for(i=es.length-1;i>=0;i--){
						$._ieMouseStyle(es[i],'','imgHover','');
					}
				}
			}
			TQE.pop(o,sender);
		}
	},
	insertflash:{
		text:'����Flash', 
		click:function($,sender){
			var o=TQE.find($.objId+'InsertFlash'),form;
			if(null==o){
				//var s='Flash��ַ:<input type="text" id="url" value="" size=40 /><br />�ߴ�:<input name="width" type="text" id="width" size="4" value="200" />&times;<input name="height" type="text" id="height" size="4" value="200" /> px';

				var okClick= function(){
					TQE.hidePop();
					var es=this.form.elements;
					$.insertFlash(es['url'].value,es['width'].value,es['height'].value);
					return false;
				};
				o = $._popFormUI($.objId+'InsertFlash',resStrInsertFlash, resStrInsertFlashUI, okClick,['width','height'], $.config.flashUploadUrl,$.config.flashFileTypes,$.config.flashFileSize, 'flash');
				TQE.drag(o);
			}
			if(!TQE.pop(o,sender)) return;
			form=TQE._tags(o,'FORM')[0];
			form.reset();
			form.elements['url'].focus();
		}
	},
	insertmusic:{
		text:'������Ƶ', 
		click:function($,sender){
			var o=TQE.find($.objId+'InsertMusic'),form;
			if(null==o){
				//var s='��Ƶ��ַ:<input type="text" id="url" value="" size=40 /><br />�Զ�����:<select id="auto_start" name="auto_start"><option value="0">�ֶ���ʼ</option><option value="1">�Զ���ʼ</option></select>';
				var okClick= function(){
					TQE.hidePop();
					var es=this.form.elements;
					$.insertMusic(es['url'].value,es['auto_start'].value);
					return false;
				};
				o = $._popFormUI($.objId+'InsertMusic',resStrInsertMusic, resStrInsertMusicUI, okClick,null, $.config.musicUploadUrl,$.config.musicFileTypes,$.config.musicFileSize, 'music');
				TQE.drag(o);
			}
			if(!TQE.pop(o,sender)) return ;
			form=TQE._tags(o,'FORM')[0];
			form.reset();
			form.elements['url'].focus();
		}
	},
	insertvideo:{
		text:'������Ƶ', 
		click:function($,sender){
			var o=TQE.find($.objId+'InsertVideo'),form;
			if(null==o){
				//var s='��Ƶ��ַ:<input type="text" id="url" value="" size=40 /><br />�ߴ�:<input name="width" type="text" id="width" size="4"  value="320"/>&times;<input name="height" type="text" id="height" size="4" value="240" /> px<br />�Զ�����:<select id="auto_start" name="auto_start"><option value="0">�ֶ���ʼ</option><option value="1">�Զ���ʼ</option></select><br />��Ƶ����:<select id="video_type" name="video_type"><option value="auto">�Զ�ʶ��</option><option value="flv">FLV</option><option value="rm">RMVB</option><option value="wm">������Ƶ</option></select>';
				var okClick= function(){
					TQE.hidePop();
					var es=this.form.elements;
					$.insertVideo(es['url'].value,es['auto_start'].value,es['width'].value,es['height'].value,es['video_type'].value);
					return false;
				};
				o = $._popFormUI($.objId+'InsertVideo',resStrInsertVideo, resStrInsertVideoUI, okClick,['width','height'], $.config.videoUploadUrl,$.config.videoFileTypes,$.config.videoFileSize, 'video');
				TQE.drag(o);
			}
			if(!TQE.pop(o,sender)) return;
			form=TQE._tags(o,'FORM')[0];
			form.reset();
			form.elements['url'].focus();
		}
	},
	cleanword:{
		text:'����Word����', 
		click:function($,sender){
			TQE.hidePop();
			$.cleanWord(true);
		}
	},
	typeset:{
		text:'һ���Ű�', 
		click:function($,sender){
			TQE.hidePop();
			$.typeset();
		}
	},
	mode:{
		text:'�鿴Դ��', 
		click:function($,sender){
			TQE.hidePop(); 
			$.changeMode();
		}
	},
	splitpage:{
		text:'�����ҳ��', 
		click:function($,sender){ 
			TQE.hidePop(); 
			//$.insertHtml($.config.tplSplitPage,'after');
			$.insertHtml($.config.tplSplitPage);
		}
	},
	group:{
		text:'', 
		click:null
	},
	separator:{
		text:'', 
		click:null
	},
	fullscreen:{
		text:'ȫ��', 
		click:function($,sender){
			TQE.hidePop();
			$.fullScreen();
		}
	},
	help:{
		text:'����', 
		click:function($,sender){
			var o=TQE.find($.objId+'TQHelpPanel');
			if(null==o){
				o = TQE.CE('DIV',$.objId+'TQHelpPanel', '<form class="ePopForm" style="line-height:20px;">TQEditor<br/>'+resStrVersion+': '+version+'<br/>'+resStrAuthor+': ������(TengQiu Li)<br/>'+resStrSite+': <a href="http://www.e512.net" target="_blank">e512.net</a><br />'+resStrMail+': <a href="mailto:litqqs@163.com">litqqs@163.com</a><br /><a href="tencent://message/?uin=7775038&Site=e512.net&Menu=no" style="clear:both;"><span class="eButton" style="background-position: -1026px 0;width:20px;" ></span> 7775038</a></form>' ,true);
				o.className='ePopPanel';
			}
			TQE.pop(o,sender);
		}
	}
};
//����ʱ��������
//TQE.rto={};
//explode names;
window.tqEditor=window.tqeditor=window.TQEditor=window.TQE=TQE;
window.isIE=isIE; 
window.isIE6=isIE6; 
window.isIE9=isIE9; 
window.isWebkit=isWebkit; 
window.isOpera=isOpera; 
window.isGecko=isGecko; 
})(window);