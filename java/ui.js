//image src replace
jQuery.fn.imgSwap = function(src_in,src_out){
	if(typeof src_out != "string") src_out="_on";
	$(this).each(function(){
		var that=$(this);
		var imgSrc=that.attr("src");
		if(imgSrc==undefined) return;
		var imgType=imgSrc.match(/.gif$|.jpg$|.png$/);
		that.attr("src",imgSrc.replace(src_in+imgType,src_out+imgType));
	});
	return this;
}

function imgRollover(selector,options){
	var opt = {
		srcOn  : "_on"
		, srcOff : ""
		, hold   : null
	}
	opt = $.extend(opt, options || {});
	function setBind(el){
		if(el.hasClass(opt.hold)) return;
		el.hover(function(){
			$(this).imgSwap(opt.srcOff,opt.srcOn);
		},function(){
			$(this).imgSwap(opt.srcOn,opt.srcOff);
		});
	}
	for(var s in selector){
		o = $(selector[s]);
		if(o.is("[src]")){
			o.each(function(){
				setBind($(this));
			});
		}else{
			o.find("[src]").each(function(){
				setBind($(this));
			});
		}
	}
}

//롤링 배너
jQuery.fn.slideImages = function(options){
	$(this).each(function(){
		var opt = {
			wrap           : $(this)
			,type          : 1
			,ul            : "ul"
			,btn_prev      : ".prev" //이전버튼
			,btn_pause     : ".pause" //일시정지
			,btn_next      : ".next" //다음버튼
			,navi          : ".navi" //네비
			,naviAuto      : true //배너 수 만큼 네비 자동 복사 true | false
			,resizingCheck : false //창사이즈 변경시 이미지 크기 재설정 true | false
			,hoverClass    : "over"
			,navSwapOn     : "on" //네비 이미지 src replace on (ico_on.gif)
			,navSwapOff    : "off" //네비 이미지 src replace off (ico_off.gif)
			,eventHandler  : "mouseenter focus" //네비 이벤트 핸들러
			,speed         : 300 //슬라이딩 속도
			,rollTimer     : 4000 //롤링 대기 시간 1/1000초
			,navAlt        : "배너" //네비가 이미지일때 alt값에 자동으로 this.navAlt값 + 숫자값으로 들어감
			,viewNum       : 1 //보여지는 이미지 수
			,fixResizing   : null //resizingCheck : true 일때 리사이징 기준이 되는 엘리먼트
			,fnBefore      : null //초기, 리사이징 될때 함수 정의

			// * type 설명
			// 1 - 왼쪽 오른쪽으로 슬라이딩
			// 2 -
			// 3 - 이미지를 fade inㆍout하면서 전환
			// 4 - 상하로 슬라이딩
		}
		opt = $.extend(opt, options || {});

		(new slideImages()).init(function(){
			this.el.wrap       = opt.wrap;
			this.el.ul         = opt.ul;
			this.el.btn_prev   = opt.btn_prev;
			this.el.btn_pause  = opt.btn_pause;
			this.el.btn_next   = opt.btn_next;
			this.el.navi       = opt.navi;
			this.type          = opt.type;
			this.naviAuto      = opt.naviAuto;
			this.resizingCheck = opt.resizingCheck;
			this.hoverClass    = opt.hoverClass;
			this.navSwapOn     = opt.navSwapOn;
			this.navSwapOff    = opt.navSwapOff;
			this.eventHandler  = opt.eventHandler;
			this.speed         = opt.speed;
			this.rollTimer     = opt.rollTimer;
			this.navAlt        = opt.navAlt;
			this.viewNum       = opt.viewNum;
			this.fnBefore      = opt.fnBefore;
			this.fixResizing   = opt.fixResizing;
		});
	});
	return this;
}

function slideImages(){
	this.el = {}
	this.el.li        = ">*";
	this.el.a         = "a";
	this.target       = "0";
	this.curr         = 0;
	this.pause       = false;
}
$.extend(slideImages.prototype,{
	bind : function(obj,evt,fn){
		var o = this;
		for(var s in obj){
			obj[s].bind(evt,fn);
		}
	}
	,firstSetting : function(){
		var o = this;
		function init(){
			if(typeof o.fnBefore === "function"){
				o.fnBefore();
			}
			o.size = o.el.li.outerWidth();
			o.vsize = o.el.li.outerHeight();
			o.totalSize = o.size*o.total+100;
			o.totalvSize = o.vsize*o.total+100;

			if(o.type === 1){
				o.el.ul.css({overflow:"hidden",position:"relative",width:o.totalSize,height:o.vsize});
				o.el.a.css({width:o.size,height:o.vsize});
				o.el.li.css({position:"absolute"});
				o.el.li.each(function(i){
					$(this).css("left",o.size*i);
				});
			}else if(o.type === 2){

			}else if(o.type === 3){
				o.el.ul.css({position:"relative",zIndex:0});
				o.el.li.css({display:"none",position:"absolute",zIndex:1,opacity:0});
				o.el.li.eq(0).css({display:"block",opacity:1});
			}else if(o.type === 4){
				o.el.ul.css({overflow:"hidden",position:"relative",width:o.size,height:o.totalvSize});
				o.el.a.css({width:o.size,height:o.vsize});
				o.el.li.css({position:"absolute"});
				o.el.li.each(function(i){
					$(this).css("top",o.vsize*i);
				});
			}
		}
		if(o.resizingCheck){
			if(o.fixResizing !== null){
				o.fixResizing = $(o.fixResizing);
			}else{
				o.fixResizing = o.el.wrap;
			}
			o.el.li.width($(o.fixResizing).outerWidth());
			$(window).resize(function(){
				o.el.li.width($(o.fixResizing).outerWidth());
				init();
				o.move();
			});
		}
		init();
	}
	,attachNavi : function(){
		var o = this;
		if(!o.el.navi.length) return;
		if(o.naviAuto){
			var html=o.el.navi.html();
			for(var i=1; i<o.total; i++){
				o.el.navi.append(html);
			}
		}
		if(o.el.navi.find("a,button").length){
			o.el.btn_navi=o.el.navi.find("a,button");
		}else{
			o.el.btn_navi=o.el.navi.find(">*");
		}
		o.el.navi.find("a").each(function(i){
			var aname = "____slideImages___"+$("*[id*='____slideImages___']").length+"0"+i;
			$(this).attr("href","#"+aname);
			o.el.a.eq(i).attr("id",aname);
		});
		o.el.navi.find("img").each(function(i){
			$(this).attr("alt",o.navAlt+" "+(i+1));
		});
	}
	,move : function(){
		var o = this;
		if(o.type === 1){
			o.target=-o.curr*o.size;
			o.el.ul.stop().animate({marginLeft:o.target+"px"},o.speed);
			o.naviEffect();
		}else if(o.type === 2){
			o.target=-o.curr*o.size;
			o.el.ul.stop().animate({marginLeft:o.target+"px"},o.speed);
			o.naviEffect();
		}else if(o.type === 3){
			var obj,obj_prev;
			obj = o.el.li.eq(o.curr);
			obj_prev = o.el.li.filter(":visible");
			obj.css("zIndex",1);
			obj_prev.css("zIndex",0);
			$(obj_prev, obj_prev.find("img")).stop().animate({opacity:0},o.speed,function(){
				$(this).css("display","none");
			});
			$(obj, obj.find("img")).stop().css("display","block").animate({opacity:1},o.speed);
			o.naviEffect();
		}else if(o.type === 4){
			o.target=-o.curr*o.vsize;
			o.el.ul.stop().animate({marginTop:o.target+"px"},o.speed);
			o.naviEffect();
		}
	}
	,naviEffect : function(){
		if(!this.el.navi.length) return;
		var o = this;
		if(o.el.btn_navi.find("img").length && o.naviAuto){
			o.el.btn_navi.filter("."+o.hoverClass).find("img").imgSwap(o.navSwapOn,o.navSwapOff);
			o.el.btn_navi.eq(o.curr).find("img").imgSwap(o.navSwapOff,o.navSwapOn);
		}
		o.el.btn_navi.removeClass(o.hoverClass);
		o.el.btn_navi.eq(o.curr).addClass(o.hoverClass);
	}
	,rolling : function(time){
		var o=this;
		var timer;
		function action(){
			clearInterval(timer);
			if(o.pause) return;
			timer=setInterval(function(){
				if(o.curr+o.viewNum>=o.total) o.curr=0
				o.curr=(o.curr+1)%(o.el.li.length);
				o.move();
			},time);
		}
		function clearTimer(){
			clearInterval(timer);
			clearInterval(timer);
		}
		o.el.wrap.mouseover(function(){
			clearTimer();
		});
		o.bind([o.el.wrap, o.el.a, o.el.btn_prev, o.el.btn_pause, o.el.btn_next],"focus",function(){
			clearTimer();
		})
		o.el.wrap.mouseleave(function(){
			action();
		});
		o.bind([o.el.wrap, o.el.a, o.el.btn_prev, o.el.btn_pause, o.el.btn_next],"blur",function(){
			action();
		})
		if(o.el.btn_navi){
			o.el.btn_navi.focus(function(){
				clearTimer();
			});
			o.el.btn_navi.blur(function(){
				action();
			});
		}
		action();
	}
	,imgReady : function(fn){
		var o,img,k;
		o = this;
		img = o.el.li.find("img:first")[0];
		k = 0;
		function chk(){
			if(++k>5){
				img = o.el.li.eq(1).find("img")[0];
			}
			if(img.complete){
				fn();
				return;
			}
			setTimeout(chk,100);
		}
		chk();
	}
	,init : function(fn){
		var o = this;
		if(typeof fn==="function"){
			fn.apply(o);
		}

		o.el.ul = o.el.wrap.find(o.el.ul);
		o.el.li = o.el.ul.find(o.el.li);
		o.el.a = o.el.li.find(o.el.a);
		o.el.btn_prev = o.el.wrap.find(o.el.btn_prev);
		o.el.btn_pause = o.el.wrap.find(o.el.btn_pause);
		o.el.btn_next = o.el.wrap.find(o.el.btn_next);
		o.el.navi = o.el.wrap.find(o.el.navi);
		o.total = o.el.li.length;

		o.attachNavi();
		o.imgReady(function(){
			o.firstSetting();
			o.move();

			o.el.btn_prev.bind("click",function(){
				if(o.curr<=0) return false;
				o.curr--;
				o.move();
				return false;
			});
			o.el.btn_next.bind("click",function(){
				if(o.curr+o.viewNum>=o.total) return false;
				o.curr++;
				o.move();
				return false;
			});
			o.el.btn_pause.bind("click",function(){
				o.pause = !o.pause;
				o.el.ul.stop();
				return false;
			});
			if(o.el.btn_navi){
				o.el.btn_navi.each(function(i){
					$(this).bind(o.eventHandler,function(){
						o.curr=i;
						o.move();
					});
				});
			}
			o.el.a.each(function(i){
				$(this).bind("focus",function(){
					o.el.ul.stop();
					o.curr=i;
					o.move();
					o.el.wrap.scrollLeft(0);
				});
			});

			o.rolling(o.rollTimer);
		});
		return this;
	}
});

//탭
jQuery.fn.tabUI = function(options){
	var opt = {
		tab : ".tabUI_tab"
		,con : ".tabUI_con"
		,tabSrcOn : "_ov"
		,tabSrcOff : ""
		,tabHover : "on"
		,eventHandler : "mouseenter focus"
		,imgHover : true
		,fn : ""
	};
	var opt = $.extend(opt, options || {});
	function init(wrap){
		var el;
		el = {
			tab : wrap.find(opt.tab)
			,con : wrap.find(opt.con)
		}
		
		el.tab.each(function(i){
			$(this).bind(opt.eventHandler,function(){
				var o = $(this);
				if(opt.fn != ""){
					opt.fn(i);
				}
				if(opt.imgHover){
					el.tab.filter("."+opt.tabHover).find("img").imgSwap(opt.tabSrcOn,opt.tabSrcOff);
					if(o.find("img").length){
						$(this).find("img").imgSwap(opt.tabSrcOff,opt.tabSrcOn);
					}
				}
				el.tab.removeClass(opt.tabHover);
				o.addClass(opt.tabHover);
				el.con.hide();
				el.con.eq(i).show();
			});
		});
		
		el.tab.eq(0).trigger(opt.eventHandler.split(" ")[0]);
		
	}
	$(this).each(function(){
		init($(this));
	});
	return this;
}

jQuery.fn.GNB = function(){
	var o,d1,d1A,d1img,d2,d2A,d1A_on,d2A_on,d2_on,subMenuBg,timer,subefc;
	o = $(this);
	if(!o.length) return;
	d1_class = ".dep1";
	d2_class = ".dep2";
	srcOff = "";
	srcOn = "_on";
	d1 = o.find(d1_class);
	d1A = d1.find(">a");
	d1img = d1A.find("img");
	d2 = o.find(d2_class);
	d2A = d2.find("a");
	subMenuBg=o.find(".subMenuBg");
	d1A_on=$("#_+#_"),d2A_on=$("#_+#_"),d2_on=$("#_+#_");
	
	//초기화
	function init(){
		activePage();
		reset();
	}
	
	//활성화
	function activePage(){
		if(typeof getPageCode != "function") return;
		if(typeof getPageCode() != "string") return;
		var pageCode=getPageCode().split(" ");
		var dep1Code=parseInt(pageCode[0]);
		var dep2Code=parseInt(pageCode[1]);
		if(dep1Code<0) return;
		d1A_on = d1A.eq(dep1Code);
		if(dep2Code<0){
			d2_on = d1A_on.siblings(d2_class);
		}else{
			d2A_on = d1A_on.next().find("a").eq(dep2Code);
			d2_on = d2A_on.parents(d2_class).eq(0);
		}
	}
	
	function overEffect(){
		subMenuBg.show();
	}
	
	function outEffect(){
		subMenuBg.hide();
	}
	
	//처음 상태
	function reset(){
		timer = setTimeout(function(){
			menuOn(d1A_on,d2_on);
			subOn(d2A_on);
		},50);
	}
	
	//1depth,2depth off
	function menuOff(){
		var curr_dep1 = d1A.filter(".on");
		curr_dep1.removeClass("on");
		d1img.each(function(){
			$(this).imgSwap(srcOn,srcOff);
		});
		d2.hide();
	}
	
	//1depth on
	function menuOn(el_dep1,el_dep2){
		if(!el_dep1.hasClass("on")){
			subefc = true;
		}else{
			subefc = false;
		}
		menuOff();
		el_dep1.addClass("on");
		el_dep1.find("img").imgSwap(srcOff,srcOn);
		if(el_dep1.attr("class") == "on" && el_dep2.length){overEffect();}
		else{outEffect();}
		el_dep2.show();
		if(subefc === true){
			el_dep2.children().css({bottom:"-20px"}).stop().animate({bottom:0},200);
		}
	}
	
	//2depth on
	function subOn(el){
		d2A.filter(".on").find("img").imgSwap(srcOn,srcOff);
		el.find("img").imgSwap(srcOff,srcOn);
		d2A.removeClass("on");
		el.addClass("on");
	}
	
	//1depth mouseover
	d1A.each(function(i){
		$(this).bind("mouseover focus", function(){
			clearReset();
			var el_dep1 = $(this);
			var el_dep2 = $(this).parent().find(d2_class);
			menuOn(el_dep1,el_dep2);
		});
	});
	
	//2depth mouseover
	d2A.bind("mouseover focus", function(){
		clearReset();
		subOn($(this));
	});
	
	d1A.filter(":first").bind("blur", function(){
		reset();
	});
	d2A.filter(":last").bind("blur", function(){
		reset();
	});
	
	//처음 상태로 가기 취소
	function clearReset(){
		clearTimeout(timer);
	}
	
	//완전히 벗어나면 처음 상태로 가기
	o.bind("mouseleave",function(){
		clearReset();
		reset();
	});
	
	init();
}

jQuery.fn.SNB = function(){
	var d1_class, d2_class;
	var o,d1,d1A,d1img,d2,d2A,d1A_on,d2A_on,d2_on,timer;
	o = $(this);
	if(!o.length) return;
	d1_class = ".dep1";
	d2_class = ".dep2";
	srcOff = "";
	srcOn = "_on";
	d1 = o.find(d1_class);
	d1A =d1.find(">a");
	d1img = d1A.find("img");
	d2 = o.find(d2_class);
	d2A = d2.find("a");
	d1A_on=$("#_+#_"),d2A_on=$("#_+#_"),d2_on=$("#_+#_");

	function imgReady(fn){
		var img,k=0,total=0;
		img = o.find("img");
		img.each(function(){
			total++;
			var that = $(this);
			function chk(){
				if(that[0].complete){
					k++;
					if(total===k){
						fn();
					}
					return;
				}
				setTimeout(chk,100);
			}
			chk();
		});
	}
	
	//초기화
	function init(){
		imgReady(function(){
			setHeight();
		});
	}
	
	function setHeight(){
		d2.css("visibility","hidden").each(function(){
			var h;
			h = $(this).outerHeight();
			$(this).attr("hgt",h).css("visibility","visible");
		});
		activePage();
		reset();
		
	}

	//활성화
	function activePage(){
		if(typeof getPageCode != "function") return;
		if(typeof getPageCode() != "string") return;
		var pageCode=getPageCode().split(" ");
		var dep1Code=parseInt(pageCode[1]);
		var dep2Code=parseInt(pageCode[2]);
		if(dep1Code<0) return;
		d1A_on = d1A.eq(dep1Code);
		if(dep2Code<0) return;
		d2A_on = d1A_on.next().find("a").eq(dep2Code);
		d2_on = d2A_on.parents(d2_class).eq(0);
	}

	//처음 상태
	function reset(){
		timer = setTimeout(function(){
			menuOn(d1A_on,d2_on);
			subOn(d2A_on);
		},50);
	}

	//1depth,2depth off
	function menuOff(){
		var curr_dep1 = d1A.filter(".on");
		curr_dep1.removeClass("on");
		d1img.each(function(){
			$(this).imgSwap(srcOn,srcOff);
		});
		d2.stop().animate({height:0},function(){
			$(this).hide();
		});
	}

	//1depth on
	function menuOn(el_dep1,el_dep2){
		menuOff();
		el_dep1.addClass("on");
		el_dep1.find("img").imgSwap(srcOff,srcOn);
		el_dep2.show().stop().animate({height:el_dep2.attr("hgt")+"px"});
	}

	//2depth on
	function subOn(el){
		d2A.filter(".on").find("img").imgSwap(srcOn,srcOff);
		el.find("img").imgSwap(srcOff,srcOn);
		d2A.removeClass("on");
		el.addClass("on");
	}

	//1depth mouseover
	d1A.each(function(i){
		$(this).bind("mouseenter focus", function(){
			clearReset();
			var el_dep1 = $(this);
			var el_dep2 = $(this).parent().find(d2_class);
			menuOn(el_dep1,el_dep2);
		});
	});

	//2depth mouseover
	d2A.bind("mouseenter focus", function(){
		clearReset();
		subOn($(this));
	});

	d1A.filter(":first").bind("blur", function(){
		reset();
	});
	d2A.filter(":last").bind("blur", function(){
		reset();
	});

	//처음 상태로 가기 취소
	function clearReset(){
		clearTimeout(timer);
	}

	//완전히 벗어나면 처음 상태로 가기
	o.bind("mouseleave",function(){
		clearReset();
		reset();
	});
	
	init();
}


//2016-09-06 선택 메뉴 효과 처리
$(document).ready(function(){
	var pathname =  document.location.pathname;
	if( pathname!=null && pathname!=undefined ){
		pathname = pathname.replace("/","");
		var obj = $("#wrap > .btnRight > a");
		if( obj.length > 0 ){
			$( obj ).each(function(){
				if( $(this).attr("href").indexOf(pathname) >= 0 ){
					$(this).addClass("on");
				}
			});
		}
	}
});