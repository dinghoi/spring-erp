

$(function(){

 $("#gnb").GNB();
 
 $(".bid").tabUI({
		tab : ".tabUI_tab" //탭
		,con : ".tabUI_con" //컨텐츠
		,tabHover : "on" //탭 오버시 클래스
		,eventHandler : "click focus" //이벤트핸들러
		,fn : "" //탭 활성화될 때 실행할 함수
	});

})

function pop_linkage() {
	var popupW = 600;
	var popupH = 720;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('pop_linkage.asp', '연계작업내역', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}
function pop_manager() {
	var popupW = 600;
	var popupH = 500;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('manager_add.asp', '업체등록', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}
function pop_codeSearch() {
	var popupW = 481;
	var popupH = 500;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('pop.asp', '품명및코드찾기', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=yes');
}
function pop_warehousing() {
	var popupW = 600;
	var popupH = 500;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('warehousing.asp', '입고내역서', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}
function pop_estimate() {
	var popupW = 500;
	var popupH = 410;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('estimate_reg.asp', '연계작업내역', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}
function pop_estimate1(bid_seq) {
	var popupW = 500;
	var popupH = 500;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('estimate_reg1.asp?bid_seq='+bid_seq, '연계작업내역', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}
function pop_wh_register(bid_seq) {
	var popupW = 600;
	var popupH = 550;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('wh_register.asp?bid_seq=' + bid_seq, '입고내역등록', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}