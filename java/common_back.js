

$(function(){

 $("#gnb").GNB();
 
 $(".bid").tabUI({
		tab : ".tabUI_tab" //��
		,con : ".tabUI_con" //������
		,tabHover : "on" //�� ������ Ŭ����
		,eventHandler : "click focus" //�̺�Ʈ�ڵ鷯
		,fn : "" //�� Ȱ��ȭ�� �� ������ �Լ�
	});

})

function pop_linkage() {
	var popupW = 600;
	var popupH = 720;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('pop_linkage.asp', '�����۾�����', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}
function pop_manager() {
	var popupW = 600;
	var popupH = 500;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('manager_add.asp', '��ü���', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}
function pop_codeSearch() {
	var popupW = 481;
	var popupH = 500;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('pop.asp', 'ǰ����ڵ�ã��', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=yes');
}
function pop_warehousing() {
	var popupW = 600;
	var popupH = 500;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('warehousing.asp', '�԰�����', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}
function pop_estimate() {
	var popupW = 500;
	var popupH = 410;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('estimate_reg.asp', '�����۾�����', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}
function pop_estimate1(bid_seq) {
	var popupW = 500;
	var popupH = 500;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('estimate_reg1.asp?bid_seq='+bid_seq, '�����۾�����', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}
function pop_wh_register(bid_seq) {
	var popupW = 600;
	var popupH = 550;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('wh_register.asp?bid_seq=' + bid_seq, '�԰������', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}