

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

function pop_area() {
	var popupW = 600;
	var popupH = 400;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('/area_search.asp', '���ڵ�˻�', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
}

function pop_juso() {
	var popupW = 800;
	var popupH = 400;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('/juso_search.asp', '�ּҷϰ˻�', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=yes, resizable=no');
}

function pop_ce() {
	var popupW = 800;
	var popupH = 300;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('/ce_reg.asp', 'CE�űԵ��', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}

function pop_user_mod() {
	var popupW = 600;
	var popupH = 420;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('/member/user_mod.asp', '����ں���', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}

function pop_id_check() {
	var popupW = 500;
	var popupH = 200;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('/ce_id_check.asp', '���̵��ߺ�Check', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}

function popup_win() {
	var popupW = 450;
	var popupH = 500;
	var left = Math.ceil((window.screen.width - popupW)/2);
	var top = Math.ceil((window.screen.height - popupH)/2);
	window.open('/popup.asp', '�˾�����', 'width='+ popupW +', height='+ popupH +', left='+ left +', top='+ top +', location=no, status=no, menubar=no, toolbar=no, scrollbars=no, resizable=no');
}

//2013-08-21
function pop_Window(theURL,winName,features){ //v2.0
	window.open(theURL,winName,features);
}

/*javascript func �߰� [����ȣ_20210714] */

//�̻�� ���� �˸�
function non_grade(){
	alert("�ش� ��ɿ� ���� ��� ������ �����ϴ�.\n�λ��� Ȥ�� �����ڿ��� ������ �ּ���.");
	return;
}

// ���ڿ����� �޸� ����
function fn_parseNumber(num){
	return ("" + num).replace(/,/g, "");
}

// ���ڿ� �޸� �ֱ�
function fn_numberFormat(num){
	var regexp = /\B(?=(\d{3})+(?!\d))/g;
	return num.toString().replace(regexp, ",");
}

// �������� Ȯ��
function fn_isNumber(val){
	var pattern = /^[0-9]*$/;
	return pattern.test(val);
}

// ���ڸ� �Է�
function fn_inputNumber(){
	if(event.keyCode < 48 || event.keyCode > 57){
		// 0 ~ 9 ������ �ƴϸ� ���� 	false
		event.returnValue = false;
	}
}

// ��ȭ��ȣ ������.
function fn_phoneFormat(obj){
	if(typeof obj == "string"){
		return fn_phoneFormatString(obj);
	} else if(typeof obj == "object"){
		obj.value = fn_phoneFormatString(obj.value);
	}
}

// ��ȭ��ȣ ������.
function fn_phoneFormatString(phoneNo){
	var regexp = /(^02.{0}|^01.{1}|[0-9]{3})([0-9]+)([0-9]{4})/;
	return phoneNo.replace(regexp, "$1-$2-$3");
}

// delim �� ���� ���ڸ� str���� �����Ѵ�.
function fn_removeCharAll(str, delim){
	var regexp = eval("/" + delim + "/gi");
	return str.replace(regexp, "");
}

var n4 = (document.layers)?true:false;
var e4 = (document.all)?true:false;

//���ڸ��Է�(onKeypress='return keyCheckdot(event)')
function keyCheck(e){
	if(n4) var keyValue = e.which
	else if(e4) var keyValue = event.keyCode

	if(((keyValue >= 48) && (keyValue <= 57)) || keyValue==13) return true;
	else return false
}

//���ڹ׵�Ʈ�Է�(onKeypress='return keyCheckdot(event)')
function keyCheckDot(e){
	if(n4) var keyValue = e.which
	else if(e4) var keyValue = event.keyCode

	if(((keyValue >= 48) && (keyValue <= 57)) || keyValue==13 || keyValue==46) return true;
	else return false
}

//��������
function Trim(string){
    for(;string.indexOf(" ")!= -1;){
        string=string.replace(" ","")
    }
    return string;
}

//�Է°˻�
function Exists(input,types){
	if(types) if(!Trim(input.value)) return false;
	return true;
}

//�����˻�+���ڰ˻�(ù���ڴ� �ݵ�ÿ���)
function EngNum(input,types){
	if(types) if(!Trim(input.value)) return false;

	var error_c=0, i, val;

	for(i=0;i<Byte(input.value);i++){
		val = input.value.charAt(i);

		if(i == 0) if(!((val>='a' && val<='z') || (val>='A' && val<='Z'))) return false;
		else if(!((val>=0 && val<=9) || (val>='a' && val<='z') || (val>='A' && val<='Z'))) return false;
	}
	return true;
}

//�����˻�+���ڰ˻�
function EngNumAll(input,types){
	if(types) if(!Trim(input.value)) return false;

	var error_c=0, i, val;

	for(i=0;i<Byte(input.value);i++){
		val = input.value.charAt(i);

		if(!((val>=0 && val<=9) || (val>='a' && val<='z') || (val>='A' && val<='Z'))) return false;
	}
	return true;
}

//�����˻�+���ڰ˻�+'_'
function EngNumAll2(input,types){
	if(types) if(!Trim(input.value)) return false;

	var error_c=0, i, val;

	for(i=0;i<Byte(input.value);i++){
		val = input.value.charAt(i);

		if(!((val>=0 && val<=9) || (val>='a' && val<='z') || (val>='A' && val<='Z') || val=='_')) return false;
	}
	return true;
}

//�����˻�
function Eng(input,types){
	if(types) if(!Trim(input.value)) return false;

	var error_c=0, i, val;

	for(i=0;i<Byte(input.value);i++){
		val = input.value.charAt(i);

		if(!((val>='a' && val<='z') || (val>='A' && val<='Z'))) return false;
	}
	return true;
}

//���ڸ��Է�
function numberonlyinput(){
	var ob = event.srcElement;
	ob.value = noSplitAndNumberOnly(ob);
	return false;
}

//��(3�������� �ĸ��� ���δ�.)
function checkNumber(){
	var ob=event.srcElement;

	ob.value = filterNum(ob.value);
	ob.value = commaSplitAndNumberOnly(ob);
	return false;
}

//������(�����ݾ� �̻��� �Ǹ� �ö���� �ʰ� �Ѵ�.)
function chkhando(money){
	var ob=event.srcElement;

	ob.value = noSplitAndNumberOnly(ob);

	if(ob.value > money) ob.value = money;
	return false;
}

//������(�Ҽ��� ��밡��)
function checkNumberDot(llen,rlen){
	if(llen == "") llen = 8;
	if(rlen == "") rlen = 2;
	var ob=event.srcElement;
	ob.value = filterNum(ob.value);

	spnumber = ob.value.split('.');

	if( spnumber.length >= llen && (spnumber[0].length >llen || spnumber[1].length >llen)) {
		ob.value = spnumber[0].substring(0,llen) + "." + spnumber[1].substring(0,rlen);
		ob.focus();
		return false;
	}else if( spnumber[0].length > llen ) {
		ob.value = spnumber[0].substring(0,llen) + ".";
		ob.focus();
		return false;
	}else if(ob.value && spnumber[0].length == 0) {
		ob.value = 0 + "." + spnumber[1].substring(0,rlen);
		ob.focus();
		return false;
	}

	ob.value = commaSplitAndAllowDot(ob);
	return false;
}

//�����Լ�
function filterNum(str){
	re = /^\$|,/g;
	return str.replace(re, "");
}

//�����Լ�(�ĸ��Ұ�)
function commaSplitAndNumberOnly(ob){
	var txtNumber = '' + ob.value;

	if(isNaN(txtNumber) || txtNumber.indexOf('.') != -1 ){
		ob.value = ob.value.substring(0, ob.value.length-1 );
		ob.value = commaSplitAndNumberOnly(ob);
		ob.focus();
		return ob.value;
	}else{
		var rxSplit = new RegExp('([0-9])([0-9][0-9][0-9][,.])');
		var arrNumber = txtNumber.split('.');
		arrNumber[0] += '.';

		do {
			arrNumber[0] = arrNumber[0].replace(rxSplit, '$1,$2');
		}while (rxSplit.test(arrNumber[0]));

		if (arrNumber.length > 1) {
			return arrNumber.join('');
		}else{
			return arrNumber[0].split('.')[0];
		}
	}
}

//�����Լ�(�ĸ�����)
function commaSplitAndAllowDot(ob){
	var txtNumber = '' + ob.value;

	if(isNaN(txtNumber)){
		ob.value = ob.value.substring(0, ob.value.length-1 );
		ob.focus();
		return ob.value;
	}else{
		var rxSplit = new RegExp('([0-9])([0-9][0-9][0-9][,.])');
		var arrNumber = txtNumber.split('.');
		arrNumber[0] += '.';

		do{
			arrNumber[0] = arrNumber[0].replace(rxSplit, '$1,$2');
		}while(rxSplit.test(arrNumber[0]));

		if (arrNumber.length > 1){
			return arrNumber.join('');
		}else{
			return arrNumber[0].split('.')[0];
		}
	}
}

//���ڸ�����
function noSplitAndNumberOnly(ob){
    var txtNumber = '' + ob.value;
    if (isNaN(txtNumber) || txtNumber.indexOf('.') != -1 ) {
        ob.value = ob.value.substring(0, ob.value.length-1 );
        ob.focus();
        return ob.value;
    }
    else return ob.value;
}

// �ػ󵵿� �´� ũ�� ���
function screensize(){
    self.moveTo(0,0);
    self.resizeTo(screen.availWidth,screen.availHeight);
}

// �ֹε�Ϲ�ȣüũ( �Է��� 1��)
function check_jumin(jumin){
    var weight = "234567892345"; // �ڸ��� weight ����
    var val = jumin.replace("-",""); // "-"(������) ����
    var sum = 0;

    if(val.length != 13){ return false;}

    for(i=0;i<12;i++){
        sum += parseInt(val.charAt(i)) * parseInt(weight.charAt(i));
    }

    var result = (11 - (sum % 11)) % 10;
    var check_val = parseInt(val.charAt(12));

    if(result != check_val){return false;}
    return true;
}

//����Ʈ�˻�
function Byte(input){
	var i, j=0;

	for(i=0;i<input.length;i++){
		val=escape(input.charAt(i)).length;

		if(val== 6) j++;
		j++;
	}
	return j;
}

//�˾��޴�
function popupmenu_show(layername, thislayer, thislayer2){
	thislayerfield.value = thislayer;
	thislayerfield2.value = thislayer2;

	var obj = document.all[layername];
	var _tmpx,_tmpy, marginx, marginy;

	_tmpx = event.clientX + parseInt(obj.offsetWidth);
	_tmpy = event.clientY + parseInt(obj.offsetHeight);
	_marginx = document.body.clientWidth - _tmpx;
	_marginy = document.body.clientHeight - _tmpy;

	if(_marginx < 0) _tmpx = event.clientX + document.body.scrollLeft + _marginx;
	else _tmpx = event.clientX + document.body.scrollLeft;

	if(_marginy < 0) _tmpy = event.clientY + document.body.scrollTop + _marginy + 20;
	else _tmpy = event.clientY + document.body.scrollTop;

	obj.style.posLeft = _tmpx - 5;
	obj.style.posTop = _tmpy;

	layer_set_visible(obj, true);
	layer_set_pos(obj, event.clientX, event.clientY);
}

function layer_set_visible(obj, flag){
	if (navigator.appName.indexOf('Netscape', 0) != -1) obj.visibility = flag ? 'show' : 'hide';
	else obj.style.visibility = flag ? 'visible' : 'hidden';
}

function layer_set_pos(obj, x, y){
	if(navigator.appName.indexOf('Netscape', 0) != -1){
		obj.left = x;
		obj.top = y;
	}else{
		obj.style.pixelLeft = x + document.body.scrollLeft;
		obj.style.pixelTop = y + document.body.scrollTop;
	}
}

//�������̵�
function move(url){
	location.href = url;
}

//�ݱ�
function toclose(){
	self.close();
}

//��ġ����
function winsize(w,h,l,t){
	if(window.opener) resizeTo(w,h);
}

//��Ŀ����ġ
function formfocus(form){
	var len = form.elements.length;

	for(i=0;i<len;i++){
		if((form.elements[i].type == "text" || form.elements[i].type == "password") && Trim(form.elements[i].value) == ""){
			form.elements[i].value = "";
			form.elements[i].focus();
			break;
		}
	}
}

// ��¥,�ð� format �Լ� = php�� date()
function date(arg_format, arg_date){
	if(!arg_date) arg_date = new Date();

	var M = new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec");
	var F = new Array("January","February","March","April","May","June","July","August","September","October","November","December");
	var K = new Array("��","��","ȭ","��","��","��","��");
	var k = new Array("��","��","��","�","��","��","��");
	var D = new Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat");
	var l = new Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday");
	var o = new Array("��","��");
	var O = new Array("����","����");
	var a = new Array("am","pm");
	var A = new Array("AM","PM");

	var org_year = arg_date.getFullYear();
	var org_month = arg_date.getMonth();
	var org_date = arg_date.getDate();
	var org_wday = arg_date.getDay();
	var org_hour = arg_date.getHours();
	var org_minute = arg_date.getMinutes();
	var org_second = arg_date.getSeconds();
	var hour = org_hour % 12; hour = (hour) ? hour : 12;
	var ampm = Math.floor(org_hour / 12);

	var value = new Array();

	value["Y"] = org_year;
	value["y"] = String(org_year).substr(2,2);
	value["m"] = String(org_month+1).replace(/^([0-9])$/,"0$1");
	value["n"] = org_month+1;
	value["d"] = String(org_date).replace(/^([0-9])$/,"0$1");
	value["j"] = org_date;
	value["w"] = org_wday;
	value["H"] = String(org_hour).replace(/^([0-9])$/,"0$1");
	value["G"] = org_hour;
	value["h"] = String(hour).replace(/^([0-9])$/,"0$1");
	value["g"] = hour;
	value["i"] = String(org_minute).replace(/^([0-9])$/,"0$1");
	value["s"] = String(org_second).replace(/^([0-9])$/,"0$1");
	value["t"] = (new Date(org_year, org_month+1, 1) - new Date(org_year, org_month, 1)) / 86400000;
	value["z"] = (new Date(org_year, org_month, org_date) - new Date(org_year, 0, 1)) / 86400000;
	value["L"] = ((new Date(org_year, 2, 1) - new Date(org_year, 1, 1)) / 86400000) - 28;
	value["M"] = M[org_month];
	value["F"] = F[org_month];
	value["K"] = K[org_wday];
	value["k"] = k[org_wday];
	value["D"] = D[org_wday];
	value["l"] = l[org_wday];
	value["o"] = o[ampm];
	value["O"] = O[ampm];
	value["a"] = a[ampm];
	value["A"] = A[ampm];

	var str = "";
	var tag = 0;

	for(i=0;i<arg_format.length;i++) {
		var chr = arg_format.charAt(i);

		switch(chr){
			case "<" : tag++; break;
			case ">" : tag--; break;
		}

		if(tag || value[chr]==null) str += chr; else str += value[chr];
	}
	return str;
}

// �ػ󵵿� �´� ũ�� ���
function screensize(){
	self.moveTo(0,0);
	self.resizeTo(screen.availWidth,screen.availHeight);
}

// �ֹε�Ϲ�ȣüũ( �Է��� 1��)
function check_jumin(jumin){
	var weight = "234567892345"; // �ڸ��� weight ����
	var val = jumin.replace("-",""); // "-"(������) ����
	var sum = 0;

	if(val.length != 13){return false;}

	for(i=0;i<12;i++){
		sum += parseInt(val.charAt(i)) * parseInt(weight.charAt(i));
	}

	var result = (11 - (sum % 11)) % 10;
	var check_val = parseInt(val.charAt(12));

	if(result != check_val){return false;}

	return true;
}

// �ֹε�Ϲ�ȣüũ( �Է��� 2��)
function check_jumin2(input, input2){
	input.value=Trim(input.value);
	input2.value=Trim(input2.value);

	var left_j=input.value;
	var right_j=input2.value;

	if(input.value.length != 6){
		alert('�ֹε�Ϲ�ȣ�� ��Ȯ�� �Է��ϼ���.');
		input.focus();

		return true;
	}

	if(right_j.length != 7){
		alert('�ֹε�Ϲ�ȣ�� ��Ȯ�� �Է��ϼ���.');
		input2.focus();

		return true;
	}

	var i2=0;

	for(var i=0;i<left_j.length;i++){
		var temp=left_j.substring(i,i+1);
		if(temp<0 || temp>9) i2++;
	}

	if((left_j== '') || (i2 != 0)){
		alert('�ֹε�Ϲ�ȣ�� �߸� �ԷµǾ����ϴ�.');
		j_left.focus();

		return true;
	}

	var i3=0;

	for(var i=0;i<right_j.length;i++){
		var temp=right_j.substring(i,i+1);
		if (temp<0 || temp>9) i3++;
	}

	if((right_j== '') || (i3 != 0)){
		alert('�ֹε�Ϲ�ȣ�� �߸� �ԷµǾ����ϴ�.');
		input2.focus();

		return true;
	}

	var l1=left_j.substring(0,1);
	var l2=left_j.substring(1,2);
	var l3=left_j.substring(2,3);
	var l4=left_j.substring(3,4);
	var l5=left_j.substring(4,5);
	var l6=left_j.substring(5,6);
	var hap=l1*2+l2*3+l3*4+l4*5+l5*6+l6*7;
	var r1=right_j.substring(0,1);
	var r2=right_j.substring(1,2);
	var r3=right_j.substring(2,3);
	var r4=right_j.substring(3,4);
	var r5=right_j.substring(4,5);
	var r6=right_j.substring(5,6);
	var r7=right_j.substring(6,7);

	hap=hap+r1*8+r2*9+r3*2+r4*3+r5*4+r6*5;
	hap=hap%11;
	hap=11-hap;
	hap=hap%10;

	if(hap != r7) {
		alert('�ֹε�Ϲ�ȣ�� �߸� �ԷµǾ����ϴ�.');
		input2.focus();

		return true;
	}

	return false;
}

// ��й�ȣ üũ
function check_passwd(input, input2, min){
	if(!input.value){
		alert('��й�ȣ�� �Է��� �ֽʽÿ�.');
		input.focus();

		return false;
	}else if(BYTE(input.value) < min){
		alert('��й�ȣ�� ���̰� �ʹ� ª���ϴ�.');
		input.focus();
		input.value='';
		input2.value='';

		return false;
	}else if(!input2.value){
		alert('Ȯ�κ�й�ȣ�� �Է��� �ֽʽÿ�.');
		input2.focus();

		return false;
	}else if(input.value != input2.value){
		alert('��й�ȣ�� ���� �ٸ��� �ԷµǾ����ϴ�.');
		input2.value='';
		input2.focus();

		return false;
	}else return true;
}

//�޸� �ֱ�(������ �ش�)
function comma(val){
	val = get_number(val);
	if(val.length <= 3) return val;

	var loop = Math.ceil(val.length / 3);
	var offset = val.length % 3;

	if(offset==0) offset = 3;
	var ret = val.substring(0, offset);

	for(i=1;i<loop;i++) {
		ret += "," + val.substring(offset, offset+3);
		offset += 3;
	}
	return ret;
}

//���ڿ����� ���ڸ� ��������
function get_number(str){
	var val = str;
	var temp = "";
	var num = "";

	for(i=0; i<val.length; i++){
		temp = val.charAt(i);
		if(temp >= "0" && temp <= "9") num += temp;
	}
	return num;
}

//�ֹε�Ϲ�ȣ�� ���̷� ��ȯ
function agechange(lno,rno){
	var refArray = new Array(18,19,19,20,20,16,16,17,17,18);
	var refyy = rno.substring(0,1);
	var refno = lno.substring(0,2);
	var biryear = refArray[refyy] * 100 + eval(refno);

	var nowDate = new Date();
	var nowyear = nowDate.getYear();

	return nowyear - biryear + 1;
}

//������ڽ� üũ�˻�
function radio_chk(input, msg){
	var len = input.length;

	for(var i=0;i<len;i++) if(input[i].checked == true && input[i].value) return true;

	alert(msg);

	return false;
}

//����Ʈ�ڽ� üũ�˻�
function select_chk(input, msg){
	if(input[0].selected == true){
		alert(msg);

		return false;
	}
	return true;
}

//��â����
function open_window(url, target, w, h, s) {
	if(s) s = 'yes';
	else s = 'no';

	var its = window.open(url,target,'width='+w+',height='+h+',top=0,left=0,scrollbars='+s);
	its.focus();
}


//��â�ݱ�
function close_win(){
	window.close();
}

//������ Ȯ��
function isBrowserCheck(){
	const agt = navigator.userAgent.toLowerCase();

	if (agt.indexOf("chrome") != -1) return 'Chrome';
	if (agt.indexOf("opera") != -1) return 'Opera';
	if (agt.indexOf("staroffice") != -1) return 'Star Office';
	if (agt.indexOf("webtv") != -1) return 'WebTV';
	if (agt.indexOf("beonex") != -1) return 'Beonex';
	if (agt.indexOf("chimera") != -1) return 'Chimera';
	if (agt.indexOf("netpositive") != -1) return 'NetPositive';
	if (agt.indexOf("phoenix") != -1) return 'Phoenix';
	if (agt.indexOf("firefox") != -1) return 'Firefox';
	if (agt.indexOf("safari") != -1) return 'Safari';
	if (agt.indexOf("skipstone") != -1) return 'SkipStone';
	if (agt.indexOf("netscape") != -1) return 'Netscape';
	if (agt.indexOf("mozilla/5.0") != -1) return 'Mozilla';
	if (agt.indexOf("msie") != -1) {
    	let rv = -1;
		if (navigator.appName == 'Microsoft Internet Explorer') {
			let ua = navigator.userAgent; var re = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
			if (re.exec(ua) != null)
				rv = parseFloat(RegExp.$1);
			}
		return 'Internet Explorer '+rv;
	}
}

/**
     * ���ڿ��� �� ���ڿ����� üũ�Ͽ� ������� �����Ѵ�.
     * @param str       : üũ�� ���ڿ�
     */
 function isEmpty(str){

	if(typeof str == "undefined" || str == null || str == "")
		return true;
	else
		return false ;
}

/**
 * ���ڿ��� �� ���ڿ����� üũ�Ͽ� �⺻ ���ڿ��� �����Ѵ�.
 * @param str           : üũ�� ���ڿ�
 * @param defaultStr    : ���ڿ��� ���������� ������ �⺻ ���ڿ�
 */
function nvl(str, defaultStr){

	if(typeof str == "undefined" || str == null || str == "")
		str = defaultStr ;

	return str ;
}