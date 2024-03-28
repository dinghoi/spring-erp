/***********************************************
   ��    �� : js_form.js
   ��    �� : ����� �� �Է°� �˻� ���� �Լ�
***********************************************/

// �� ���ڿ����� �˻�
function isempty (data) { 
	for (var i = 0; i < data.length; i++) 
		if (data.substring(i, i+1) != " ")
			return false;
	return true;
}


// �ѱ����� �˻� (Ȥ�� ASCII�� ������ ���� 2����Ʈ ����)
function ishangul (chr) {
	var		key_eg ;

	key_eg = (escape (chr)).charAt (1) ;

	switch (key_eg) {
		case "u":
		case "b":
			return true ;
		default:
			return false ;
	}
}

// ���ڿ� ���� (�ѱ��� 2�� ���̸� ���� ������ ����)
function h_stringlength(string) {   
	char_cnt = 0;   

	for(var i = 0; i < string.length; i++) {
		var chr = string.substr(i,1);   
		chr = escape(chr);   
		key_eg = chr.charAt(1);   

		switch (key_eg) {   
			case "u":   
				key_num = chr.substr(2,(chr.length - 1)) ;   
							
//				if((key_num < "AC00") || (key_num > "D7A3"))
//					return -1 ;   
//				else
					char_cnt = char_cnt + 2;   
				break;   

			case "B":   
				char_cnt = char_cnt + 2;   
				break;   

			default:   
				char_cnt = char_cnt + 1;   
		}
	}
	return char_cnt ;
}


// text ���� �Է»��ڿ� onlyemail �Ӽ��� �ִ� ���
// �ùٸ� �̸������� �˻�
// onlyemail
function formemailcheck (form) {
	var			what, i, j ;
	var			emailexp ;

//	emailexp = new RegExp ("^(..+)@(..+)\\.(..+)$") ;
	emailexp = new RegExp ("^[A-Za-z0-9-_\\.]{2,}@[A-Za-z0-9-_\\.]{2,}\\.[A-Za-z0-9-_]{2,}$") ;

	for (i = 0; i < form.length; i++) {
		what = form.elements[i] ;

		if ( what.onlyemail != null ) {
			if ( what.type == "text" ) {
				if ( what.value == "" )
					continue ;

				if ( !emailexp.test (what.value) ) {
					alert (what.errname + "�� �ùٸ��� �ʽ��ϴ�") ;
					what.focus () ;
					return false ;
				}

				for (j = 0; j < what.value.length; j++) {
					if ( ishangul (what.value.charAt (j)) ) {
						alert (what.errname + "�� �ùٸ��� �ʽ��ϴ�") ;
						what.focus () ;
						return false ;
					}
				}
			}
		}
	}
	return true ;
}

function formemailcheck_and_submit (form) {
	if ( formemailcheck (form) )
		form.submit () ;
}


// text, password, textarea, file ���� �Է»��ڿ� notnull �Ӽ��� �ִ� ���
// ���� ����ִ��� �˻�
// ���� ���� ��� true, ���� ���� ��� false �� ����
// notnull
function formnullcheck (form) {
	var			what, i ;

	for (i = 0; i < form.length; i++) {
		what = form.elements[i] ;

		if ( what.notnull != null ) {
			if ( what.type == "text" || what.type == "password" 
				|| what.type == "textarea" || what.type == "file" )
				if ( isempty (what.value) ) {
					alert (what.errname + "��(��) �Է��ϼ���") ;
					what.focus () ;
					return false ;
				}				
		}
	}
	return true ;
}	

// text, password, textarea, file ���� �Է»����� �ּұ��� �� �ִ���� �˻�
// �ּұ��� �� �ִ���� �׽�Ʈ ����� true, ���н� false �� ����
// maxlength, minlength
function formlencheck(form){
	var what, i;
	var str_len, errname;

	for(i = 0; i < form.length; i++){
		what = form.elements[i];

		if(what.type === "text" || what.type === "password" || what.type === "textarea" || what.type === "file"){
			if(what.value === "" ){
				continue;
			}

			if(what.type === "file"){
				var	rslash_pos;
				var	fname;

				rslash_pos = what.value.lastIndexOf("\\", what.value.length);

				if(rslash_pos === -1){
					alert(what.errname + "�� �����̸��� �߸��Ǿ����ϴ�\n�ٽ��ѹ� Ȯ���ϼ���");
					what.focus();

					return false;
				}

				fname = what.value.substr(rslash_pos + 1);
				str_len = h_stringlength(fname);
				errname = what.errname + " �����̸�";
			}else{
				str_len = h_stringlength(what.value);
				errname = what.errname;
			}		
			
//			if ( str_len == -1 ) {
//				alert ("�߸��� ���ڿ��Դϴ�.") ;
//				what.focus () ;
//				return false ;
//			}			

			//�˻��� ���� �˻� ���� ����[����ȣ_20210707]
			// �ּ� ���� �˻�
			if(what.minLength !== null){
				/*if(str_len < what.minLength){
					alert(errname + "��(��) ���� " + what.minLength + "��, �ѱ� " + Math.ceil(what.minLength / 2) + "�� �̻� �Է��ؾ� �մϴ�");
					what.focus();
					return false;
				}*/
				if(str_len < 1){
					alert("�˻����(��) ���� 1��, �ѱ� 1�� �̻� �Է��ؾ� �մϴ�");
					what.focus();

					return false;
				}
			}

			// �ִ� ���� �˻�
			/*if(str_len > what.maxLength){
				alert(errname + "��(��) ���� " + what.maxLength + "��, �ѱ� " + parseInt(what.maxLength / 2) + "�� ���� �Է��� �� �ֽ��ϴ�");
				what.focus();

				return false;
			}*/
			if(str_len > 50){
				alert("�˻����(��) ���� 50��, �ѱ� 25�� �̻� �Է��ؾ� �մϴ�");
				what.focus();

				return false;
			}
		}
	}

	return true ;			
}

// null �˻� �� ������ submit
function formnullcheck_and_submit (form) {
	if ( formnullcheck (form) )
		form.submit () ;
}

// ���� �˻� �� ������ submit
function formlencheck_and_submit (form) {
	if ( formlencheck (form) )
		form.submit () ;
}


// �������� �˻�. ������ ��� true ����
function isnumber (num) {
	return !isNaN (num) ;
}

// �������� �˻�, ������ ��� true ����
function isinteger (num) {
	if ( !isnumber (num) )
		return false ;
	
	if ( num.indexOf (".") != -1 )
		return false ;
		
	return true ;
}

// ���� �������� �˻�, ����� ��� true ����
function isposint (num) {
	if ( !isinteger (num) )
		return false ;

	if ( parseInt (num) < 0 )
		return false ;

	return true ;
}

// ���� �������� �˻�, ������ ��� true ����
function isnegint (num) {
	if ( !isinteger (num) )
		return false ;

	if ( parseInt (num) > 0 )
		return false ;

	return true ;
}


// text, password, textarea, file ������ ����
// ����(����, ����, ���, ����)���� �˻�
// �ּҰ�, �ִ밪 �˻�
// onlynum, onlyint, onlyposint, onlynegint, minnum, maxnum
function formnumcheck (form) {
	var			what, i ;

	for (i = 0; i < form.length; i++) {
		what = form.elements[i] ;

		if ( what.type == "text" || what.type == "password" 
			|| what.type == "textarea" || what.type == "file" ) {

			// ���� �������� �˻�
			if ( what.onlyposint != null ) {
				if ( !isposint (what.value) ) {
					alert (what.errname + "��(��) " +
							"���� ������ �Է��� �� �ֽ��ϴ�") ;
					what.focus () ;
					return false ;
				}
			}

			// ���� �������� �˻�
			if ( what.onlynegint != null ) {
				if ( !isnegint (what.value) ) {
					alert (what.errname + "��(��) " +
							"���� ������ �Է��� �� �ֽ��ϴ�") ;
					what.focus () ;
					return false ;
				}
			}

			// �������� �˻�
			if ( what.onlyint != null ) {
				if ( !isinteger (what.value) ) {
					alert (what.errname + "��(��) " +
							"������ �Է��� �� �ֽ��ϴ�") ;
					what.focus () ;
					return false ;
				}
			}

			// �������� �˻�
			if ( what.onlynum != null ) {
				if ( !isnumber (what.value) ) {
					alert (what.errname + "��(��) " +
							"���ڸ� �Է��� �� �ֽ��ϴ�") ;
					what.focus () ;
					return false ;
				}
			}

			// �ּҰ� �˻�
			if ( what.minnum != null ) {
				if ( parseInt (what.value) < parseInt (what.minnum) ) {
					alert (what.errname + "��(��) " +
							what.minnum +
							" �̻��� ���ڸ� �Է��ؾ� �մϴ�") ;
					what.focus () ;
					return false ;
				}
			}

			// �ִ밪 �˻�
			if ( what.maxnum != null ) {
				if ( parseInt (what.value) > parseInt (what.maxnum) ) {
					alert (what.errname + "��(��) " +
							what.maxnum +
							" ������ ���ڸ��� �Է��� �� �ֽ��ϴ�") ;
					what.focus () ;
					return false ;
				}
			}
		}
	}
	return true ;
}

// �������� �˻��� submit
function formnumcheck_and_submit (form) {
	if ( formnumcheck (form) )
		form.submit () ;
}


// str �� perm_char �� �ִ� ���ڵ�� ������
// ��� true, �ƴϸ� false ����
function check_perm_char (str, perm_char) {
	var			i, chr ;

	if ( str == "" )
		return true ;
	
	for (i = 0; i < str.length; i++) {
		chr = str.substr (i, 1) ;

		if ( perm_char.indexOf (chr) == -1 )
			return false ;
	}
	return true ;
}

// str �� rej_char �� �ִ� ���ڵ��� ����
// ��� true, �ƴϸ� false ����
function check_reject_char (str, rej_char) {
	var			i, chr ;

	if ( str == "" )
		return true ;

	for (i = 0; i < rej_char.length; i++) {
		chr = rej_char.substr (i, 1) ;

		if ( str.indexOf (chr) != -1 )
			return false ;
	}
	return true ;
}

function get_str_breaked_by_comma (str) {
	var			i, rtn_str ;
	
	if ( str == "" )
		return "" ;

	rtn_str = "" ;
	for (i = 0; i < str.length; i++) {
		rtn_str += str.substr (i, 1) ;

		if ( i != str.length - 1 )
			rtn_str += ", " ;
	}
	return rtn_str ;
}


// text, password, textarea, file ������ ����
// �㰡 ���� (permchr) �Ǵ� �ź� ���� (rejchr)
// �� �ִ��� �˻�
// permchr, rejchr
function formstrcheck (form) {
	var			what, i ;

	for (i = 0; i < form.length; i++) {
		what = form.elements[i] ;

		if ( what.type == "text" || what.type == "password" 
			|| what.type == "textarea" || what.type == "file" ) {
	
			// �㰡 ���� �˻�
			if ( what.permchr != null ) {
				if ( !check_perm_char (what.value, what.permchr) ) {
					alert (what.errname + "��(��) " +
							"������ ���� '" + 
							get_str_breaked_by_comma (what.permchr) +
							"' �� �� �� �ֽ��ϴ�") ;
					what.focus () ;
					return false ;	
				}
			}

			// �ź� ���� �˻�
			if ( what.rejchr != null ) {
				if ( !check_reject_char (what.value, what.rejchr) ) {
					alert (what.errname + "��(��) " +
							"������ ���� '" + 
							get_str_breaked_by_comma (what.rejchr) +
							"' �� �� �� �����ϴ�") ;
					what.focus () ;
					return false ;	
				}
			}
		}
	}
	return true ;
}

// �� ���ڿ� �˻� �� submit ()
function formstrcheck_and_submit (form) {
	if ( formstrcheck (form) )
		form.submit ()
}


// text, password, textarea, file ������ ����
// �����̿��� ���ڰ� �ִ��� �˻�
// onlyeng
function formengcheck (form) {
	var			what, i ;
	var			cap_letter ;
	var			low_letter ;
	var			num_letter ;
	var			sign_letter ;
	var			eng_letter ;
	
	cap_letter = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" ;
	low_letter = "abcdefghijklmnopqrstuvwxyz" ;
	num_letter = "0123456789" ;
	sign_letter = "`~!@#$%^&*()-_=+\\|[{]};:'\",<.>/?" ;
	white_space = " \n\r\t" ;
	eng_letter = cap_letter + low_letter +
		num_letter + sign_letter + eng_letter +
		white_space ;
	
	for (i = 0; i < form.length; i++) {
		what = form.elements[i] ;

		if ( what.type == "text" || what.type == "password" 
			|| what.type == "textarea" || what.type == "file" ) {

			// �������� �˻�
			if ( what.onlyeng != null ) {
				if ( !check_perm_char (what.value, eng_letter) ) {
					alert (what.errname + "��(��) " +
							"��� �Է��� �� �ֽ��ϴ�") ;
					what.focus () ;
					return false ;
				}
			}
		}
	}
	return true ;
}

// �������� �˻��ϰ� submit
function formengcheck_and_submit (form) {
	if ( formengcheck (form) )
		form.submit () ;
}


// ��� �˻� ����
function formcheck(form){
	if(!formnullcheck (form) || !formlencheck (form) || !formnumcheck (form) || !formstrcheck (form) || !formengcheck (form) || !formemailcheck (form)){
		return false;
	}

	return true;
}

// ��� �˻� �� ������ submit
function formcheck_and_submit (form) {
	if ( formcheck (form) )
		form.submit () ;
}

/* Trim �ϴ� �Լ� */
function TrimString(SrcString)
{

   /* ���� Ʈ��   */
   len = SrcString.length;
   for(i=0;i<len;i++)
   {
      if(SrcString.substring(0,1) == " ")
      {
         SrcString = SrcString.substring(1);
      }
      else
      {
         break;
      }
   }

   /* ������ Ʈ��   */
   len = SrcString.length;
   for(i=len;i>0;i--)
   {
      if(SrcString.substring(i-1) == " ")
      {
         SrcString = SrcString.substring(0,i-1);
      }
      else
      {
         break;
      }
   }

   return SrcString;
}
function roundXL(n, digits) {
  if (digits >= 0) return parseFloat(n.toFixed(digits)); // �Ҽ��� �ݿø�

  digits = Math.pow(10, digits); // ������ �ݿø�
  var t = Math.round(n * digits) / digits;

  return parseFloat(t.toFixed(0));
}
function plusComma(txtObj){
	if (txtObj.value.length<1) {
		txtObj.value=txtObj.value.replace(/,/g,"");
		txtObj.value=txtObj.value.replace(/\D/g,"");
	}
	var num = txtObj.value;
	if (num == "--" ||  num == "." ) num = "";
	if (num != "" ) {
		temp=new String(num);
		if(temp.length<1) return "";
					
		// ����ó��
		if(temp.substr(0,1)=="-") minus="-";
			else minus="";
					
		// �Ҽ�������ó��
		dpoint=temp.search(/\./);
				
		if(dpoint>0)
		{
		// ù��° ������ .�� �������� �ڸ��� ���������� ���� ����
		dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
		temp=temp.substr(0,dpoint);
		}else dpointVa="";
					
		// �����ܹ̿��� ����
		temp=temp.replace(/\D/g,"");
		zero=temp.search(/[1-9]/);
				
		if(zero==-1) return "";
		else if(zero!=0) temp=temp.substr(zero);
					
		if(temp.length<4) return minus+temp+dpointVa;
		buf="";
		while (true)
		{
		if(temp.length<3) { buf=temp+buf; break; }
			
		buf=","+temp.substr(temp.length-3)+buf;
		temp=temp.substr(0, temp.length-3);
		}
		if(buf.substr(0,1)==",") buf=buf.substr(1);
				
		//return minus+buf+dpointVa;
		txtObj.value = minus+buf+dpointVa;
	}else txtObj.value = "0";					
}
function plusComma1(txtObj){
	if (txtObj.value.length<1) {
		txtObj.value=txtObj.value.replace(/,/g,"");
		txtObj.value=txtObj.value.replace(/\D/g,"");
	}
	var num = txtObj.value;
	if (num == "--" ||  num == "." ) num = "";
	if (num != "" ) {
		temp=new String(num);
		if(temp.length<1) return "";
					
		// ����ó��
		if(temp.substr(0,1)=="-") minus="-";
			else minus="";
					
		// �Ҽ�������ó��
		dpoint=temp.search(/\./);
				
		if(dpoint>0)
		{
		// ù��° ������ .�� �������� �ڸ��� ���������� ���� ����
		dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
		temp=temp.substr(0,dpoint);
		}else dpointVa="";
					
		// �����ܹ̿��� ����
		temp=temp.replace(/\D/g,"");
		zero=temp.search(/[1-9]/);
				
		if(zero==-1) return "";
		else if(zero!=0) temp=temp.substr(zero);
					
		if(temp.length<4) return minus+temp+dpointVa;
		buf="";
		while (true)
		{
		if(temp.length<3) { buf=temp+buf; break; }
			
		buf=","+temp.substr(temp.length-3)+buf;
		temp=temp.substr(0, temp.length-3);
		}
		if(buf.substr(0,1)==",") buf=buf.substr(1);
				
		//return minus+buf+dpointVa;
		txtObj.value = minus+buf+dpointVa;
	}else txtObj.value = "";					
}
function checkNum(obj) 
{
	 	var word = obj.value;
	 	var str = "-1234567890";
	 	for (i=0;i< word.length;i++){
	 	    if(str.indexOf(word.charAt(i)) < 0){
	 	        alert("���� ���ո� �����մϴ�..");
	 	        obj.value="";
	 	        obj.focus();
	 	        return false;
	 	    }
	 	}
}
function nbytes(str) 
{ // korean chars count as two bytes, others count as one. 
    var i,sum=0; 
    for(i=0;i<str.length;i++) sum += (str.charCodeAt(i) > 255 ? 2 : 1); 
    return sum; 
} 
function checklength(x,limit) 
{ // use onkeyup event 
    if((y=nbytes(x.value)) >= limit) alert('���� '+limit+' �ִ밪 : ' + y); 
    if((y=nbytes(x.value)) == limit) alert('���� '+limit+' �ʰ� : ' + y); 
} 


//2016-09-07 callback ȸ��� ����Ʈ selectBox ����
function setCompanySelect(targetId, defaultKey, data){
	var opt = "<option value=\"\">��ü</option>";
	var isSelected = "";

	if( data ){
		var result = data.result;

		$.each(result, function(i, item){
			if( defaultKey!="" && defaultKey == item.orgCompany ){
				isSelected = " selected=\"selected\"";
			}else{
				isSelected = "";
			}
			opt+= "<option value=\""+item.orgCompany+"\"" + isSelected + ">"+item.orgCompany+"</option>";
			
		});
	}
	$("#"+targetId).html(opt);
}

//2016-09-07 get company list
function getOrg(type, company, bonbu, saupbu, team, defaultKey, targetId, callback){

	var params = "srchType="+type
					+ "&company="+company
					+ "&bonbu="+bonbu
					+ "&saupbu="+saupbu
					+ "&team="+team;

	$.ajax({
		url:"/include/ajax/getOrg.asp"
		,type:'post'
		,data: params
		,contentType: "application/x-www-form-urlencoded; charset=euc-kr"
		,dataType: "json"
		,success:function(data, status, request){
			callback(targetId, defaultKey, data);
		}
		,error:function(jqXHR, status, errorThrown){
			alert("������ �߻��Ͽ����ϴ�.\n�����ڵ� : " + status + " : " + errorThrown);
		}
	});
}