/***********************************************
   파    일 : js_form.js
   설    명 : 사용자 폼 입력값 검사 관련 함수
***********************************************/

// 빈 문자열인지 검사
function isempty (data) { 
	for (var i = 0; i < data.length; i++) 
		if (data.substring(i, i+1) != " ")
			return false;
	return true;
}


// 한글인지 검사 (혹은 ASCII를 제외한 문자 2바이트 문자)
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

// 문자열 길이 (한글을 2의 길이를 갖는 것으로 적용)
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


// text 형식 입력상자에 onlyemail 속성이 있는 경우
// 올바른 이메일인지 검사
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
					alert (what.errname + "이 올바르지 않습니다") ;
					what.focus () ;
					return false ;
				}

				for (j = 0; j < what.value.length; j++) {
					if ( ishangul (what.value.charAt (j)) ) {
						alert (what.errname + "이 올바르지 않습니다") ;
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


// text, password, textarea, file 형식 입력상자에 notnull 속성이 있는 경우
// 값이 비어있는지 검사
// 빈값이 없을 경우 true, 빈값이 있을 경우 false 를 리턴
// notnull
function formnullcheck (form) {
	var			what, i ;

	for (i = 0; i < form.length; i++) {
		what = form.elements[i] ;

		if ( what.notnull != null ) {
			if ( what.type == "text" || what.type == "password" 
				|| what.type == "textarea" || what.type == "file" )
				if ( isempty (what.value) ) {
					alert (what.errname + "을(를) 입력하세요") ;
					what.focus () ;
					return false ;
				}				
		}
	}
	return true ;
}	

// text, password, textarea, file 형식 입력상자의 최소길이 및 최대길이 검사
// 최소길이 및 최대길이 테스트 통과시 true, 실패시 false 를 리턴
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
					alert(what.errname + "의 파일이름이 잘못되었습니다\n다시한번 확인하세요");
					what.focus();

					return false;
				}

				fname = what.value.substr(rslash_pos + 1);
				str_len = h_stringlength(fname);
				errname = what.errname + " 파일이름";
			}else{
				str_len = h_stringlength(what.value);
				errname = what.errname;
			}		
			
//			if ( str_len == -1 ) {
//				alert ("잘못된 문자열입니다.") ;
//				what.focus () ;
//				return false ;
//			}			

			//검색어 길이 검사 오류 수정[허정호_20210707]
			// 최소 길이 검사
			if(what.minLength !== null){
				/*if(str_len < what.minLength){
					alert(errname + "은(는) 영문 " + what.minLength + "자, 한글 " + Math.ceil(what.minLength / 2) + "자 이상 입력해야 합니다");
					what.focus();
					return false;
				}*/
				if(str_len < 1){
					alert("검색어는(은) 영문 1자, 한글 1자 이상 입력해야 합니다");
					what.focus();

					return false;
				}
			}

			// 최대 길이 검사
			/*if(str_len > what.maxLength){
				alert(errname + "은(는) 영문 " + what.maxLength + "자, 한글 " + parseInt(what.maxLength / 2) + "자 까지 입력할 수 있습니다");
				what.focus();

				return false;
			}*/
			if(str_len > 50){
				alert("검색어는(은) 영문 50자, 한글 25자 이상 입력해야 합니다");
				what.focus();

				return false;
			}
		}
	}

	return true ;			
}

// null 검사 후 성공시 submit
function formnullcheck_and_submit (form) {
	if ( formnullcheck (form) )
		form.submit () ;
}

// 길이 검사 후 성공시 submit
function formlencheck_and_submit (form) {
	if ( formlencheck (form) )
		form.submit () ;
}


// 숫자인지 검사. 숫자일 경우 true 리턴
function isnumber (num) {
	return !isNaN (num) ;
}

// 정수인지 검사, 정수일 경우 true 리턴
function isinteger (num) {
	if ( !isnumber (num) )
		return false ;
	
	if ( num.indexOf (".") != -1 )
		return false ;
		
	return true ;
}

// 양의 정수인지 검사, 양수일 경우 true 리턴
function isposint (num) {
	if ( !isinteger (num) )
		return false ;

	if ( parseInt (num) < 0 )
		return false ;

	return true ;
}

// 음의 정수인지 검사, 음수일 경우 true 리턴
function isnegint (num) {
	if ( !isinteger (num) )
		return false ;

	if ( parseInt (num) > 0 )
		return false ;

	return true ;
}


// text, password, textarea, file 형식의 값이
// 숫자(숫자, 정수, 양수, 음수)인지 검사
// 최소값, 최대값 검사
// onlynum, onlyint, onlyposint, onlynegint, minnum, maxnum
function formnumcheck (form) {
	var			what, i ;

	for (i = 0; i < form.length; i++) {
		what = form.elements[i] ;

		if ( what.type == "text" || what.type == "password" 
			|| what.type == "textarea" || what.type == "file" ) {

			// 양의 정수인지 검사
			if ( what.onlyposint != null ) {
				if ( !isposint (what.value) ) {
					alert (what.errname + "은(는) " +
							"양의 정수만 입력할 수 있습니다") ;
					what.focus () ;
					return false ;
				}
			}

			// 음의 정수인지 검사
			if ( what.onlynegint != null ) {
				if ( !isnegint (what.value) ) {
					alert (what.errname + "은(는) " +
							"음의 정수만 입력할 수 있습니다") ;
					what.focus () ;
					return false ;
				}
			}

			// 정수인지 검사
			if ( what.onlyint != null ) {
				if ( !isinteger (what.value) ) {
					alert (what.errname + "은(는) " +
							"정수만 입력할 수 있습니다") ;
					what.focus () ;
					return false ;
				}
			}

			// 숫자인지 검사
			if ( what.onlynum != null ) {
				if ( !isnumber (what.value) ) {
					alert (what.errname + "은(는) " +
							"숫자만 입력할 수 있습니다") ;
					what.focus () ;
					return false ;
				}
			}

			// 최소값 검사
			if ( what.minnum != null ) {
				if ( parseInt (what.value) < parseInt (what.minnum) ) {
					alert (what.errname + "은(는) " +
							what.minnum +
							" 이상의 숫자를 입력해야 합니다") ;
					what.focus () ;
					return false ;
				}
			}

			// 최대값 검사
			if ( what.maxnum != null ) {
				if ( parseInt (what.value) > parseInt (what.maxnum) ) {
					alert (what.errname + "은(는) " +
							what.maxnum +
							" 이하의 숫자만을 입력할 수 있습니다") ;
					what.focus () ;
					return false ;
				}
			}
		}
	}
	return true ;
}

// 숫자인지 검사후 submit
function formnumcheck_and_submit (form) {
	if ( formnumcheck (form) )
		form.submit () ;
}


// str 이 perm_char 에 있는 문자들로 구성될
// 경우 true, 아니면 false 리턴
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

// str 에 rej_char 에 있는 문자들이 없을
// 경우 true, 아니면 false 리턴
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


// text, password, textarea, file 형식의 값에
// 허가 문자 (permchr) 또는 거부 문자 (rejchr)
// 가 있는지 검사
// permchr, rejchr
function formstrcheck (form) {
	var			what, i ;

	for (i = 0; i < form.length; i++) {
		what = form.elements[i] ;

		if ( what.type == "text" || what.type == "password" 
			|| what.type == "textarea" || what.type == "file" ) {
	
			// 허가 문자 검사
			if ( what.permchr != null ) {
				if ( !check_perm_char (what.value, what.permchr) ) {
					alert (what.errname + "은(는) " +
							"다음의 문자 '" + 
							get_str_breaked_by_comma (what.permchr) +
							"' 만 쓸 수 있습니다") ;
					what.focus () ;
					return false ;	
				}
			}

			// 거부 문자 검사
			if ( what.rejchr != null ) {
				if ( !check_reject_char (what.value, what.rejchr) ) {
					alert (what.errname + "은(는) " +
							"다음의 문자 '" + 
							get_str_breaked_by_comma (what.rejchr) +
							"' 를 쓸 수 없습니다") ;
					what.focus () ;
					return false ;	
				}
			}
		}
	}
	return true ;
}

// 폼 문자열 검사 후 submit ()
function formstrcheck_and_submit (form) {
	if ( formstrcheck (form) )
		form.submit ()
}


// text, password, textarea, file 형식의 값에
// 영어이외의 문자가 있는지 검사
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

			// 영어인지 검사
			if ( what.onlyeng != null ) {
				if ( !check_perm_char (what.value, eng_letter) ) {
					alert (what.errname + "은(는) " +
							"영어만 입력할 수 있습니다") ;
					what.focus () ;
					return false ;
				}
			}
		}
	}
	return true ;
}

// 영어인지 검사하고 submit
function formengcheck_and_submit (form) {
	if ( formengcheck (form) )
		form.submit () ;
}


// 모든 검사 수행
function formcheck(form){
	if(!formnullcheck (form) || !formlencheck (form) || !formnumcheck (form) || !formstrcheck (form) || !formengcheck (form) || !formemailcheck (form)){
		return false;
	}

	return true;
}

// 모든 검사 후 성공시 submit
function formcheck_and_submit (form) {
	if ( formcheck (form) )
		form.submit () ;
}

/* Trim 하는 함수 */
function TrimString(SrcString)
{

   /* 왼쪽 트림   */
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

   /* 오른쪽 트림   */
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
  if (digits >= 0) return parseFloat(n.toFixed(digits)); // 소수부 반올림

  digits = Math.pow(10, digits); // 정수부 반올림
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
					
		// 음수처리
		if(temp.substr(0,1)=="-") minus="-";
			else minus="";
					
		// 소수점이하처리
		dpoint=temp.search(/\./);
				
		if(dpoint>0)
		{
		// 첫번째 만나는 .을 기준으로 자르고 숫자제외한 문자 삭제
		dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
		temp=temp.substr(0,dpoint);
		}else dpointVa="";
					
		// 숫자이외문자 삭제
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
					
		// 음수처리
		if(temp.substr(0,1)=="-") minus="-";
			else minus="";
					
		// 소수점이하처리
		dpoint=temp.search(/\./);
				
		if(dpoint>0)
		{
		// 첫번째 만나는 .을 기준으로 자르고 숫자제외한 문자 삭제
		dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
		temp=temp.substr(0,dpoint);
		}else dpointVa="";
					
		// 숫자이외문자 삭제
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
	 	        alert("숫자 조합만 가능합니다..");
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
    if((y=nbytes(x.value)) >= limit) alert('길이 '+limit+' 최대값 : ' + y); 
    if((y=nbytes(x.value)) == limit) alert('길이 '+limit+' 초과 : ' + y); 
} 


//2016-09-07 callback 회사명 리스트 selectBox 생성
function setCompanySelect(targetId, defaultKey, data){
	var opt = "<option value=\"\">전체</option>";
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
			alert("에러가 발생하였습니다.\n상태코드 : " + status + " : " + errorThrown);
		}
	});
}