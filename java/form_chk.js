
// 빈 문자열인지 검사

function result_form_chk()
{

	if(document.form1.acpt_user.value =="") {
		alert('사용자를 입력하세요');
		form1.acpt_user.focus();
		return false;}
	if(document.form1.addr.value =="") {
		alert('나머지 주소를 입력하세요');
		form1.addr.focus();
		return false;}
	if(document.form1.as_memo.value =="") {
		alert('장애내용을 입력하세요');
		form1.as_memo.focus();
		return false;}
	if(document.form1.request_date.value =="") {
		alert('요청일을 입력하세요');
		form1.request_date.focus();
		return false;}
	if(document.form1.visit_date.value =="") {
		alert('방문예정일을 입력하세요');
		form1.visit_date.focus();
		return false;}
	if(document.form1.request_date.value =="") {
		alert('요청일을 입력하세요');
		form1.request_date.focus();
		return false;}
	if(document.form1.as_process.value =="") {
		alert('처리현황을 입력하세요');
		form1.as_process.focus();
		return false;}
	if(document.form1.as_type.value =="") {
		alert('A/S 유형을 입력하세요');
		form1.as_type.focus();
		return false;}
	if(document.form1.as_process.value =="입고" || document.form1.as_process.value =="연기") 
		if(document.form1.into_reason.value =="") {
			alert('입고 및 연기 사유를 입력하세요');
			form1.into_reason.focus();
		return false;}

	{
	a=confirm('입력하시겠습니까?')
	if (a==true) {
		return true;
	}
	return false;
	}
}


function as_form_chk()
{
	if(document.form1.tel_no.value =="") {
		alert('전화번호를 입력하세요');
		form1.tel_no.focus();
		return false;}
	if(document.form1.company.value =="") {
		alert('회사를 입력하세요');
		form1.company.focus();
		return false;}	
	if(document.form1.dept.value =="") {
		alert('조직명을 입력하세요');
		form1.dept.focus();
		return false;}
	if(document.form1.sido.value =="") {
		alert('지역 검색을 하세요');
		form1.area_view.focus();
		return false;}
	if(document.form1.addr.value =="") {
		alert('나머지 주소를 입력하세요');
		form1.addr.focus();
		return false;}
	if(document.form1.acpt_user.value =="") {
		alert('사용자를 입력하세요');
		form1.acpt_user.focus();
		return false;}
	if(document.form1.as_memo.value =="") {
		alert('장애내용을 입력하세요');
		form1.as_memo.focus();
		return false;}
	if(document.form1.request_date.value =="") {
		alert('요청일을 입력하세요');
		form1.request_date.focus();
		return false;}
	if(document.form1.visit_date.value =="") {
		alert('방문예정일을 입력하세요');
		form1.visit_date.focus();
		return false;}

	{
	a=confirm('입력하시겠습니까?')
	if (a==true) {
		return true;
	}
	return false;
	}
}


function user_form_chk()
{
	if(document.form1.id.value =="") {
		alert('아이디를 입력하세요');
		form1.id.focus();
		return false;}
	if(document.form1.pass.value =="") {
		alert('패스워드를 입력하세요');
		form1.pass.focus();
		return false;}	
	if(document.form1.pass.value != document.form1.re_pass.value) {
		alert('패스워드를 확인하세요');
		form1.re_pass.focus();
		return false;}		
	if(document.form1.user_name.value =="") {
		alert('사용자를 입력하세요');
		form1.user_name.focus();
		return false;}
	if(document.form1.hp.value =="") {
		alert('핸드폰을 입력하세요');
		form1.hp.focus();
		return false;}
	if(document.form1.email.value =="") {
		alert('이메일을 입력하세요');
		form1.email.focus();
		return false;}
	{
	a=confirm('입력하시겠습니까?')
	if (a==true) {
		return true;
	}
	return false;
	}
}
function user_mod_chk()
{
	if(document.form1.pass.value != document.form1.re_pass.value) {
		alert('비밀번호를 확인하세요');
		form1.re_pass.focus();
		return false;}		
	if(document.form1.mod_pass.value != document.form1.mod_re_pass.value) {
		alert('변경 비밀번호가 서로 다릅니다.');
		form1.user_name.focus();
		return false;}
	if(document.form1.hp.value =="") {
		alert('핸드폰을 입력하세요');
		form1.hp.focus();
		return false;}
	if(document.form1.email.value =="") {
		alert('이메일을 입력하세요');
		form1.email.focus();
		return false;}
	{
	a=confirm('수정하시겠습니까?')
	if (a==true) {
		return true;
	}
	return false;
	}
}

function juso_form_chk()
{
	if(document.form1.tel_no.value =="") {
		alert('전화번호를 입력하세요');
		form1.tel_no.focus();
		return false;}
	if(document.form1.company.value =="") {
		alert('회사를 입력하세요');
		form1.company.focus();
		return false;}	
	if(document.form1.dept.value =="") {
		alert('부서를 입력하세요');
		form1.dept.focus();
		return false;}
	if(document.form1.sido.value =="") {
		alert('지역 검색을 하세요');
		form1.area_view.focus();
		return false;}
	if(document.form1.addr.value =="") {
		alert('나머지 주소를 입력하세요');
		form1.addr.focus();
		return false;}
	{
	a=confirm('입력하시겠습니까?')
	if (a==true) {
		return true;
	}
	return false;
	}
}

function board_form_chk()
{
	if(document.form1.title.value =="") {
		alert('제목을를 입력하세요');
		form1.title.focus();
		return false;}
	if(document.form1.fbody.value =="") {
		alert('내용를 입력하세요');
		form1.fbody.focus();
		return false;}	
	if(document.form1.pass.value =="") {
		alert('패스워드를 입력하세요');
		form1.pass.focus();
		return false;}

	{
	a=confirm('입력하시겠습니까?')
	if (a==true) {
		return true;
	}
	return false;
	}
}

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

function formemailcheck1 (arg) {
	var			what, i, j ;
	var			emailexp, arg ;

//	emailexp = new RegExp ("^(..+)@(..+)\\.(..+)$") ;
	emailexp = new RegExp ("^[A-Za-z0-9-_\\.]{2,}@[A-Za-z0-9-_\\.]{2,}\\.[A-Za-z0-9-_]{2,}$") ;

	what = arg ;
	
	if ( what.onlyemail != null ) {
		if ( what.type == "text" ) {
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

function formnullcheck1 (arg) {
	var			what, i, arg ;

	what = arg ;
	if ( what.notnull != null ) {
		if ( what.type == "text" || what.type == "password" 
			|| what.type == "textarea" || what.type == "file" )
			if ( isempty (what.value) ) {
				alert (what.errname + "을(를) 입력하세요") ;
				what.focus () ;
				return false ;
			}				
	}
	return true ;
}	

// text, password, textarea, file 형식 입력상자의 최소길이 및 최대길이 검사
// 최소길이 및 최대길이 테스트 통과시 true, 실패시 false 를 리턴
// maxlength, minlength
function formlencheck (form) {
	var			what, i ;
	var			str_len, errname ;

	for (i = 0; i < form.length; i++) {
		what = form.elements[i] ;

		if ( what.type == "text" || what.type == "password" 
			|| what.type == "textarea" || what.type == "file" ) {

			if ( what.value == "" )
				continue ;

			if ( what.type == "file" ) {
				var			rslash_pos ;
				var			fname ;

				rslash_pos = what.value.lastIndexOf ("\\", what.value.length) ;
				if ( rslash_pos == -1 ) {
					alert (what.errname + "의 파일이름이 "
							+ "잘못되었습니다\n"
							+ "다시한번 확인하세요") ;
					what.focus () ;
					return false ;
				}

				fname = what.value.substr (rslash_pos + 1) ;
				str_len = h_stringlength (fname) ;
				errname = what.errname + " 파일이름" ;
			} else {
				str_len = h_stringlength (what.value) ;
				errname = what.errname ;
			}

//			if ( str_len == -1 ) {
//				alert ("잘못된 문자열입니다.") ;
//				what.focus () ;
//				return false ;
//			}

			// 최소 길이 검사
			if ( what.minlength != null ) {
				if ( str_len < what.minlength ) {
					alert (errname + "은(는) 영문(숫자) "
						+ what.minlength + "자, 한글 "
						+ Math.ceil (what.minlength / 2)
						+ "자 이상 입력해야 합니다") ;
					what.focus () ;
					return false ;
				}
			}

			// 최대 길이 검사
			if ( str_len > what.maxLength ) {
				alert (errname + "은(는) 영문(숫자) "
					+ what.maxLength + "자, 한글 "
					+ parseInt (what.maxLength / 2) 
					+ "자 까지 입력할 수 있습니다") ;
				what.focus () ;
				return false ;						
			}
		}
	}
	return true ;			
}

function formlencheck1 (arg) {
	var			what, i ;
	var			str_len, errname, arg ;

	what = arg ;
	
	if ( what.type == "text" || what.type == "password" 
		|| what.type == "textarea" || what.type == "file" ) {
	
		if ( what.type == "file" ) {
			var			rslash_pos ;
			var			fname ;
	
			rslash_pos = what.value.lastIndexOf ("\\", what.value.length) ;
			if ( rslash_pos == -1 ) {
				alert (what.errname + "의 파일이름이 "
						+ "잘못되었습니다\n"
						+ "다시한번 확인하세요") ;
				what.focus () ;
				return false ;
			}
	
			fname = what.value.substr (rslash_pos + 1) ;
			str_len = h_stringlength (fname) ;
			errname = what.errname + " 파일이름" ;
		} else {
			str_len = h_stringlength (what.value) ;
			errname = what.errname ;
		}
	
		// 최소 길이 검사
		if ( what.minlength != null ) {
			if ( str_len < what.minlength ) {
				alert (errname + "은(는) 영문(숫자) "
					+ what.minlength + "자, 한글 "
					+ Math.ceil (what.minlength / 2)
					+ "자 이상 입력해야 합니다") ;
				what.focus () ;
				return false ;
			}
		}
	
		// 최대 길이 검사
		if ( str_len > what.maxLength ) {
			alert (errname + "은(는) 영문(숫자) "
				+ what.maxLength + "자, 한글 "
				+ parseInt (what.maxLength / 2) 
				+ "자 까지 입력할 수 있습니다") ;
			what.focus () ;
			return false ;						
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

function formnumcheck1 (form) {
	var			what, i, arg ;

	what = arg ;
	
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
	return true ;
}

function formnumcheck1 (arg) {
	var			what, i, arg ;

	what = arg ;
	
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

function formstrcheck1 (arg) {
	var			what, i, arg ;

	what = arg ;
	
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

function formengcheck1 (arg) {
	var			what, i, arg ;
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
	
	what = arg ;
	
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
	return true ;
}

// 영어인지 검사하고 submit
function formengcheck_and_submit (form) {
	if ( formengcheck (form) )
		form.submit () ;
}


// 모든 검사 수행
function formcheck (form) {
	if ( !formnullcheck (form) || !formlencheck (form) 
			|| !formnumcheck (form) || !formstrcheck (form) 
			|| !formengcheck (form) || !formemailcheck (form) 
			|| !formregnocheck (form) )
		return false ;

	return true ;
}

function formcheck1 (arg) {
	if ( !formnullcheck1 (arg) || !formlencheck1 (arg) 
			|| !formnumcheck1 (arg) || !formstrcheck1 (arg) 
			|| !formengcheck1 (arg) || !formemailcheck1 (arg) 
			|| !formregnocheck1 (arg) )
		return false ;

	return true ;
}

// 모든 검사 후 성공시 submit
function formcheck_and_submit (form) {
	if ( formcheck (form) )
		form.submit () ;
}

 <!-----달력 새창열기---------->
     function calendar_window(ref)
   { window.open(ref,"calendar",'width=140,height=120, toolbar=no,status=no,directories=no,menubar=no,scrollbars=yes,resizable=no'); }

//주민번호검사
function formregnocheck(form) {
	var			what, i, j ;
 	var			IDtot = 0
 	var			IDAdd = "234567892345"
 
	for (i = 0; i < form.length; i++) {
		what = form.elements[i] ;
		if ( what.type == "text" ) {
			if ( what.onlyregno != null ) {

		  	jslastchar =  parseInt(what.value.charAt(12))
  
		  	for (j = 0; j <= 11; j++) {                        
				 IDtot = IDtot + parseInt(what.value.charAt(j)) * parseInt(IDAdd.charAt(j))
		  	}
  
  			IDtot = 11 - (IDtot % 11)

				if (IDtot > 9) {
			    IDtot = IDtot.toString()
				  IDtot = IDtot.charAt(1)
				}

		 	 if (jslastchar !=  IDtot) {
			    alert(what.errname + "가 올바르지 않습니다");
					what.focus();
		  		return false;
		  	}
			}
		}
	}
	return true ;
}

function formregnocheck1(arg) {
	var			what, i, j, arg ;
 	var			IDtot = 0
 	var			IDAdd = "234567892345"
 
	what = arg ;
	if ( what.type == "text" || what.type == "hidden" ) {
		if ( what.onlyregno != null ) {
	
	  	jslastchar =  parseInt(what.value.charAt(12))
	
	  	for (j = 0; j <= 11; j++) {                        
			 IDtot = IDtot + parseInt(what.value.charAt(j)) * parseInt(IDAdd.charAt(j))
	  	}
	
				IDtot = 11 - (IDtot % 11)
	
			if (IDtot > 9) {
		    IDtot = IDtot.toString()
			  IDtot = IDtot.charAt(1)
			}
	
	 	 if (jslastchar !=  IDtot) {
		    alert(what.errname + "가 올바르지 않습니다");
				what.focus();
	  		return false;
	  	}
		}
	}
	return true ;
}

//주민등록번호 체크 '최삼락
function reg_no_check(reg_no1, reg_no2)
{
    var codesum = 0;
    var coderet = 0;
    var month, day;
    var sex_code;
    //자리수 확인
    if ( reg_no1.length != 6 || reg_no2.length != 7 )
    {
		return false;
    }
    //숫자인지
    if ( isNaN( reg_no1 ) || isNaN( reg_no2 ) )
    {
		return false;
    }
    
    sex_code = reg_no2.substring(0,1)
    if ( !(sex_code == '1' || sex_code == '2' || sex_code == '3' || sex_code == '4') )
    {
		return false;
    }
    
    if (sex_code == "3" || sex_code == "4") {
		year = "20" + reg_no1.substring(0,2);
    }
    else {
		year = "19" + reg_no1.substring(0,2);
    }
    
    year = Number(year) ;
    month = Number(reg_no1.substring(2,4) );
    day = Number(reg_no1.substring(4,6));
    
    if ( month > 12){
		return false;
	}
	
	if ( month == 1 ||  month == 3 ||  month == 5 ||  month == 7 ||  month == 8 ||  month == 10 ||  month == 12 ){
		if ( day > 31 ){
			return false;
		}
	}
	if ( month == 4 ||  month == 6 || month == 9 ||  month == 11){
		if ( day > 30 ){
			return false;
		}
	}
	if ( month == 2 && year % 4 == 0 &&  day > 29){
				return false;
	}
	if ( month == 2 && year % 4 != 0 &&  day > 28){
				return false;
	}
	
    codesum = ( eval( reg_no1.substring( 0, 1 ) ) * 2 )
            + ( eval( reg_no1.substring( 1, 2 ) ) * 3 )
            + ( eval( reg_no1.substring( 2, 3 ) ) * 4 )
            + ( eval( reg_no1.substring( 3, 4 ) ) * 5 )
            + ( eval( reg_no1.substring( 4, 5 ) ) * 6 )
            + ( eval( reg_no1.substring( 5, 6 ) ) * 7 )
            + ( eval( reg_no2.substring( 0, 1 ) ) * 8 )
            + ( eval( reg_no2.substring( 1, 2 ) ) * 9 )
            + ( eval( reg_no2.substring( 2, 3 ) ) * 2 )
            + ( eval( reg_no2.substring( 3, 4 ) ) * 3 )
            + ( eval( reg_no2.substring( 4, 5 ) ) * 4 )
            + ( eval( reg_no2.substring( 5, 6 ) ) * 5 );
    coderet = 11 - ( eval( codesum % 11 ) );
    if ( coderet >= 10 )
        coderet = coderet - 10;
    if ( coderet >= 10 )
        coderet = coderet - 10;
    if ( eval( reg_no2.substring( 6, 7 ) ) != coderet )
    {
		return false;
    }
    else
		return true;
}

//외국인 주민번호 체크
function fgn_no_chksum(reg_no) {
    var sum = 0;
    var odd = 0;
    buf = new Array(13);
    for ( var i = 0; i < 13; i++) buf[i] = parseInt(reg_no.charAt(i));
	
    odd = buf[7]*10 + buf[8];
    
    if (odd%2 != 0) {
      return false;
    }

    if ( (buf[11] != 6) && (buf[11] != 7) && (buf[11] != 8) && (buf[11] != 9) ) {
      return false;
    }
    	
    multipliers = [2,3,4,5,6,7,8,9,2,3,4,5];
    for (i = 0, sum = 0; i < 12; i++) sum += (buf[i] *= multipliers[i]);


    sum=11-(sum%11);
    
    if (sum>=10) sum-=10;

    sum += 2;

    if (sum>=10) sum-=10;

    if ( sum != buf[12]) {
        return false;
    }
    else {
        return true;
    }

}