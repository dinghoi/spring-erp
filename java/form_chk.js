
// �� ���ڿ����� �˻�

function result_form_chk()
{

	if(document.form1.acpt_user.value =="") {
		alert('����ڸ� �Է��ϼ���');
		form1.acpt_user.focus();
		return false;}
	if(document.form1.addr.value =="") {
		alert('������ �ּҸ� �Է��ϼ���');
		form1.addr.focus();
		return false;}
	if(document.form1.as_memo.value =="") {
		alert('��ֳ����� �Է��ϼ���');
		form1.as_memo.focus();
		return false;}
	if(document.form1.request_date.value =="") {
		alert('��û���� �Է��ϼ���');
		form1.request_date.focus();
		return false;}
	if(document.form1.visit_date.value =="") {
		alert('�湮�������� �Է��ϼ���');
		form1.visit_date.focus();
		return false;}
	if(document.form1.request_date.value =="") {
		alert('��û���� �Է��ϼ���');
		form1.request_date.focus();
		return false;}
	if(document.form1.as_process.value =="") {
		alert('ó����Ȳ�� �Է��ϼ���');
		form1.as_process.focus();
		return false;}
	if(document.form1.as_type.value =="") {
		alert('A/S ������ �Է��ϼ���');
		form1.as_type.focus();
		return false;}
	if(document.form1.as_process.value =="�԰�" || document.form1.as_process.value =="����") 
		if(document.form1.into_reason.value =="") {
			alert('�԰� �� ���� ������ �Է��ϼ���');
			form1.into_reason.focus();
		return false;}

	{
	a=confirm('�Է��Ͻðڽ��ϱ�?')
	if (a==true) {
		return true;
	}
	return false;
	}
}


function as_form_chk()
{
	if(document.form1.tel_no.value =="") {
		alert('��ȭ��ȣ�� �Է��ϼ���');
		form1.tel_no.focus();
		return false;}
	if(document.form1.company.value =="") {
		alert('ȸ�縦 �Է��ϼ���');
		form1.company.focus();
		return false;}	
	if(document.form1.dept.value =="") {
		alert('�������� �Է��ϼ���');
		form1.dept.focus();
		return false;}
	if(document.form1.sido.value =="") {
		alert('���� �˻��� �ϼ���');
		form1.area_view.focus();
		return false;}
	if(document.form1.addr.value =="") {
		alert('������ �ּҸ� �Է��ϼ���');
		form1.addr.focus();
		return false;}
	if(document.form1.acpt_user.value =="") {
		alert('����ڸ� �Է��ϼ���');
		form1.acpt_user.focus();
		return false;}
	if(document.form1.as_memo.value =="") {
		alert('��ֳ����� �Է��ϼ���');
		form1.as_memo.focus();
		return false;}
	if(document.form1.request_date.value =="") {
		alert('��û���� �Է��ϼ���');
		form1.request_date.focus();
		return false;}
	if(document.form1.visit_date.value =="") {
		alert('�湮�������� �Է��ϼ���');
		form1.visit_date.focus();
		return false;}

	{
	a=confirm('�Է��Ͻðڽ��ϱ�?')
	if (a==true) {
		return true;
	}
	return false;
	}
}


function user_form_chk()
{
	if(document.form1.id.value =="") {
		alert('���̵� �Է��ϼ���');
		form1.id.focus();
		return false;}
	if(document.form1.pass.value =="") {
		alert('�н����带 �Է��ϼ���');
		form1.pass.focus();
		return false;}	
	if(document.form1.pass.value != document.form1.re_pass.value) {
		alert('�н����带 Ȯ���ϼ���');
		form1.re_pass.focus();
		return false;}		
	if(document.form1.user_name.value =="") {
		alert('����ڸ� �Է��ϼ���');
		form1.user_name.focus();
		return false;}
	if(document.form1.hp.value =="") {
		alert('�ڵ����� �Է��ϼ���');
		form1.hp.focus();
		return false;}
	if(document.form1.email.value =="") {
		alert('�̸����� �Է��ϼ���');
		form1.email.focus();
		return false;}
	{
	a=confirm('�Է��Ͻðڽ��ϱ�?')
	if (a==true) {
		return true;
	}
	return false;
	}
}
function user_mod_chk()
{
	if(document.form1.pass.value != document.form1.re_pass.value) {
		alert('��й�ȣ�� Ȯ���ϼ���');
		form1.re_pass.focus();
		return false;}		
	if(document.form1.mod_pass.value != document.form1.mod_re_pass.value) {
		alert('���� ��й�ȣ�� ���� �ٸ��ϴ�.');
		form1.user_name.focus();
		return false;}
	if(document.form1.hp.value =="") {
		alert('�ڵ����� �Է��ϼ���');
		form1.hp.focus();
		return false;}
	if(document.form1.email.value =="") {
		alert('�̸����� �Է��ϼ���');
		form1.email.focus();
		return false;}
	{
	a=confirm('�����Ͻðڽ��ϱ�?')
	if (a==true) {
		return true;
	}
	return false;
	}
}

function juso_form_chk()
{
	if(document.form1.tel_no.value =="") {
		alert('��ȭ��ȣ�� �Է��ϼ���');
		form1.tel_no.focus();
		return false;}
	if(document.form1.company.value =="") {
		alert('ȸ�縦 �Է��ϼ���');
		form1.company.focus();
		return false;}	
	if(document.form1.dept.value =="") {
		alert('�μ��� �Է��ϼ���');
		form1.dept.focus();
		return false;}
	if(document.form1.sido.value =="") {
		alert('���� �˻��� �ϼ���');
		form1.area_view.focus();
		return false;}
	if(document.form1.addr.value =="") {
		alert('������ �ּҸ� �Է��ϼ���');
		form1.addr.focus();
		return false;}
	{
	a=confirm('�Է��Ͻðڽ��ϱ�?')
	if (a==true) {
		return true;
	}
	return false;
	}
}

function board_form_chk()
{
	if(document.form1.title.value =="") {
		alert('�������� �Է��ϼ���');
		form1.title.focus();
		return false;}
	if(document.form1.fbody.value =="") {
		alert('���븦 �Է��ϼ���');
		form1.fbody.focus();
		return false;}	
	if(document.form1.pass.value =="") {
		alert('�н����带 �Է��ϼ���');
		form1.pass.focus();
		return false;}

	{
	a=confirm('�Է��Ͻðڽ��ϱ�?')
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

function formemailcheck1 (arg) {
	var			what, i, j ;
	var			emailexp, arg ;

//	emailexp = new RegExp ("^(..+)@(..+)\\.(..+)$") ;
	emailexp = new RegExp ("^[A-Za-z0-9-_\\.]{2,}@[A-Za-z0-9-_\\.]{2,}\\.[A-Za-z0-9-_]{2,}$") ;

	what = arg ;
	
	if ( what.onlyemail != null ) {
		if ( what.type == "text" ) {
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

function formnullcheck1 (arg) {
	var			what, i, arg ;

	what = arg ;
	if ( what.notnull != null ) {
		if ( what.type == "text" || what.type == "password" 
			|| what.type == "textarea" || what.type == "file" )
			if ( isempty (what.value) ) {
				alert (what.errname + "��(��) �Է��ϼ���") ;
				what.focus () ;
				return false ;
			}				
	}
	return true ;
}	

// text, password, textarea, file ���� �Է»����� �ּұ��� �� �ִ���� �˻�
// �ּұ��� �� �ִ���� �׽�Ʈ ����� true, ���н� false �� ����
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
					alert (what.errname + "�� �����̸��� "
							+ "�߸��Ǿ����ϴ�\n"
							+ "�ٽ��ѹ� Ȯ���ϼ���") ;
					what.focus () ;
					return false ;
				}

				fname = what.value.substr (rslash_pos + 1) ;
				str_len = h_stringlength (fname) ;
				errname = what.errname + " �����̸�" ;
			} else {
				str_len = h_stringlength (what.value) ;
				errname = what.errname ;
			}

//			if ( str_len == -1 ) {
//				alert ("�߸��� ���ڿ��Դϴ�.") ;
//				what.focus () ;
//				return false ;
//			}

			// �ּ� ���� �˻�
			if ( what.minlength != null ) {
				if ( str_len < what.minlength ) {
					alert (errname + "��(��) ����(����) "
						+ what.minlength + "��, �ѱ� "
						+ Math.ceil (what.minlength / 2)
						+ "�� �̻� �Է��ؾ� �մϴ�") ;
					what.focus () ;
					return false ;
				}
			}

			// �ִ� ���� �˻�
			if ( str_len > what.maxLength ) {
				alert (errname + "��(��) ����(����) "
					+ what.maxLength + "��, �ѱ� "
					+ parseInt (what.maxLength / 2) 
					+ "�� ���� �Է��� �� �ֽ��ϴ�") ;
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
				alert (what.errname + "�� �����̸��� "
						+ "�߸��Ǿ����ϴ�\n"
						+ "�ٽ��ѹ� Ȯ���ϼ���") ;
				what.focus () ;
				return false ;
			}
	
			fname = what.value.substr (rslash_pos + 1) ;
			str_len = h_stringlength (fname) ;
			errname = what.errname + " �����̸�" ;
		} else {
			str_len = h_stringlength (what.value) ;
			errname = what.errname ;
		}
	
		// �ּ� ���� �˻�
		if ( what.minlength != null ) {
			if ( str_len < what.minlength ) {
				alert (errname + "��(��) ����(����) "
					+ what.minlength + "��, �ѱ� "
					+ Math.ceil (what.minlength / 2)
					+ "�� �̻� �Է��ؾ� �մϴ�") ;
				what.focus () ;
				return false ;
			}
		}
	
		// �ִ� ���� �˻�
		if ( str_len > what.maxLength ) {
			alert (errname + "��(��) ����(����) "
				+ what.maxLength + "��, �ѱ� "
				+ parseInt (what.maxLength / 2) 
				+ "�� ���� �Է��� �� �ֽ��ϴ�") ;
			what.focus () ;
			return false ;						
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

function formnumcheck1 (form) {
	var			what, i, arg ;

	what = arg ;
	
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
	return true ;
}

function formnumcheck1 (arg) {
	var			what, i, arg ;

	what = arg ;
	
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

function formstrcheck1 (arg) {
	var			what, i, arg ;

	what = arg ;
	
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
	return true ;
}

// �������� �˻��ϰ� submit
function formengcheck_and_submit (form) {
	if ( formengcheck (form) )
		form.submit () ;
}


// ��� �˻� ����
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

// ��� �˻� �� ������ submit
function formcheck_and_submit (form) {
	if ( formcheck (form) )
		form.submit () ;
}

 <!-----�޷� ��â����---------->
     function calendar_window(ref)
   { window.open(ref,"calendar",'width=140,height=120, toolbar=no,status=no,directories=no,menubar=no,scrollbars=yes,resizable=no'); }

//�ֹι�ȣ�˻�
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
			    alert(what.errname + "�� �ùٸ��� �ʽ��ϴ�");
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
		    alert(what.errname + "�� �ùٸ��� �ʽ��ϴ�");
				what.focus();
	  		return false;
	  	}
		}
	}
	return true ;
}

//�ֹε�Ϲ�ȣ üũ '�ֻ��
function reg_no_check(reg_no1, reg_no2)
{
    var codesum = 0;
    var coderet = 0;
    var month, day;
    var sex_code;
    //�ڸ��� Ȯ��
    if ( reg_no1.length != 6 || reg_no2.length != 7 )
    {
		return false;
    }
    //��������
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

//�ܱ��� �ֹι�ȣ üũ
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