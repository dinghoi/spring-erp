<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")
car_no = request("car_no")

ins_date = request("ins_date")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * FROM car_info where car_no = '"+car_no+"'"
Set Rs = DbConn.Execute(SQL)
if not Rs.EOF or not Rs.BOF then
     owner_emp_name = ""
	 owner_emp_no = rs("owner_emp_no")
	 if rs("owner_emp_name") = "" or isnull(rs("owner_emp_name")) then
	     Sql="select * from emp_master where emp_no = '"&owner_emp_no&"'"
	     Set rs_emp=DbConn.Execute(Sql)
				 owner_emp_name = rs_emp("emp_name")
	   else 
			     owner_emp_name = rs("owner_emp_name")
	 end if
	 if rs("last_check_date") = "1900-01-01"  then
	         last_check_date = ""
	    else 
		     last_check_date = rs("last_check_date")
	 end if
	 if rs("end_date") = "1900-01-01" then
	         end_date = ""
	    else 
		     end_date = rs("end_date")
	 end if
	 if rs("car_year") = "1900-01-01" then
	         car_year = ""
	    else 
	         car_year = rs("car_year")
	 end if
	 car_name = rs("car_name")
     car_reg_date = rs("car_reg_date")
     car_use_dept = rs("car_use_dept")
     oil_kind = rs("oil_kind")  
     car_owner = rs("car_owner") 
end if
rs.close()

ins_amount = 0
ins_company = ""
ins_last_date = ""
ins_man1 = ""
ins_man2 = ""
ins_object = ""
ins_self = ""
ins_injury = ""
ins_self_car = ""
ins_age = ""
ins_comment = ""
ins_scramble = ""

view_condi = ""

title_line = "차량 보험등록"

if u_type = "U" then

	sql = "select * from car_insurance where ins_car_no = '" + car_no + "' and ins_date = '" + ins_date + "'"
	set rs = dbconn.execute(sql)

    ins_car_no = rs("ins_car_no")
	ins_date = rs("ins_date")
	
	ins_old_date = rs("ins_date")
	
    ins_amount = rs("ins_amount")
    ins_company = rs("ins_company")
    ins_last_date = rs("ins_last_date")
    ins_man1 = rs("ins_man1")
    ins_man2 = rs("ins_man2")
    ins_object = rs("ins_object")
    ins_self = rs("ins_self")
    ins_injury = rs("ins_injury")
    ins_self_car = rs("ins_self_car")
    ins_age = rs("ins_age")
    ins_comment = rs("ins_comment")
    ins_scramble = rs("ins_scramble")
	ins_contract_yn = rs("ins_contract_yn")
	if ins_contract_yn = "N" then 
	       ins_comment = "필요시 제안사에서 운영" 
	   else 
	       ins_comment = ""
    end if
	rs.close()

	title_line = "차량 보험변경"  
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=ins_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=ins_last_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=end_date%>" );
			});	  
			$(function() {    $( "#datepicker3" ).datepicker();
												$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker3" ).datepicker("setDate", "<%=car_year%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.ins_date.value =="") {
					alert('보험가입일을 선택하세요');
					frm.ins_date.focus();
					return false;}		
				if(document.frm.ins_company.value =="" ) {
					alert('보험회사를 입력하세요');
					frm.ins_company.focus();
					return false;}
				if(document.frm.ins_amount.value =="") {
					alert('보험가입액을 입력하세요');
					frm.ins_amount.focus();
					return false;}
				if(document.frm.ins_last_date.value =="") {
					alert('보험만기일을 선택하세요');
					frm.ins_last_date.focus();
					return false;}		
				if(document.frm.ins_date.value > document.frm.ins_last_date.value) {
						alert('보험만기일이 보험가입일보다 빠릅니다');
						frm.ins_last_date.focus();
						return false;}
				if(document.frm.ins_man1.value =="") {
					alert('대인1 계약사항을 입력하세요');
					frm.ins_man1.focus();
					return false;}			
				if(document.frm.ins_man2.value =="") {
					alert('대인2 계약사항을 입력하세요');
					frm.ins_man2.focus();
					return false;}			
				if(document.frm.ins_object.value =="" ) {
					alert('대물 계약사항을 입력하세요');
					frm.ins_object.focus();
					return false;}
				if(document.frm.ins_self.value =="" ) {
					alert('자기 계약사항을 입력하세요');
					frm.ins_self.focus();
					return false;}
				if(document.frm.ins_injury.value =="" ) {
					alert('무상해 계약사항을 입력하세요');
					frm.ins_injury.focus();
					return false;}
				if(document.frm.ins_self_car.value =="" ) {
					alert('자차 계약사항을 입력하세요');
					frm.ins_self_car.focus();
					return false;}
			
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function num_chk(txtObj){
				ins_amt = parseInt(document.frm.ins_amount.value.replace(/,/g,""));		
				ins_amt = String(ins_amt);
				num_len = ins_amt.length;
				sil_len = num_len;
				ins_amt = String(ins_amt);
				if (ins_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) ins_amt = ins_amt.substr(0,num_len -3) + "," + ins_amt.substr(num_len -3,3);
				if (sil_len > 6) ins_amt = ins_amt.substr(0,num_len -6) + "," + ins_amt.substr(num_len -6,3) + "," + ins_amt.substr(num_len -2,3);

				document.frm.ins_amount.value = ins_amt; 

				if (txtObj.value.length >= 2) {
					if (txtObj.value.substr(0,1) == "0"){
						txtObj.value=txtObj.value.substr(1,1);
					}
				}
				if (txtObj.value.length<5) {
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
			
			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U') 
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_insurance_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first" style="background:#FFFFE6">차량번호</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_no%>&nbsp;
                                <input name="car_no" type="hidden" value="<%=car_no%>" style="width:150px" readonly="true"></td>
								<th style="background:#FFFFE6">차종</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_name%>&nbsp;
                                <input name="car_name" type="hidden" value="<%=car_name%>" style="width:150px" readonly="true"></td>
							</tr>
                           	<tr>
								<th class="first" style="background:#FFFFE6">차량연식</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_year%>&nbsp;
                                <input name="car_year" type="hidden" value="<%=car_year%>" style="width:70px" readonly="true"></td>
                                <th style="background:#FFFFE6">차량등록일</th>
								<td class="left" bgcolor="#FFFFE6"><%=car_reg_date%>&nbsp;
                                <input name="car_reg_date" type="hidden" value="<%=car_reg_date%>" style="width:70px" readonly="true"></td>
							</tr>   
                            <tr>
								<th class="first">보험가입일</th>
								<td class="left"><input name="ins_date" type="text" value="<%=ins_date%>" style="width:70px" id="datepicker"></td>
                                <th class="first">보험만기일</th>
								<td class="left"><input name="ins_last_date" type="text" value="<%=ins_last_date%>" style="width:70px" id="datepicker1"></td>
							</tr>            
							<tr>
								<th class="first">보험회사</th>
								<td class="left">
                                <input name="ins_company" type="text" value="<%=ins_company%>" style="width:150px" onKeyUp="checklength(this,30)">
                                </td>
								<th>보험가입액</th>
								<td class="left">
                                <input name="ins_amount" type="text" id="ins_amount" style="width:80px;text-align:right" value="<%=formatnumber(ins_amount,0)%>" onKeyUp="num_chk(this);">
                                </td>
							</tr>
							<tr>
								<th class="first">대인1</th>
								<td class="left">
                                <input name="ins_man1" type="text" value="<%=ins_man1%>" style="width:150px" onKeyUp="checklength(this,30)"></td>
								<th>대인2</th>
                                <td class="left">
								<input name="ins_man2" type="text" value="<%=ins_man2%>" style="width:150px" onKeyUp="checklength(this,10)"></td>
							</tr>
                            <tr>
								<th class="first">대물</th>
								<td class="left">
                                <input name="ins_object" type="text" value="<%=ins_object%>" style="width:150px" onKeyUp="checklength(this,10)"></td>
								<th>자기</th>
                                <td class="left">
								<input name="ins_self" type="text" value="<%=ins_self%>" style="width:150px" onKeyUp="checklength(this,10)"></td>
							</tr>
                            <tr>
								<th class="first">무상해</th>
								<td class="left">
                                <input name="ins_injury" type="text" value="<%=ins_injury%>" style="width:150px" onKeyUp="checklength(this,10)"></td>
								<th>자차</th>
                                <td class="left">
								<input name="ins_self_car" type="text" value="<%=ins_self_car%>" style="width:150px" onKeyUp="checklength(this,10)"></td>
							</tr>
                            <tr>
								<th class="first">연령</th>
								<td class="left">
                                <input name="ins_age" type="text" value="<%=ins_age%>" style="width:150px" onKeyUp="checklength(this,10)"></td>
								<th>긴급출동</th>
                                <td class="left">
								<input name="ins_scramble" type="text" value="<%=ins_scramble%>" style="width:150px" onKeyUp="checklength(this,10)"></td>
							</tr>
                             <tr>
								<th class="first">계약내용<br>포함유무</th>
                                <td class="left">
								<input type="radio" name="ins_contract_yn" value="Y" <% if ins_contract_yn = "Y" then %>checked<% end if %>>포함 
              					<input name="ins_contract_yn" type="radio" value="N" <% if ins_contract_yn = "N" then %>checked<% end if %>>미포함
                                </td>
                                <th>비고</th>
								<td class="left">
                                <input name="ins_comment" type="text" value="<%=ins_comment%>" style="width:200px" readonly="true"></td>
							</tr> 
                      </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="ins_old_date" value="<%=ins_old_date%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

