<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")

as_date = request("as_date")
as_seq = request("as_seq")
as_car_no = request("as_car_no")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

car_no = as_car_no

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

as_cause = ""
as_solution = ""
as_amount = 0
as_amount_sign = "현금"
as_repair_pre_yn = ""
as_car_name = ""
as_owner_emp_no = ""
as_owner_emp_name = ""
as_use_org_name = ""

view_condi = ""

title_line = "차량 AS등록"

if u_type = "U" then

	sql = "select * from car_as where as_car_no = '" + car_no + "' and as_date = '" + as_date + "' and as_seq = '" + as_seq + "'"
	set rs = dbconn.execute(sql)

    as_car_no = rs("as_car_no")
	as_date = rs("as_date")
	as_seq = rs("as_seq")
	
	as_cause = rs("as_cause")
    as_solution = rs("as_solution")
    as_amount = rs("as_amount")
	as_amount_sign = rs("as_amount_sign")
	as_repair_pre_yn = rs("as_repair_pre_yn")
    as_car_name = rs("as_car_name")
    as_owner_emp_no = rs("as_owner_emp_no")
    as_owner_emp_name = rs("as_owner_emp_name")
    as_use_org_name = rs("as_use_org_name")

	rs.close()

	title_line = "차량 AS변경"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
        <title>인사관리 시스템</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=as_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=last_check_date%>" );
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
				if(document.frm.as_date.value =="" ) {
					alert('A/S일자를 입력하세요');
					frm.as_date.focus();
					return false;}
				if(document.frm.as_cause.value =="") {
					alert('증상을 입력하세요');
					frm.as_cause.focus();
					return false;}
				if(document.frm.as_solution.value =="") {
					alert('수리내용을 입력하세요');
					frm.as_solution.focus();
					return false;}
				if(document.frm.as_amount.value =="") {
					alert('수리비용을 입력하세요');
					frm.as_amount.focus();
					return false;}
				
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.as_repair_pre_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("비용(선)사용 구분을 선택하세요");
					return false;
				}	
				
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function num_chk(txtObj){
				as_amt = parseInt(document.frm.as_amount.value.replace(/,/g,""));		
				as_amt = String(as_amt);
				num_len = as_amt.length;
				sil_len = num_len;
				as_amt = String(as_amt);
				if (as_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) as_amt = as_amt.substr(0,num_len -3) + "," + as_amt.substr(num_len -3,3);
				if (sil_len > 6) as_amt = as_amt.substr(0,num_len -6) + "," + as_amt.substr(num_len -6,3) + "," + as_amt.substr(num_len -2,3);
				document.frm.as_amount.value = as_amt; 
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
				<form action="insa_car_as_save.asp" method="post" name="frm">
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
								<td class="left" bgcolor="#FFFFE6">
                                <input name="car_no" type="text" value="<%=as_car_no%>" style="width:150px" readonly="true">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_car_info_select.asp?gubun=<%="caras"%>','carinfoselect','scrollbars=yes,width=600,height=400')">차량찾기</a>
                                </td>
								<th style="background:#FFFFE6">차종</th>
								<td class="left" bgcolor="#FFFFE6">
                                <input name="car_name" type="text" value="<%=car_name%>" style="width:150px" readonly="true"></td>
							</tr>
                           	<tr>
								<th class="first" style="background:#FFFFE6">차량연식</th>
								<td class="left" bgcolor="#FFFFE6">
                                <input name="car_year" type="text" value="<%=car_year%>" style="width:70px" readonly="true">&nbsp;</td>
                                <th style="background:#FFFFE6">차량등록일</th>
								<td class="left" bgcolor="#FFFFE6">
                                <input name="car_reg_date" type="text" value="<%=car_reg_date%>" style="width:70px" readonly="true">&nbsp;</td>
							</tr>   
                            <tr>
								<th class="first" style="background:#FFFFE6">현 운행자</th>
								<td colspan="3" class="left" bgcolor="#FFFFE6">
                                <input name="owner_emp_name" type="text" value="<%=owner_emp_name%>" style="width:70px" readonly="true">
                          		-
                                <input name="owner_emp_no" type="text" value="<%=owner_emp_no%>" style="width:70px" readonly="true">
                                -
                                <input name="car_use_dept" type="text" value="<%=car_use_dept%>" style="width:90px" readonly="true">
                                </td>
							</tr>             
                            <tr>
								<th class="first">AS발생일</th>
								<td colspan="3" class="left"><input name="as_date" type="text" value="<%=as_date%>" style="width:70px" id="datepicker">
                                </td>
							</tr>  
							<tr>
								<th class="first">증상</th>
								<td colspan="3" class="left">
                                <input name="as_cause" type="text" value="<%=as_cause%>" style="width:600px" onKeyUp="checklength(this,50)">
                                </td>
							</tr>
                            <tr>
								<th class="first">수리내용</th>
								<td colspan="3" class="left">
                                <input name="as_solution" type="text" value="<%=as_solution%>" style="width:600px" onKeyUp="checklength(this,100)">
                                </td>
							</tr>
                            <tr>
								<th class="first">수리비용</th>
                                <td class="left">
                                <input name="as_amount" type="text" id="as_amount" style="width:90px;text-align:right" value="<%=formatnumber(as_amount,0)%>" onKeyUp="num_chk(this);">
                                <input type="radio" name="as_amount_sign" value="현금" <% if as_amount_sign = "현금" then %>checked<% end if %> style="width:40px" id="Radio1">현금
                                <input type="radio" name="as_amount_sign" value="카드" <% if as_amount_sign = "카드" then %>checked<% end if %> style="width:40px" id="Radio2">카드
                                </td>
                                <th>비용(선)사용구분</th>
							    <td class="left"><input type="radio" name="as_repair_pre_yn" value="N" <% if as_repair_pre_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio">개인
                                  <input type="radio" name="as_repair_pre_yn" value="Y" <% if as_repair_pre_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio2">회사
                                </td>
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
                <input type="hidden" name="oil_kind" value="<%=oil_kind%>" ID="Hidden1">
                <input type="hidden" name="car_owner" value="<%=car_owner%>" ID="Hidden1">
                
                <input type="hidden" name="old_as_car_no" value="<%=as_car_no%>" ID="Hidden1">
                <input type="hidden" name="old_as_date" value="<%=as_date%>" ID="Hidden1">
                <input type="hidden" name="old_as_seq" value="<%=as_seq%>" ID="Hidden1">
                
			</form>
		</div>				
	</body>
</html>

