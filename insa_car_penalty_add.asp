<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")
car_no = request("car_no")

pe_date = request("pe_date")
pe_seq = request("pe_seq")

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

pe_comment = ""
pe_place = ""
pe_amount = 0
pe_in_date = ""
pe_in_amt = 0
pe_default = ""
pe_notice_date = ""
pe_notice = ""
pe_bigo = ""

view_condi = ""

title_line = "차량 과태료 등록"

if u_type = "U" then

	sql = "select * from car_penalty where pe_car_no = '" + car_no + "' and pe_date = '" + pe_date + "' and pe_seq = '" + pe_seq + "'"
	set rs = dbconn.execute(sql)

    pe_comment = rs("pe_comment")
    pe_place = rs("pe_place")
    pe_amount = rs("pe_amount")
    pe_in_date = rs("pe_in_date")
	pe_in_amt = rs("pe_in_amt")
    pe_default = rs("pe_default")
    pe_notice_date = rs("pe_notice_date")
    pe_notice = rs("pe_notice")
    pe_bigo = rs("pe_bigo")
	if rs("pe_in_date") = "1900-01-01"  then
	       pe_in_date = ""
	   else 
	       pe_in_date = rs("pe_in_date")
	end if
	if rs("pe_notice_date") = "1900-01-01" then
	       pe_notice_date = ""
	   else 
	       pe_notice_date = rs("pe_notice_date")
	end if
	rs.close()

	title_line = "차량 과태료 변경"
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
												$( "#datepicker" ).datepicker("setDate", "<%=pe_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=pe_in_date%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
												$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker2" ).datepicker("setDate", "<%=pe_notice_date%>" );
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
				if(document.frm.pe_date.value =="" ) {
					alert('위반일자를 입력하세요');
					frm.pe_date.focus();
					return false;}
				if(document.frm.pe_comment.value =="") {
					alert('위반내용을 입력하세요');
					frm.pe_comment.focus();
					return false;}
				if(document.frm.pe_place.value =="") {
					alert('위반장소을 입력하세요');
					frm.pe_place.focus();
					return false;}
				if(document.frm.pe_amount.value =="") {
					alert('과태료를 입력하세요');
					frm.pe_amount.focus();
					return false;}
				if(document.frm.pe_in_date.value != "") {
					if(document.frm.pe_in_amt.value == 0) {
							alert('납입금액을 입력하세요');
							frm.pe_in_amt.focus();
							return false;}}
							
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function num_chk(txtObj){
				pe_amt = parseInt(document.frm.pe_amount.value.replace(/,/g,""));		
				in_amt = parseInt(document.frm.pe_in_amt.value.replace(/,/g,""));		
				
				pe_amt = String(pe_amt);
				num_len = pe_amt.length;
				sil_len = num_len;
				pe_amt = String(pe_amt);
				if (pe_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) pe_amt = pe_amt.substr(0,num_len -3) + "," + pe_amt.substr(num_len -3,3);
				if (sil_len > 6) pe_amt = pe_amt.substr(0,num_len -6) + "," + pe_amt.substr(num_len -6,3) + "," + pe_amt.substr(num_len -2,3);
				document.frm.pe_amount.value = pe_amt; 
				
				in_amt = String(in_amt);
				num_len = in_amt.length;
				sil_len = num_len;
				in_amt = String(in_amt);
				if (in_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) in_amt = in_amt.substr(0,num_len -3) + "," + in_amt.substr(num_len -3,3);
				if (sil_len > 6) in_amt = in_amt.substr(0,num_len -6) + "," + in_amt.substr(num_len -6,3) + "," + in_amt.substr(num_len -2,3);
				document.frm.pe_in_amt.value = in_amt; 
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
				<form action="insa_car_penalty_save.asp" method="post" name="frm">
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
                                <input name="car_no" type="text" value="<%=car_no%>" style="width:150px" readonly="true">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_car_info_select.asp?gubun=<%="carpt"%>','carinfoselect','scrollbars=yes,width=600,height=400')">차량찾기</a>
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
                                -
                                <input name="car_owner" type="text" value="<%=car_owner%>" style="width:90px" readonly="true">
                                </td>
							</tr>             
                            <tr>
								<th class="first">위반일시</th>
								<td class="left"><input name="pe_date" type="text" value="<%=pe_date%>" style="width:70px" id="datepicker">
                                </td>
                                <th class="first">위반내용</th>
								<td class="left">
                                <select name="pe_comment" id="pe_comment" type="text" value="<%=pe_comment%>" style="width:150px">
                                  <option value="">선택</option>
								  <option value="속도위반" <%If pe_comment = "속도위반" then %>selected<% end if %>>속도위반</option>
								  <option value="신호위반" <%If pe_comment = "신호위반" then %>selected<% end if %>>신호위반</option>
								  <option value="주정차위반" <%If pe_comment = "주정차위반" then %>selected<% end if %>>주정차위반</option>
                                  <option value="통행료미납" <%If pe_comment = "통행료미납" then %>selected<% end if %>>통행료미납</option>
                                  <option value="주차료미납" <%If pe_comment = "주차료미납" then %>selected<% end if %>>주차료미납</option>
                                  <option value="기타" <%If pe_comment = "기타" then %>selected<% end if %>>기타</option>
							    </select>
                                </td>
							</tr>  
							<tr>
								<th class="first">위반장소</th>
								<td colspan="3" class="left">
                                <input name="pe_place" type="text" value="<%=pe_place%>" style="width:600px" onKeyUp="checklength(this,100)">
                                </td>
							</tr>
                            <tr>
								<th class="first">과태료</th>
                                <td colspan="3" class="left">
                                <input name="pe_amount" type="text" id="pe_amount" style="width:90px;text-align:right" value="<%=formatnumber(pe_amount,0)%>" onKeyUp="num_chk(this);">
                                </td>
							</tr>
                            <tr>
								<th class="first">납입일자</th>
								<td class="left"><input name="pe_in_date" type="text" value="<%=pe_in_date%>" style="width:70px" id="datepicker1">
                                </td>
                                <th>납입액</th>
								<td class="left">
                                <input name="pe_in_amt" type="text" id="pe_in_amt" style="width:90px;text-align:right" value="<%=formatnumber(pe_in_amt,0)%>" onKeyUp="num_chk(this);">
                                </td>
							</tr>
                            <tr>
								<th class="first">통보일자</th>
								<td class="left"><input name="pe_notice_date" type="text" value="<%=pe_notice_date%>" style="width:70px" id="datepicker2">
                                </td>
                                <th>통보방법</th>
								<td class="left">
                                <input name="pe_notice" type="text" value="<%=pe_notice%>" style="width:200px" onKeyUp="checklength(this,30)">
                                </td>
							</tr>
                            <tr>
								<th class="first">미납내용</th>
								<td class="left">
                                <input name="pe_default" type="text" value="<%=pe_default%>" style="width:200px" onKeyUp="checklength(this,30)">
                                </td>
                                <th>비고</th>
								<td class="left">
                                <input name="pe_bigo" type="text" value="<%=pe_bigo%>" style="width:200px" onKeyUp="checklength(this,30)">
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
                <input type="hidden" name="pe_seq" value="<%=pe_seq%>" ID="Hidden1">
 			</form>
		</div>				
	</body>
</html>

