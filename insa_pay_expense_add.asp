<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

u_type = request("u_type")
view_condi = request("view_condi")
ex_emp_no = request("ex_emp_no")
ex_emp_name = request("ex_emp_name")
ex_date = request("ex_date")
ex_deduct_id = request("ex_deduct_id")
ex_code = request("ex_code")
rever_yymm = request("rever_yyyymm")
ex_pay_date = request("ex_pay_date")

if ex_deduct_id = "G" then 
    etc_type = "60"
   else
    etc_type = "65"
end if

	ex_code_name = ""
	ex_tax_id = ""
	ex_company = ""
	ex_bonbu = ""
	ex_saupbu = ""
	ex_team = ""
	ex_reside_place = ""
	ex_reside_company = ""
	ex_org_name = ""
	ex_comment = ""

	ex_work_cnt = 0
	ex_amount = 0

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "SELECT * FROM emp_master where emp_no = '"&ex_emp_no&"'"
Set rs_emp = DbConn.Execute(SQL)
if not rs_emp.eof then
        emp_name = rs_emp("emp_name")
    	emp_first_date = rs_emp("emp_first_date")
		emp_in_date = rs_emp("emp_in_date")
		emp_type = rs_emp("emp_type")
		emp_grade = rs_emp("emp_grade")
		emp_position = rs_emp("emp_position")
		emp_company = rs_emp("emp_company")
		emp_bonbu = rs_emp("emp_bonbu")
		emp_saupbu = rs_emp("emp_saupbu")
		emp_team = rs_emp("emp_team")
		emp_org_code = rs_emp("emp_org_code")
		emp_org_name = rs_emp("emp_org_name")
		emp_reside_place = rs_emp("emp_reside_place")
		emp_reside_company = rs_emp("emp_reside_company")
   else
		emp_name = ""
		emp_first_date = ""
		emp_in_date = ""
		emp_type = ""
		emp_grade = ""
		emp_position = ""
		emp_company = ""
		emp_bonbu = ""
		emp_saupbu = ""
		emp_team = ""
		emp_org_code = ""
		emp_org_name = ""
		emp_reside_place = ""
		emp_reside_company = ""
end if
rs_emp.close()

title_line = "지급/공제 등록"

if u_type = "U" then

	sql = "select * from pay_expense where (ex_date = '"+ex_date+"') and (ex_emp_no = '"+ex_emp_no+"') and (ex_deduct_id = '"+ex_deduct_id+"') and (ex_code = '"+ex_code+"')"
	set rs = dbconn.execute(sql)

'테스트를하기위한..
    if not rs.eof then     
    ex_code_name = rs("ex_code_name")
	rever_yymm = rs("rever_yymm")
    rever_yymm = rs("rever_yymm")
	ex_tax_id = rs("ex_tax_id")
	ex_emp_name = rs("ex_emp_name")
	ex_company = rs("ex_company")
	ex_bonbu = rs("ex_bonbu")
	ex_saupbu = rs("ex_saupbu")
	ex_team = rs("ex_team")
	ex_reside_place = rs("ex_reside_place")
	ex_reside_company = rs("ex_reside_company")
	ex_org_name = rs("ex_org_name")
	ex_comment = rs("ex_comment")
	
	ex_amount = cint(rs("ex_amount"))
	ex_work_cnt = cint(rs("ex_work_cnt"))

	rs.close()
		
	end if

	title_line = "지급/공제 변경"
end if

pay_curr_amt = pmg_give_tot - de_deduct_tot

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
												$( "#datepicker" ).datepicker("setDate", "<%=ex_pay_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=ex_date%>" );
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
				if(document.frm.ex_emp_no.value =="" ) {
					alert('사번을 입력하세요');
					frm.ex_emp_no.focus();
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
				ex_amt = parseInt(document.frm.ex_amount.value.replace(/,/g,""));		
				ex_amt = String(ex_amt);
				num_len = ex_amt.length;
				sil_len = num_len;
				ex_amt = String(ex_amt);
				if (ex_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) ex_amt = ex_amt.substr(0,num_len -3) + "," + ex_amt.substr(num_len -3,3);
				if (sil_len > 6) ex_amt = ex_amt.substr(0,num_len -6) + "," + ex_amt.substr(num_len -6,3) + "," + ex_amt.substr(num_len -2,3);
				document.frm.ex_amount.value = ex_amt; 
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
				<form action="insa_pay_expense_save.asp" method="post" name="frm">
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
								<th class="first">사번</th>
								<td class="left">
                                <input name="ex_emp_no" type="text" value="<%=ex_emp_no%>" style="width:90px" readonly="true">
                                <a href="#" class="btnType03" onClick="pop_Window('insa_emp_select.asp?gubun=<%="payexp"%>&view_condi=<%=view_condi%>','orgempselect','scrollbars=yes,width=600,height=400')">직원검색</a>
                                </td>
								<th >성명</th>
								<td class="left" >
                                <input name="ex_emp_name" type="text" value="<%=ex_emp_name%>" style="width:90px" readonly="true"></td>
							</tr>
                           	<tr>
								<th class="first">소속</th>
								<td colspan="3" class="left">
                                <input name="ex_company" type="text" value="<%=ex_company%>" style="width:120px" readonly="true">
                                <input name="ex_org_name" type="text" value="<%=ex_org_name%>" style="width:100px" readonly="true">
                                <input name="ex_org_code" type="text" value="<%=ex_org_code%>" style="width:40px" readonly="true">
                                <input name="ex_bonbu" type="hidden" value="<%=ex_bombu%>" style="width:120px" readonly="true">
                                <input name="ex_saupbu" type="hidden" value="<%=ex_saupbu%>" style="width:120px" readonly="true">
                                <input name="ex_team" type="hidden" value="<%=ex_team%>" style="width:120px" readonly="true">
                                <input name="ex_reside_place" type="hidden" value="<%=ex_reside_place%>" style="width:120px" readonly="true">
                                <input name="ex_reside_company" type="hidden" value="<%=ex_reside_company%>" style="width:120px" readonly="true">
                                </td>
							</tr>    
                            <tr>
								<th class="first">귀속년월</th>
								<td class="left" ><input name="rever_yymm" type="text" value="<%=rever_yymm%>" style="width:70px" readonly="true"></td>
                                <th >지급일</th>
								<td class="left"><input name="ex_pay_date" type="text" value="<%=ex_pay_date%>" style="width:70px" id="datepicker"></td>
							</tr>             
							<tr>
								<th class="first" style="background:#F5FFFA">발생일</th>
								<td class="left"><input name="ex_date" type="text" value="<%=ex_date%>" style="width:70px" id="datepicker1"></td>
                                <th style="background:#F5FFFA">항목</th>
                                <td class="left">
                         <%
					        Sql="select * from emp_etc_code where emp_etc_type = '"+etc_type+"' order by emp_etc_code asc"
					        Rs_etc.Open Sql, Dbconn, 1
					     %>
					            <select name="ex_code_name" id="ex_code_name" style="width:130px">
                                   <option value="" <% if ex_code_name = "" then %>selected<% end if %>>선택</option>
                	     <% 
							do until rs_etc.eof 
			  		     %>
                				   <option value='<%=rs_etc("emp_etc_name")%>' 
								   <%If ex_code_name = rs_etc("emp_etc_name") then 
								    ex_code = rs_etc("emp_etc_code")
									ex_tax_id = rs_etc("emp_tax_id")
								   %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                		 <%
									rs_etc.movenext()  
							loop 
							rs_etc.Close()
						 %>
            		            </select>    
                                </td>
							</tr>
                        	<tr>
								<th class="first" style="background:#F5FFFA">금액</th>
								<td class="left">
                                <input name="ex_amount" type="text" value="<%=formatnumber(ex_amount,0)%>" style="width:100px;text-align:right" onKeyUp="num_chk(this);"></td>
								<th style="background:#F8F8FF">근무일수</th>
                                <td class="left">
								<input name="ex_work_cnt" type="text" value="<%=formatnumber(ex_work_cnt,0)%>" style="width:30px;text-align:right"></td>
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
                <input type="hidden" name="ex_tax_id" value="<%=ex_tax_id%>" ID="Hidden1">
                <input type="hidden" name="ex_code" value="<%=ex_code%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

