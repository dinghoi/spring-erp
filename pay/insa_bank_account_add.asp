<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim u_type, emp_name, person_no1, person_no2
Dim emp_type, emp_pay_type, bank_code, bank_name, account_no, account_holder
Dim curr_date, rs_emp, emp_bonbu, emp_saupbu, emp_team, emp_org_code, emp_org_name
Dim title_line, rsBank

u_type = f_Request("u_type")
emp_no = f_Request("emp_no")
emp_name = f_Request("emp_name")
person_no1 = f_Request("emp_person1")
person_no2 = f_Request("emp_person2")

emp_type = ""
emp_pay_type = ""
bank_code = ""
bank_name = ""
account_no = ""
account_holder = emp_name

curr_date = Mid(CStr(Now()), 1, 10)

objBuilder.Append "SELECT emp_no, emp_name, emp_company, emp_bonbu, emp_saupbu, emp_team, emp_org_code, emp_org_name, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emp_no = '"&emp_no&"' "

Set rs_emp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rs_emp.eof Then
	emp_no = rs_emp("emp_no")
	emp_name = rs_emp("emp_name")
	emp_company = rs_emp("org_company")
	emp_bonbu = rs_emp("org_bonbu")
	emp_saupbu = rs_emp("org_saupbu")
	emp_team = rs_emp("org_team")
	emp_org_code = rs_emp("emp_org_code")
	emp_org_name = rs_emp("org_name")
Else
	emp_name = ""
	emp_company = ""
	emp_bonbu = ""
	emp_saupbu = ""
	emp_team = ""
	emp_org_code = ""
	emp_org_name = ""
End If

rs_emp.Close() : Set rs_emp = Nothing

title_line = " 직원 은행계좌 등록 "



If u_type = "U" Then

	objBuilder.Append "SELECT emp_type, emp_pay_type, bank_code, person_no1, person_no2, bank_name, account_no, account_holder "
	objBuilder.Append "FROM pay_bank_account WHERE emp_no = '"&emp_no&"' "

	Set rsBank = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	emp_type = rsBank("emp_type")
    emp_pay_type = rsBank("emp_pay_type")
    bank_code = rsBank("bank_code")
	person_no1 = rsBank("person_no1")
	person_no2 = rsBank("person_no2")
    bank_name = rsBank("bank_name")
    account_no = rsBank("account_no")
	account_holder = rsBank("account_holder")

	rsBank.Close() : Set rsBank = Nothing

	title_line = " 직원 은행계좌 변경 "
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
			function goAction(){
			   window.close();
			}

			function goBefore(){
			   history.back();
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.bank_name.value ==""){
					alert('은행명을 선택하세요');
					frm.bank_name.focus();
					return false;
				}

				if(document.frm.account_no ==""){
					alert('계좌번호를 선택하세요');
					frm.account_no.focus();
					return false;
				}

				if(document.frm.account_holder.value ==""){
					alert('예금주를 입력하세요');
					frm.account_holder.focus();
					return false;
				}

				{
					a=confirm('입력하시겠습니까?');

					if(a==true){
						return true;
					}
					return false;
				}
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/pay/insa_bank_account_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="*" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">사번</th>
                      <td class="left" bgcolor="#FFFFE6"><%=emp_no%></td>
					  <input name="emp_no" type="hidden" id="emp_no" size="14" value="<%=emp_no%>" readonly="true"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6"><%=emp_name%></td>
					  <input name="emp_name" type="hidden" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th style="background:#FFFFE6">주민등록<br>번호</th>
                      <td colspan="5" class="left" bgcolor="#FFFFE6"><%=person_no1%> - <%=person_no2%></td>
					  <input name="person_no1" type="hidden" id="person_no1" size="14" value="<%=person_no1%>" readonly="true">
                      <input name="person_no2" type="hidden" id="person_no2" size="14" value="<%=person_no2%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>은행명</th>
                      <td colspan="5" class="left">
					<%
					Dim rs_etc
					' Sql="select emp_etc_name from emp_etc_code where emp_etc_type = '50' order by emp_etc_code asc"
					objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code WHERE emp_etc_type = '50' ORDER BY emp_etc_code ASC "

					Set rs_etc = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()
					%>
					  <select name="bank_name" id="bank_name" style="width:130px">
                         <option value="" <%If bank_name = "" Then %>selected<%End If %>>선택</option>
                			  <%
								Do Until rs_etc.EOF
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If bank_name = rs_etc("emp_etc_name") Then %>selected<%End If %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.MoveNext()
								Loop
								rs_etc.Close() : Set rs_etc = Nothing

								DBConn.Close() : Set DBConn = Nothing
							  %>
            		  </select>
                      </td>
                    </tr>
                    <tr>
                      <th>계좌번호</th>
                      <td colspan="5" class="left">
					  <input name="account_no" type="text" id="account_no" size="20" value="<%=account_no%>"></td>
                    </tr>
                    <tr>
                      <th>예금주</th>
                      <td colspan="5" class="left">
					  <input name="account_holder" type="text" id="account_holder" size="14" value="<%=account_holder%>"></td>
					  </td>
			    	</tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align="center">
				<%
				'if end_sw = "N" then
				%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" /></span>
        		<%
				'end if
				%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" />
			</form>
		</div>
	</body>
</html>