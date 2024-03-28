<%@ Language=VBScript %>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

'If view_condi = "" Then
'	view_condi = "케이원정보통신"
'End If

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

'sql = "select * from emp_master WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"&view_condi&"') and (emp_no < '900000') ORDER BY emp_no ASC"
sql = "select * from emp_master WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') ORDER BY emp_no ASC"
Rs.Open Sql, Dbconn, 1

title_line = " 직원 현황 "

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
			function getPageCode(){
				return "1 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("필드조건을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

</head> 

<% '<body bgcolor="#00274f" text="#eefde3"> %>
<body>

				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
                            <col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" style="font-size:11px;">사번</th>
								<th scope="col" style="font-size:11px;">성  명</th>
                                <th scope="col" style="font-size:11px;">직급</th>
							</tr>
						</thead>
					<tbody>
						<%
						do until rs.eof

						%>
							<tr>
								<td style="font-size:11px;"><%=rs("emp_no")%></td>
                                <td style="font-size:11px;">
                                <a href="insa_emp_infor_view.asp?emp_no=<%=rs("emp_no")%>"
    target="right"><%=rs("emp_name")%></a>
                                </td>
                                <td style="font-size:11px;"><%=rs("emp_grade")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>
