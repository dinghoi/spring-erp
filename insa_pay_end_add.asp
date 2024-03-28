<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim month_tab(24,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)

curr_date = curr_year + curr_month

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	pmg_yymm = request.form("pmg_yymm")
  else
	pmg_yymm = request("pmg_yymm")
end if

if pmg_yymm = "" then
	ck_sw = "n"
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
end if

etc_type = "99"
etc_code = "9999"

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from emp_etc_code where emp_etc_code = '" + etc_code + "'"
Rs.Open Sql, Dbconn, 1
'Response.write Sql
emp_payend_date = rs("emp_payend_date")
emp_payend_yn = rs("emp_payend_yn")
if emp_payend_yn = "Y" then  ' 윤성희:급여마감을 풀어 달라고 하면 emp_payend_yn 를 'N' 으로 바꾸어주면된다. emp_payend_date도 전달로 맞추고..(select * from emp_etc_code where emp_etc_code = '9999' )
	emp_payend = "마감"
else
  emp_payend = ""
end if

'Response.write sql

title_line = " 급여지급월 마감등록 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
			   window.close () ;
			}
	    </script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.pmg_yymm.value == "") {
					alert ("조건을 입력하시기 바랍니다");
					return false;
				}
				return true;
			}

			function pay_end_date(val) {

            if (!confirm("마감등록을 하시겠습니까 ?")) return;
            var frm = document.frm;
			document.frm.pmg_yymm1.value = document.getElementById(val).value;

            document.frm.action = "insa_pay_end_ok.asp";
            document.frm.submit();
            }
		</script>

	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_end_add.asp?ck_sw=<%="n"%>" method="post" name="frm">
                <fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<label>
								<strong>귀속년월 : </strong>
								<%
								'	Response.write pmg_yymm
								%>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">기 마감월</th>
								<td class="left"><%=rs("emp_payend_date")%></td>
							</tr>
							<tr>
								<th class="first">마감여부</th>
 								<td class="left"><%=emp_payend%>&nbsp;</td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="마감등록" onclick="pay_end_date('pmg_yymm');return false;" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="pmg_yymm1" value="<%=pmg_yymm%>" ID="Hidden1">
                <input type="hidden" name="etc_code" value="<%=etc_code%>" ID="Hidden1">
				</form>
		</div>
	</body>
</html>

