<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

inc_yyyy = cint(mid(now(),1,4)) - 1

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
rs_emp.Open Sql, Dbconn, 1
emp_in_date = rs_emp("emp_in_date")
emp_name = rs_emp("emp_name")
emp_grade = rs_emp("emp_grade")
emp_position = rs_emp("emp_position")
emp_company = rs_emp("emp_company")
emp_org_name = rs_emp("emp_org_name")

title_line = "연말정산 최종제출 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
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
				if (document.frm.view_condi.value == "") {
					alert ("조건을 입력하시기 바랍니다");
					return false;
				}	
				return true;
			}
			
			function insa_pay_yeartax_final(val) {

            if (!confirm("최종제출시 더이상 수정이 불가 입니다. 최종제출 하시겠습니까 ?")) return;
            var frm = document.frm;
			document.frm.emp_no1.value = document.getElementById(val).value;
			
            document.frm.action = "insa_pay_yeartax_final_submit_ok.asp";
            document.frm.submit();
            }	
		</script>

	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_user_password.asp?ck_sw=<%="n"%>" method="post" name="frm">
                <fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong><%=inc_yyyy%> 년 귀속 연말정산 최종제출 </strong>
								<label>
        						<input name="emp_no" type="hidden" id="emp_no" value="<%=emp_no%>" style="width:40px" readonly="true">
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
                <h3 class="stit">※ 최종제출을 하시면 <%=inc_yyyy%>  년 귀속 연말정산 등록을 모두 마침니다.<br>&nbsp;<br></h3>                
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th colspan="2" class="left" style=" border-bottom:1px solid #ffffff;">최종제출 이후에는 더이상 등록 및 수정이 불가 합니다(단, 조회는 가능)<br><br>최종제출을 하시려면 마래 버튼을 글릭하시기 바랍니다<br><br>최종제출을 하시면 소득공제명서 및 의료비/기부금/신용카드명세서를 출력할 수 있습니다.</th>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="최종제출" onclick="insa_pay_yeartax_final('emp_no');return false;" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="inc_yyyy" value="<%=inc_yyyy%>" ID="Hidden1">
                <input type="hidden" name="emp_no1" value="<%=emp_no1%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

