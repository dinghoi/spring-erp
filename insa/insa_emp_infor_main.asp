<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

curr_date = datevalue(mid(cstr(now()),1,10))

If view_condi = "" Then
	view_condi = "케이원정보통신"
End If

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

sql = "select * from emp_master WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"&view_condi&"') and (emp_no < '900000') ORDER BY emp_no ASC"
Rs.Open Sql, Dbconn, 1

title_line = " 직원 현황 "

emp_no = "100001"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
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
			/*function getPageCode(){
				return "0 1";
			}*/

			/*function frmcheck () {
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
			}*/
		</script>

	</head>

<frameset rows="20%,80%">
  <frame  src="/insa/insa_emp_infor_top.asp" scrolling="no" mane="top">
<frameset cols="150,*" border="0">
  <frame name="left" src="/insa/insa_emp_infor_left.asp" target="main">
  <frame name="right" src="/insa/insa_emp_infor_view.asp?emp_no=<%=emp_no%>" >
</frameset>

  <noframes>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#nclude virtual = "/include/insa_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa/insa_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>회사 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">

                			  <%
								do until rs_org.eof
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()
								loop
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
              <input type="hidden" name="user_id">
		      <input type="hidden" name="pass">
			</form>
         </div>
	</div>
	</body>
  </noframes>
<!--</frameset>-->
</html>
</frameset>