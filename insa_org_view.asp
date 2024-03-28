<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
org_code = request("org_code")
org_name = request("org_name")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs_memb = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_owner = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 조직 상세조회 "

	Sql="select * from emp_org_mst where org_code = '"&org_code&"'"
	Set rs=DbConn.Execute(Sql)

    org_level = rs("org_level")
    org_name = rs("org_name")
    org_date = rs("org_date")
	org_end_date = rs("org_end_date")
    org_empno = rs("org_empno")
    org_empname = rs("org_emp_name")
    org_company = rs("org_company")
    org_bonbu = rs("org_bonbu")
    org_saupbu = rs("org_saupbu")
    org_team = rs("org_team")
    owner_org = rs("org_owner_org")
    owner_empno = rs("org_owner_empno")
    owner_empname = rs("org_owner_empname")
	org_reside_company = rs("org_reside_company")
	org_table_org = rs("org_table_org")
    tel_ddd = rs("org_tel_ddd")
    tel_no1 = rs("org_tel_no1")
    tel_no2 = rs("org_tel_no2")
	org_sido = rs("org_sido")
    org_gugun = rs("org_gugun")
    org_dong = rs("org_dong")
    org_addr = rs("org_addr")
    org_end_date = rs("org_end_date")
    org_reg_date = rs("org_reg_date")
	org_reg_user = rs("org_reg_user")
    org_mod_date = rs("org_mod_date")
    org_mod_user = rs("org_mod_user")
	rs.close()
    
	Sql="select * from emp_org_mst where org_code = '"&owner_org&"'"
	Set rs_owner=DbConn.Execute(Sql)

    owner_orgname = rs_owner("org_name")
	rs_owner.close()

Sql = "select * from emp_master where emp_no = '"&org_empno&"'"
Set rs_emp = DbConn.Execute(SQL)

emp_image = ""
if	rs_emp.eof or rs_emp.bof then
		mg_ce = "ERROR"
	  else
		emp_image = rs_emp("emp_image")
	end if
	rs_emp.close()						


photo_image = "/emp_photo/" + emp_image


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
												$( "#datepicker" ).datepicker("setDate", "<%=org_end_date%>" );
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
				if(document.frm.org_end_date.value =="") {
					alert('조직폐쇄일을 입력하세요');
					frm.org_end_date.focus();
					return false;}
				
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
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
				<form action="insa_org_view.asp?org_code=<%=org_code%>&org_name=<%=org_name%>&u_type=<%=u_type%>" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="12%" >
							<col width="12%" >
							<col width="12%" >
							<col width="12%" >
							<col width="12%" >
                            <col width="13%" >
                            <col width="12%" >
						</colgroup>
						<tbody>
							<tr>
								<td rowspan="4" class="first">
                                <img src="<%=photo_image%>" width=110 height=120 alt="">
                                </td>
                                <th>조직코드</th>
                                <td class="left"><%=org_code%>&nbsp;)<%=org_level%>&nbsp;</td>
                                <th>조직명</th>
                                <td colspan="2" class="left"><%=org_name%>&nbsp;</td>
                                <th>조직생성일</th>
                                <td class="left"><%=org_date%>&nbsp;</td>
                             </tr>
                             <tr>
								<th>조직장사번</th>
                                <td class="left"><%=org_empno%>&nbsp;</td>
                                <th>조직장성명</th>
                                <td colspan="2" class="left"><%=org_empname%>&nbsp;</td>
                                <th>조직폐쇄일</th>
                                <td class="left"><%=org_end_date%>&nbsp;</td>
                                </td>
                             </tr>
                             <tr>
                                <th>소속</th>
                                <td colspan="6" class="left"><%=org_company%>&nbsp;&nbsp;&nbsp;<%=org_bonbu%>&nbsp;&nbsp;&nbsp;<%=org_saupbu%>&nbsp;&nbsp;&nbsp;<%=org_team%>&nbsp;</td>
                             </tr>
							<tr>
								<th>상위조직</th>
                                <td colspan="3" class="left"><%=owner_org%>&nbsp;)<%=owner_orgname%>&nbsp;</td>
                                <th>상위조직장</th>
                                <td colspan="2" class="left"><%=owner_empno%>&nbsp;)<%=owner_empname%>&nbsp;</td>
                             </tr>
                             <tr>
								<th class="first">대표전화</th>
                                <td colspan="2" class="left"><%=tel_ddd%>&nbsp;-&nbsp;<%=tel_no1%>&nbsp;-&nbsp;<%=tel_no2%>&nbsp;</td>
								<th>주소</th>
								<td colspan="4" class="left"><%=org_sido%>&nbsp;&nbsp;<%=org_gugun%>&nbsp;&nbsp;<%=org_dong%>&nbsp;&nbsp;<%=org_addr%>&nbsp;</td>
                                </td>
                             </tr>
                             <tr>
                                <th class="first">입력일자</th>
                                <td colspan="2" class="left"><%=org_reg_date%>&nbsp;</td>
                                <th>입력자 명</th>
                                <td class="left"><%=org_reg_user%>&nbsp;</td>
                                <th>조직 T.O</th>
                                <td colspan="2" class="left"><%=org_table_org%>&nbsp;</td>
                             </tr>
                             <tr>
                                <th class="first">변경일자</th>
                                <td colspan="2" class="left"><%=org_mod_date%>&nbsp;</td>
                                <th>변경자 명</th>
                                <td class="left"><%=org_mod_user%>&nbsp;</td>
                                <th>상주처 회사</th>
                                <td colspan="2" class="left"><%=org_reside_company%>&nbsp;</td>
                             </tr>
			        </tbody>
			      </table>
				</div>
                   	<br>
               		<div align=right>
						<a href="#" class="btnType04" onclick="javascript:goAction()" >닫기</a>&nbsp;&nbsp;
					</div>
                    <br>               				
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="org_code" value="<%=org_code%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

