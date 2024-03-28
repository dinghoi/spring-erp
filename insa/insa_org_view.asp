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
Dim org_code, title_line
Dim org_level, org_date, org_end_date, org_empno
Dim org_empname, org_company, org_bonbu, org_saupbu, org_team
Dim owner_org, owner_empno, owner_empname, org_reside_company
Dim org_table_company, tel_ddd, tel_no1, tel_no2, org_sido
Dim org_gugun, org_dong, org_addr, org_reg_date, org_reg_user
Dim rsOrg, org_mod_date, org_mod_user, org_table_org
Dim rs_owner, owner_orgname, rs_emp, emp_image, photo_image
Dim mg_ce, strSql, arrOrg

org_code = request("org_code")

title_line = " 조직 상세 조회 "

strSql = "CALL USP_INSA_ORG_VIEW('"&org_code&"')"

Set rsOrg = DBConn.Execute(strSql)

If Not rsOrg.EOF Then
	arrOrg = rsOrg.getRows()
End If

Call Rs_Close(rsOrg)

If IsArray(arrOrg) Then
	org_level = arrOrg(0, 0)
	org_name = arrOrg(1, 0)
	org_date = arrOrg(2, 0)
	org_end_date = arrOrg(3, 0)
	org_empno = arrOrg(4, 0)
	org_empname = arrOrg(5, 0)
	org_company = arrOrg(6, 0)
	org_bonbu = arrOrg(7, 0)
	org_saupbu = arrOrg(8, 0)
	org_team = arrOrg(9, 0)
	owner_org = arrOrg(10, 0)
	owner_empno = arrOrg(11, 0)
	owner_empname = arrOrg(12, 0)
	org_reside_company = arrOrg(13, 0)
	org_table_org = arrOrg(14, 0)
	tel_ddd = arrOrg(15, 0)
	tel_no1 = arrOrg(16, 0)
	tel_no2 = arrOrg(17, 0)
	org_sido = arrOrg(18, 0)
	org_gugun = arrOrg(19, 0)
	org_dong = arrOrg(20, 0)
	org_addr = arrOrg(21, 0)
	org_reg_date = arrOrg(22, 0)
	org_reg_user = arrOrg(23, 0)
	org_mod_date = arrOrg(24, 0)
	org_mod_user = arrOrg(25, 0)
	owner_orgname = arrOrg(26, 0)
	emp_image = arrOrg(27, 0)
End If

DBConn.Close() : Set DBConn = Nothing

photo_image = "/emp_photo/" & emp_image
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>인사 관리 시스템</title>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>

</head>
<body>
	<div id="container">
		<h3 class="insa"><%=title_line%></h3><br/>
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
						<img src="<%=photo_image%>" width="110" height="120" alt="">
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
		<div align="right">
			<a href="#" class="btnType04" onclick="close_win();" >닫기</a>&nbsp;&nbsp;
		</div>
		<br>
	</div>
</body>
</html>