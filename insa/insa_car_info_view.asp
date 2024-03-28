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
Dim drvuser_tab(10,10)
Dim insur_tab(10,20)
Dim as_tab(10,10)
Dim pe_tab(10,10)
Dim drv_tab(10,10)

Dim car_no, car_name, car_year, car_reg_date, oil_kind, title_line
Dim i, j, k
Dim rs_user, rs_ins, rs_as, rs_pe, rsCar

car_no = f_Request("car_no")
car_name = f_Request("car_name")
car_year = f_Request("car_year")
car_reg_date = f_Request("car_reg_date")
oil_kind = f_Request("oil_kind")

title_line = " 차량 정보 "

'운행자 정보
For i = 0 To 10
'	drvuser_tab(i) = ""
'	drvuser_tab(i) = 0
	For j = 0 To 10
		drvuser_tab(i, j) = ""
'		drvuser_tab(i,j) = 0
	Next
Next

'Sql="select * from car_drive_user where use_car_no = '"&car_no&"' order by use_car_no,use_date,use_owner_emp_no DESC"
objBuilder.Append "SELECT use_date, use_emp_name, use_owner_emp_no, use_company, "
objBuilder.Append "	use_org_name, use_org_code, use_emp_grade, use_end_date "
objBuilder.Append "FROM car_drive_user "
objBuilder.Append "WHERE use_car_no = '"&car_no&"' "
objBuilder.Append "ORDER BY use_car_no, use_date, use_owner_emp_no DESC "

Set rs_user = Server.CreateObject("ADODB.RecordSet")
rs_user.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

k = 0
While Not rs_user.EOF
	k = k + 1

	drvuser_tab(k, 1) = rs_user("use_date")
	drvuser_tab(k, 2) = rs_user("use_emp_name")
	drvuser_tab(k, 3) = rs_user("use_owner_emp_no")
	drvuser_tab(k, 4) = rs_user("use_company")
	drvuser_tab(k, 5) = rs_user("use_org_name")
	drvuser_tab(k, 6) = rs_user("use_org_code")
	drvuser_tab(k, 7) = rs_user("use_emp_grade")
	drvuser_tab(k, 8) = rs_user("use_end_date")

	rs_user.MoveNext()
Wend
rs_user.Close() : Set rs_user = Nothing

'보험 정보
For i = 0 To 10
'	insur_tab(i) = ""
'	insur_tab(i) = 0
	For j = 0 To 20
		insur_tab(i, j) = ""
'		insur_tab(i,j) = 0
	Next
Next

'Sql="select * from car_insurance where ins_car_no = '"&car_no&"' order by ins_car_no,ins_date DESC"
objBuilder.Append "SELECT ins_date, ins_amount, ins_company, ins_last_date, ins_man1, ins_man2, "
objBuilder.Append "	ins_object, ins_self, ins_injury, ins_self_car, ins_age, ins_scramble, "
objBuilder.Append "	ins_contract_yn, ins_comment "
objBuilder.Append "FROM car_insurance "
objBuilder.Append "WHERE ins_car_no = '"&car_no&"' "
objBuilder.Append "ORDER BY ins_car_no, ins_date DESC "

Set rs_ins = Server.CreateObject("ADODB.RecordSet")
rs_ins.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

k = 0
While Not rs_ins.EOF
	k = k + 1

	insur_tab(k, 1) = rs_ins("ins_date")
	insur_tab(k, 2) = rs_ins("ins_amount")
	insur_tab(k, 3) = rs_ins("ins_company")
	insur_tab(k, 4) = rs_ins("ins_last_date")
	insur_tab(k, 5) = rs_ins("ins_man1")
	insur_tab(k, 6) = rs_ins("ins_man2")
	insur_tab(k, 7) = rs_ins("ins_object")
	insur_tab(k, 8) = rs_ins("ins_self")
	insur_tab(k, 9) = rs_ins("ins_injury")
	insur_tab(k, 10) = rs_ins("ins_self_car")
	insur_tab(k, 11) = rs_ins("ins_age")
	insur_tab(k, 12) = rs_ins("ins_scramble")

	If rs_ins("ins_contract_yn") = "Y" Then
		insur_tab(k, 13) = "계약내용포함"
	Else
		insur_tab(k, 13) = "계약내용미포함" & rs_ins("ins_comment")
	End If

	rs_ins.MoveNext()
Wend
rs_ins.Close() : Set rs_ins = Nothing

'AS 정보
For i = 0 To 10
'	as_tab(i) = ""
'	as_tab(i) = 0
	For j = 0 To 10
		as_tab(i, j) = ""
'		as_tab(i,j) = 0
	Next
Next

'Sql="select * from car_as where as_car_no = '"&car_no&"' order by as_car_no,as_date,as_seq DESC"
objBuilder.Append "SELECT as_date, as_cause, as_solution, as_amount "
objBuilder.Append "	as_amount_sign, as_owner_emp_no, as_owner_emp_name "
objBuilder.Append "FROM car_as "
objBuilder.Append "WHERE as_car_no = '"&car_no&"' "
objBuilder.Append "ORDER BY as_car_no, as_date, as_seq DESC "

Set rs_as = Server.CreateObject("ADODB.RecordSet")
rs_as.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

k = 0
While Not rs_as.EOF
	k = k + 1

	as_tab(k, 1) = rs_as("as_date")
	as_tab(k, 2) = rs_as("as_cause")
	as_tab(k, 3) = rs_as("as_solution")
	as_tab(k, 4) = rs_as("as_amount")
	as_tab(k, 5) = rs_as("as_amount_sign")
	as_tab(k, 6) = rs_as("as_owner_emp_no")
	as_tab(k, 7) = rs_as("as_owner_emp_name")

	rs_as.MoveNext()
Wend
rs_as.Close() : Set rs_as = Nothing

'과태료 정보
For i = 0 To 10
'	pe_tab(i) = ""
'	pe_tab(i) = 0
	For j = 0 To 10
		pe_tab(i, j) = ""
'		pe_tab(i,j) = 0
	Next
Next

'Sql="select * from car_penalty where pe_car_no = '"&car_no&"' order by pe_car_no,pe_date,pe_seq DESC"

objBuilder.Append "SELECT pe_date, pe_amount, pe_comment, pe_place, pe_default, "
objBuilder.Append "	pe_in_date, pe_notice_date, pe_notice, pe_owner_emp_name, pe_owner_emp_no "
objBuilder.Append "FROM car_penalty "
objBuilder.Append "WHERE pe_car_no = '"&car_no&"' "
objBuilder.Append "ORDER BY pe_car_no, pe_date, pe_seq DESC "

Set rs_pe = Server.CreateObject("ADODB.RecordSet")
rs_pe.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

k = 0
While Not rs_pe.EOF
	k = k + 1

	pe_tab(k, 1) = rs_pe("pe_date")
	pe_tab(k, 2) = rs_pe("pe_amount")
	pe_tab(k, 3) = rs_pe("pe_comment")
	pe_tab(k, 4) = rs_pe("pe_place")
	pe_tab(k, 5) = rs_pe("pe_default")
	pe_tab(k, 6) = rs_pe("pe_in_date")
	pe_tab(k, 7) = rs_pe("pe_notice_date")
	pe_tab(k, 8) = rs_pe("pe_notice")
	pe_tab(k, 9) = rs_pe("pe_owner_emp_name")
	pe_tab(k, 10) = rs_pe("pe_owner_emp_no")

	rs_pe.MoveNext()
Wend
rs_pe.Close() : Set rs_pe = Nothing

'sql = "select * from car_info where car_no = '" + car_no + "'"
objBuilder.Append "SELECT car_owner, car_company, start_date, car_use, car_use_dept, "
objBuilder.Append "	buy_gubun, rental_company, last_check_date, last_km, "
objBuilder.Append "	car_status, end_date, car_comment "
objBuilder.Append "FROM car_info "
objBuilder.Append "WHERE car_no = '"&car_no&"'"

Set rsCar = Server.CreateObject("ADODB.RecordSet")
rsCar.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
		/*
			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm() {
				if(document.frm.in_carno.value =="") {
					alert('차량명을 입력하세요');
					frm.in_name.focus();
					return false;}
				{
					return true;
				}
			}
		*/
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<!--<form action="insa_car_drvuser_view.asp?car_no=<%=car_no%>" method="post" name="frm">-->
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>차량번호 : </strong>
								<label>
        						<input name="in_carno" type="text" id="in_carno" value="<%=car_no%>" style="width:100px; text-align:left" readonly="true">
								</label>
                            <strong>차종/연식/취득일 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=car_name%>" style="width:100px; text-align:left" readonly="true">
                                -
                                <input name="in_year" type="text" id="in_year" value="<%=car_year%>" style="width:70px; text-align:left" readonly="true">
                                -
                                <input name="car_reg_date" type="text" id="car_reg_date" value="<%=car_reg_date%>" style="width:70px; text-align:left" readonly="true">
                                -
                                <input name="oil_kind" type="text" id="oil_kind" value="<%=oil_kind%>" style="width:50px; text-align:left" readonly="true">
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%">
							<col width="15%">
                            <col width="15%">
                            <col width="15%">
                            <col width="15%">
                            <col width="*">
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">소유구분</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">소유회사</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">등록일</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">용도</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">사용부서</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">구매</th>
							</tr>
                            <tr>
                                <th class="first" scope="col">최종검사일</th>
                                <th scope="col">운행Km</th>
                                <th scope="col">차량상태</th>
                                <th scope="col">처분일</th>
                                <th colspan="2" scope="col">차량정보</th>
							</tr>
						</thead>
						<tbody>
						<%
							Do Until rsCar.EOF or rsCar.BOF
						%>
							<tr>
								<td><%=rsCar("car_owner")%>&nbsp;</td>
								<td><%=rsCar("car_company")%>&nbsp;</td>
                                <td><%=rsCar("start_date")%>&nbsp;</td>
                                <td><%=rsCar("car_use")%>&nbsp;</td>
                                <td><%=rsCar("car_use_dept")%>&nbsp;</td>
                                <td><%=rsCar("buy_gubun")%>(<%=rsCar("rental_company")%>)&nbsp;</td>
							</tr>
                            <tr>
								<td><%=rsCar("last_check_date")%>&nbsp;</td>
								<td><%=FormatNumber(rsCar("last_km"), 0)%>&nbsp;</td>
                                <td><%=rsCar("car_status")%>&nbsp;</td>
                                <td><%=rsCar("end_date")%>&nbsp;</td>
                                <td colspan="2"><%=rsCar("car_comment")%>&nbsp;</td>
							</tr>
							<%
								rsCar.MoveNext()
							Loop
							rsCar.Close() : Set rsCar = Nothing
							%>
						</tbody>
					</table>
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%">
							<col width="15%">
                            <col width="15%">
                            <col width="*">
                            <col width="15%">
                            <col width="15%">
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">운행시작일</th>
                                <th scope="col">운행자</th>
                                <th scope="col">소속회사</th>
                                <th scope="col">부서</th>
                                <th scope="col">직위</th>
                                <th scope="col">종료일</th>
							</tr>
						</thead>
						<tbody>
						<%
						For i = 0 To 10
                        	If drvuser_tab(i, 1) <> "" Then
						%>
							<tr>
								<td><%=drvuser_tab(i, 1)%>&nbsp;</td>
								<td><%=drvuser_tab(i, 2)%>(<%=drvuser_tab(i, 3)%>)&nbsp;</td>
                                <td><%=drvuser_tab(i, 4)%>&nbsp;</td>
                                <td><%=drvuser_tab(i, 5)%>(<%=drvuser_tab(i, 6)%>)&nbsp;</td>
                                <td><%=drvuser_tab(i, 7)%>&nbsp;</td>
                                <td><%=drvuser_tab(i, 8)%>&nbsp;</td>
							</tr>
						<%
							End If
						Next
						%>
						</tbody>
					</table>
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%">
							<col width="15%">
                            <col width="15%">
                            <col width="15%">
                            <col width="15%">
                            <col width="15%">
                            <col width="*">
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">보험가입일</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">보험료</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">보험회사</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">보험만기일</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">대인1</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">대인2</th>
                                <th scope="col" style=" border-bottom:1px solid #e3e3e3;">대물</th>
							</tr>
                            <tr>
                                <th class="first" scope="col">자기</th>
                                <th scope="col">무상해</th>
                                <th scope="col">자차</th>
                                <th scope="col">연령</th>
                                <th scope="col">긴급출동</th>
                                <th colspan="2" scope="col">계약내용포함</th>
							</tr>
						</thead>
						<tbody>
						<%
						For i = 0 To 10
                        	If insur_tab(i, 1) <> "" Then
						%>
							<tr>
								<td><%=insur_tab(i, 1)%>&nbsp;</td>
								<td><%=FormatNumber(insur_tab(i, 2), 0)%>&nbsp;</td>
                                <td><%=insur_tab(i, 3)%>&nbsp;</td>
                                <td><%=insur_tab(i, 4)%>&nbsp;</td>
                                <td><%=insur_tab(i, 5)%>&nbsp;</td>
                                <td><%=insur_tab(i, 6)%>&nbsp;</td>
                                <td><%=insur_tab(i, 7)%>&nbsp;</td>
							</tr>
                            <tr>
								<td><%=insur_tab(i, 8)%>&nbsp;</td>
								<td><%=insur_tab(i, 9)%>&nbsp;</td>
                                <td><%=insur_tab(i, 10)%>&nbsp;</td>
                                <td><%=insur_tab(i, 11)%>&nbsp;</td>
                                <td><%=insur_tab(i, 12)%>&nbsp;</td>
                                <td colspan="2"><%=insur_tab(i, 13)%>&nbsp;</td>
							</tr>
						<%
							End If
						Next
						%>
						</tbody>
					</table>
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%">
							<col width="15%">
                            <col width="20%">
                            <col width="*">
                            <col width="20%">
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">AS일자</th>
                                <th scope="col">수리비용</th>
                                <th scope="col">증상/원인</th>
                                <th scope="col">수리내용</th>
                                <th scope="col">운행자</th>
							</tr>
						</thead>
						<tbody>
						<%
						For i = 0 To 10
                        	If as_tab(i, 1) <> "" Then
						%>
							<tr>
								<td><%=as_tab(i, 1)%>&nbsp;</td>
								<td><%=FormatNumber(insur_tab(i, 4), 0)%>&nbsp;(<%=as_tab(i, 5)%>)</td>
                                <td class="left"><%=as_tab(i, 2)%>&nbsp;</td>
                                <td class="left"><%=as_tab(i, 3)%>&nbsp;</td>
                                <td><%=as_tab(i, 7)%>(<%=as_tab(i, 6)%>)&nbsp;</td>
							</tr>
						<%
							End If
						Next
						%>
						</tbody>
					</table>
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="9%">
							<col width="9%">
                            <col width="10%">
                            <col width="*">
                            <col width="10%">
                            <col width="9%">
                            <col width="20%">
                            <col width="12%">
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">위반일자</th>
                                <th scope="col">과태료</th>
                                <th scope="col">위반내용</th>
                                <th scope="col">위반장소</th>
                                <th scope="col">미납</th>
                                <th scope="col">납입일자</th>
                                <th scope="col">통보일자</th>
                                <th scope="col">운행자</th>
							</tr>
						</thead>
						<tbody>
						<%
						For i = 0 To 10
                        	If pe_tab(i, 1) <> "" Then
						%>
							<tr>
								<td><%=pe_tab(i, 1)%>&nbsp;</td>
								<td><%=FormatNumber(pe_tab(i, 2), 0)%>&nbsp</td>
                                <td class="left"><%=pe_tab(i, 3)%>&nbsp;</td>
                                <td class="left"><%=pe_tab(i, 4)%>&nbsp;</td>
                                <td><%=pe_tab(i, 5)%>&nbsp;</td>
                                <td><%=pe_tab(i, 6)%>&nbsp;</td>
                                <td><%=pe_tab(i, 7)%>(<%=pe_tab(i, 8)%>)&nbsp;</td>
                                <td><%=pe_tab(i, 9)%>(<%=pe_tab(i, 10)%>)&nbsp;</td>
							</tr>
						<%
							End If
						Next
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div align="right">
                    <br/>
						<a href="#" class="btnType04" onclick="javascript:toclose();" >닫기</a>&nbsp;&nbsp;
					</div>
                    </td>
			      </tr>
			  </table>
         </div>
	<!--</form>-->
	  </div>
	</body>
</html>