<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
dim drvuser_tab(10,10)
dim insur_tab(10,20)
dim as_tab(10,10)
dim pe_tab(10,10)
dim drv_tab(10,10)

car_no = request("car_no")
car_name = request("car_name")
car_year = request("car_year")
car_reg_date = request("car_reg_date")
oil_kind = request("oil_kind")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs_user = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set rs_as = Server.CreateObject("ADODB.Recordset")
Set rs_pe = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from car_info where car_no = '" + car_no + "'"
Rs.Open Sql, Dbconn, 1

'운행자 db
for i = 0 to 10
'	drvuser_tab(i) = ""
'	drvuser_tab(i) = 0
	for j = 0 to 10
		drvuser_tab(i,j) = ""
'		drvuser_tab(i,j) = 0
	next
next

	k = 0
    Sql="select * from car_drive_user where use_car_no = '"&car_no&"' order by use_car_no,use_date,use_owner_emp_no DESC"
	rs_user.Open Sql, Dbconn, 1	
	while not rs_user.eof
		k = k + 1
		drvuser_tab(k,1) = rs_user("use_date")
		drvuser_tab(k,2) = rs_user("use_emp_name")
		drvuser_tab(k,3) = rs_user("use_owner_emp_no")
		drvuser_tab(k,4) = rs_user("use_compay")
		drvuser_tab(k,5) = rs_user("use_org_name")
		drvuser_tab(k,6) = rs_user("use_org_code")
		drvuser_tab(k,7) = rs_user("use_emp_grade")
		drvuser_tab(k,8) = rs_user("use_end_date")
		rs_user.movenext()
	Wend
    rs_user.close()	
	
'보험 db
for i = 0 to 10
'	insur_tab(i) = ""
'	insur_tab(i) = 0
	for j = 0 to 20
		insur_tab(i,j) = ""
'		insur_tab(i,j) = 0
	next
next

	k = 0
    Sql="select * from car_insurance where ins_car_no = '"&car_no&"' order by ins_car_no,ins_date DESC"
	rs_ins.Open Sql, Dbconn, 1	
	while not rs_ins.eof
		k = k + 1
		insur_tab(k,1) = rs_ins("ins_date")
		insur_tab(k,2) = rs_ins("ins_amount")
		insur_tab(k,3) = rs_ins("ins_company")
		insur_tab(k,4) = rs_ins("ins_last_date")
		insur_tab(k,5) = rs_ins("ins_man1")
		insur_tab(k,6) = rs_ins("ins_man2")
		insur_tab(k,7) = rs_ins("ins_object")
		insur_tab(k,8) = rs_ins("ins_self")
		insur_tab(k,9) = rs_ins("ins_injury")
		insur_tab(k,10) = rs_ins("ins_self_car")
		insur_tab(k,11) = rs_ins("ins_age")
		insur_tab(k,12) = rs_ins("ins_scramble")
		if rs_ins("ins_contract_yn") = "Y" then 
		       insur_tab(k,13) = "계약내용포함"
		   else	   
		       insur_tab(k,13) = "계약내용미포함" + rs_ins("ins_comment")
	    end if
		rs_ins.movenext()
	Wend
    rs_ins.close()		

'AS db
for i = 0 to 10
'	as_tab(i) = ""
'	as_tab(i) = 0
	for j = 0 to 10
		as_tab(i,j) = ""
'		as_tab(i,j) = 0
	next
next

	k = 0
    Sql="select * from car_as where as_car_no = '"&car_no&"' order by as_car_no,as_date,as_seq DESC"
	rs_as.Open Sql, Dbconn, 1	
	while not rs_as.eof
		k = k + 1
		as_tab(k,1) = rs_as("as_date")
		as_tab(k,2) = rs_as("as_cause")
		as_tab(k,3) = rs_as("as_solution")
		as_tab(k,4) = rs_as("as_amount")
		as_tab(k,5) = rs_as("as_amount_sign")
		as_tab(k,6) = rs_as("as_owner_emp_no")
		as_tab(k,7) = rs_as("as_owner_emp_name")
		rs_as.movenext()
	Wend
    rs_as.close()	

'과태료 db
for i = 0 to 10
'	pe_tab(i) = ""
'	pe_tab(i) = 0
	for j = 0 to 10
		pe_tab(i,j) = ""
'		pe_tab(i,j) = 0
	next
next

	k = 0
    Sql="select * from car_penalty where pe_car_no = '"&car_no&"' order by pe_car_no,pe_date,pe_seq DESC"
	rs_pe.Open Sql, Dbconn, 1	
	while not rs_pe.eof
		k = k + 1
		pe_tab(k,1) = rs_pe("pe_date")
		pe_tab(k,2) = rs_pe("pe_amount")
		pe_tab(k,3) = rs_pe("pe_comment")
		pe_tab(k,4) = rs_pe("pe_place")
		pe_tab(k,5) = rs_pe("pe_default")
		pe_tab(k,6) = rs_pe("pe_in_date")
		pe_tab(k,7) = rs_pe("pe_notice_date")
		pe_tab(k,8) = rs_pe("pe_notice")
		pe_tab(k,9) = rs_pe("pe_owner_emp_name")
		pe_tab(k,10) = rs_pe("pe_owner_emp_no")
		rs_pe.movenext()
	Wend
    rs_pe.close()	

title_line = " 차량 정보 "

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
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
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
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_drvuser_view.asp?car_no=<%=car_no%>" method="post" name="frm">
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
							do until rs.eof or rs.bof
						%>
							<tr>
								<td><%=rs("car_owner")%>&nbsp;</td>
								<td><%=rs("car_company")%>&nbsp;</td>
                                <td><%=rs("start_date")%>&nbsp;</td>
                                <td><%=rs("car_use")%>&nbsp;</td>
                                <td><%=rs("car_use_dept")%>&nbsp;</td>
                                <td><%=rs("buy_gubun")%>(<%=rs("rental_company")%>)&nbsp;</td>
							</tr>
                            <tr>
								<td><%=rs("last_check_date")%>&nbsp;</td>
								<td><%=formatnumber(rs("last_km"),0)%>&nbsp;</td>
                                <td><%=rs("car_status")%>&nbsp;</td>
                                <td><%=rs("end_date")%>&nbsp;</td>
                                <td colspan="2"><%=rs("car_comment")%>&nbsp;</td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
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
						for i = 0 to 10 
                        	if	drvuser_tab(i,1) <> "" then
						%>
							<tr>
								<td><%=drvuser_tab(i,1)%>&nbsp;</td>
								<td><%=drvuser_tab(i,2)%>(<%=drvuser_tab(i,3)%>)&nbsp;</td>
                                <td><%=drvuser_tab(i,4)%>&nbsp;</td>
                                <td><%=drvuser_tab(i,5)%>(<%=drvuser_tab(i,6)%>)&nbsp;</td>
                                <td><%=drvuser_tab(i,7)%>&nbsp;</td>
                                <td><%=drvuser_tab(i,8)%>&nbsp;</td>
							</tr>
						<%
							end if
						next
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
						for i = 0 to 10 
                        	if	insur_tab(i,1) <> "" then
						%>
							<tr>
								<td><%=insur_tab(i,1)%>&nbsp;</td>
								<td><%=formatnumber(insur_tab(i,2),0)%>&nbsp;</td>
                                <td><%=insur_tab(i,3)%>&nbsp;</td>
                                <td><%=insur_tab(i,4)%>&nbsp;</td>
                                <td><%=insur_tab(i,5)%>&nbsp;</td>
                                <td><%=insur_tab(i,6)%>&nbsp;</td>
                                <td><%=insur_tab(i,7)%>&nbsp;</td>
							</tr>
                            <tr>
								<td><%=insur_tab(i,8)%>&nbsp;</td>
								<td><%=insur_tab(i,9)%>&nbsp;</td>
                                <td><%=insur_tab(i,10)%>&nbsp;</td>
                                <td><%=insur_tab(i,11)%>&nbsp;</td>
                                <td><%=insur_tab(i,12)%>&nbsp;</td>
                                <td colspan="2"><%=insur_tab(i,13)%>&nbsp;</td>
							</tr>
						<%
							end if
						next
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
						for i = 0 to 10 
                        	if	as_tab(i,1) <> "" then
						%>
							<tr>
								<td><%=as_tab(i,1)%>&nbsp;</td>
								<td><%=formatnumber(insur_tab(i,4),0)%>&nbsp;(<%=as_tab(i,5)%>)</td>
                                <td class="left"><%=as_tab(i,2)%>&nbsp;</td>
                                <td class="left"><%=as_tab(i,3)%>&nbsp;</td>
                                <td><%=as_tab(i,7)%>(<%=as_tab(i,6)%>)&nbsp;</td>
							</tr>
						<%
							end if
						next
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
						for i = 0 to 10 
                        	if	pe_tab(i,1) <> "" then
						%>
							<tr>
								<td><%=pe_tab(i,1)%>&nbsp;</td>
								<td><%=formatnumber(pe_tab(i,2),0)%>&nbsp</td>
                                <td class="left"><%=pe_tab(i,3)%>&nbsp;</td>
                                <td class="left"><%=pe_tab(i,4)%>&nbsp;</td>
                                <td><%=pe_tab(i,5)%>&nbsp;</td>
                                <td><%=pe_tab(i,6)%>&nbsp;</td>
                                <td><%=pe_tab(i,7)%>(<%=pe_tab(i,8)%>)&nbsp;</td>
                                <td><%=pe_tab(i,9)%>(<%=pe_tab(i,10)%>)&nbsp;</td>
							</tr>
						<%
							end if
						next
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div align=right>
                    <br/>
						<a href="#" class="btnType04" onclick="javascript:goAction()" >닫기</a>&nbsp;&nbsp;
					</div>              
                    </td>
			      </tr>
			  </table>
         </div>	
	</form>
	  </div>				
	</body>
</html>

