<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim curr_date, title_line, rs_emp, arrTemp

curr_date = DateValue(Mid(CStr(Now()), 1, 10))
title_line = "개인 인사 정보"

objBuilder.Append "Call USP_PERSON_INSA_INFO('"&emp_no&"')"

Set rs_emp = DBConn.Execute(objBuilder.ToString())
'Call Rs_Open(rs_emp, DBConn, objBuilder.ToString())
objBuilder.Clear()

If Not rs_emp.EOF Then
	arrTemp = rs_emp.getRows()
End If

Call Rs_Close(rs_emp)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무관리</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}

			//개인 인사 정보 팝업[허정호_20210812]
			function insaCardPopView(id){
				var url = '/person/insa_individual_card00.asp';
				var param = '?emp_no='+id;
				var features = 'scrollbars=yes,width=1300,height=650';
				var win_name = '인사 카드';

				url += param;

				pop_Window(url, win_name, features);
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/person/insa_person_mg.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" />
							<col width="6%" />
							<col width="6%" />
							<col width="6%" />
							<col width="6%" />
							<col width="6%" />
							<col width="9%" />
							<col width="6%" />
							<col width="6%" />
							<col width="6%" />
							<col width="7%" />
							<col width="6%" />
							<col width="4%" />
							<col width="*"  />
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직위</th>
								<th scope="col">직책</th>
								<th scope="col">입사일</th>
								<th scope="col">소속</th>
								<th scope="col">최초입사일</th>
								<th scope="col">소속발령일</th>
								<th scope="col">승진일</th>
								<th scope="col">상주처</th>
								<th scope="col">생년월일</th>
								<th scope="col">구분</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim i
						Dim emp_name, emp_grade, emp_job, emp_position, emp_in_date
						Dim emp_first_date, emp_reside_place, emp_birthday, emp_type, emp_org_baldate
						Dim emp_grade_date, org_company, org_bonbu, org_saupbu, org_team

						If IsArray(arrTemp) Then
							For i = 0 To UBound(arrTemp, 2)
								emp_name = arrTemp(0, i)
								emp_grade = arrTemp(1, i)
								emp_job = arrTemp(2, i)
								emp_position = arrTemp(3, i)
								emp_in_date = arrTemp(4, i)
								emp_first_date = arrTemp(5, i)
								emp_reside_place = arrTemp(6, i)
								emp_birthday = arrTemp(7, i)
								emp_type = arrTemp(8, i)
								emp_org_baldate = arrTemp(9, i)
								emp_grade_date = arrTemp(10, i)
								org_company = arrTemp(11, i)
								org_bonbu = arrTemp(12, i)
								org_saupbu = arrTemp(13, i)
								org_team = arrTemp(14, i)

								If emp_org_baldate = "1900-01-01" Then
									emp_org_baldate = ""
								End If

								If emp_grade_date = "1900-01-01" Then
									emp_grade_date = ""
								End If
						%>
							<tr>
								<td class="first"><%=emp_no%></td>
								<td>
									<a href="#" onclick="insaCardPopView('<%=emp_no%>');"><%=emp_name%></a>
								</td>
								<td><%=emp_grade%>&nbsp;</td>
								<td><%=emp_job%>&nbsp;</td>
								<td><%=emp_position%>&nbsp;</td>
								<td><%=emp_in_date%>&nbsp;</td>
								<td><%=org_name%>&nbsp;</td>
								<td><%=emp_first_date%>&nbsp;</td>
								<td><%=emp_org_baldate%>&nbsp;</td>
								<td><%=emp_grade_date%>&nbsp;</td>
								<td><%=emp_reside_place%>&nbsp;</td>
								<td><%=emp_birthday%>&nbsp;</td>
								<td class="left"><%=emp_type%>&nbsp;</td>
								<td class="left">
								<%
								Call EmpOrgInSaupbuText(org_company, org_bonbu, org_saupbu, org_team)
								%>
								</td>
							</tr>
						<%
							Next
						Else
							Response.Write "<tr><td colspan='14' style='text-weight:bold;'>조회된 내역이 없습니다.</td></tr>"
						End If
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</div>
	</body>
</html>
<!--#include virtual="/common/inc_footer.asp"-->