<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim sch_tab(10,10)
Dim car_tab(20,10)
Dim qul_tab(20,10)

Dim acpt_emp_no, curr_date, be_pg1
Dim i, k, j
Dim rs, rs_sch, rs_car, rs_qul, title_line
Dim photo_image, emp_person2, sex_id, emp_sex
Dim emp_military_date1, emp_military_date2, emp_marry_date
Dim emp_email, date_sw, acpt_user

emp_no = request("emp_no")

acpt_emp_no = user_id
curr_date = Mid(CStr(Now()), 1, 10)
be_pg1 = "/insa/insa_card00.asp"

objBuilder.Append "SELECT emtt.emp_image, emtt.emp_person1, emtt.emp_person2, emtt.emp_military_date1, "
objBuilder.Append "	emtt.emp_military_date2, emtt.emp_marry_date, emtt.emp_email, emtt.emp_no, emtt.emp_org_code, "
objBuilder.Append "	emtt.emp_org_name, emtt.emp_jikgun, emtt.emp_jikmu, emtt.emp_name, emtt.emp_position, "
objBuilder.Append "	emtt.emp_grade, emtt.emp_job, emtt.emp_grade_date, emtt.emp_first_date, emtt.emp_in_date, "
objBuilder.Append "	emtt.emp_tel_ddd, emtt.emp_tel_no1, emtt.emp_tel_no2, emtt.emp_sido, emtt.emp_gugun, "
objBuilder.Append "	emtt.emp_dong, emtt.emp_addr, emtt.emp_gunsok_date, emtt.emp_end_gisan, emtt.emp_hp_ddd, "
objBuilder.Append "	emtt.emp_hp_no1, emtt.emp_hp_no2, emtt.emp_military_id, emtt.emp_military_grade, "
objBuilder.Append "	emtt.emp_military_comm, emtt.emp_faith, emtt.emp_ename, "
objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_team, eomt.org_name "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emtt.emp_no = '"&emp_no&"' "

Set rs = DbConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rs.EOF Or Not rs.BOF Then
    If rs("emp_image") = "" Or IsNull(rs("emp_image")) Then
		photo_image = ""
	Else
		photo_image = "/emp_photo/" & rs("emp_image")
    End If

    emp_person2 = rs("emp_person2")

	If emp_person2 <> "" Then
		sex_id = Mid(CStr(emp_person2), 1, 1)

		If sex_id = "1" Then
			emp_sex = "남"
		Else
			emp_sex = "여"
		End If
	End If

    If rs("emp_military_date1") = "1900-01-01" Then
           emp_military_date1 = ""
           emp_military_date2 = ""
    Else
           emp_military_date1 = rs("emp_military_date1")
           emp_military_date2 = rs("emp_military_date2")
    End If

    If rs("emp_marry_date") = "1900-01-01" Then
           emp_marry_date = ""
    Else
     	   emp_marry_date = rs("emp_marry_date")
    End If

	'학력사항 db
	For i = 0 To 10
	'	com_tab(i) = ""
	'	com_sum(i) = 0
		For j = 0 To 10
			sch_tab(i,j) = ""
	'		com_in(i,j) = 0
		Next
	Next

	k = 0

	objBuilder.Append "SELECT sch_start_date, sch_end_date, sch_school_name, sch_dept, sch_major, "
	objBuilder.Append "	sch_sub_major, sch_degree, sch_finish "
	objBuilder.Append "FROM emp_school "
	objBuilder.Append "WHERE sch_empno = '"&emp_no&"' "
	objBuilder.Append "ORDER BY sch_empno, sch_seq ASC "

	Set rs_sch = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	While Not rs_sch.EOF
		k = k + 1

		sch_tab(k, 1) = rs_sch("sch_start_date")
		sch_tab(k, 2) = rs_sch("sch_end_date")
		sch_tab(k, 3) = rs_sch("sch_school_name")
		sch_tab(k, 4) = rs_sch("sch_dept")
		sch_tab(k, 5) = rs_sch("sch_major")
		sch_tab(k, 6) = rs_sch("sch_sub_major")
		sch_tab(k, 7) = rs_sch("sch_degree")
		sch_tab(k, 8) = rs_sch("sch_finish")

		rs_sch.MoveNext()
	Wend
	rs_sch.Close() : Set rs_sch = Nothing

	'경력사항 db
	For i = 0 To 20
		For j = 0 To 10
			car_tab(i, j) = ""
		Next
	Next

	k = 0

	objBuilder.Append "SELECT career_join_date, career_end_date, career_office, "
	objBuilder.Append "	career_dept, career_position, career_task "
	objBuilder.Append "FROM emp_career "
	objBuilder.Append "WHERE career_empno = '"&emp_no&"' "
	objBuilder.Append "ORDER BY career_empno, career_seq ASC "

	Set rs_car = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	While Not rs_car.EOF
		k = k + 1

		car_tab(k, 1) = rs_car("career_join_date")
		car_tab(k, 2) = rs_car("career_end_date")
		car_tab(k, 3) = rs_car("career_office")
		car_tab(k, 4) = rs_car("career_dept")
		car_tab(k, 5) = rs_car("career_position")
		car_tab(k, 6) = rs_car("career_task")

		rs_car.MoveNext()
	Wend
	rs_car.Close() : Set rs_car = Nothing

	'자격사항 db
	For i = 0 To 20
		For j = 0 To 10
			qul_tab(i, j) = ""
		Next
	Next

	k = 0

	objBuilder.Append "SELECT qual_type, qual_grade, qual_pass_date, qual_org, qual_no "
	objBuilder.Append "FROM emp_qual "
	objBuilder.Append "WHERE qual_empno = '"&emp_no&"' "
	objBuilder.Append "ORDER BY qual_empno, qual_seq ASC "

	Set rs_qul = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	While Not rs_qul.EOF
		k = k + 1

		qul_tab(k, 1) = rs_qul("qual_type")
		qul_tab(k, 2) = rs_qul("qual_grade")
		qul_tab(k, 3) = rs_qul("qual_pass_date")
		qul_tab(k, 4) = rs_qul("qual_org")
		qul_tab(k, 5) = rs_qul("qual_no")

		rs_qul.MoveNext()
	Wend
	rs_qul.Close() : Set rs_qul = Nothing
End If

title_line = "인사 기록 카드"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html lang="ko">
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

	<script type="text/javascript">
		function getPageCode(){
			return "1 1";
		}

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
			if(document.frm.condi.value == ""){
				alert ("소속을 선택하시기 바랍니다");
				return false;
			}
			return true;
		}

		//인사기록카드 출력 팝업[허정호_20210811]
		function insaCardPopView(id){
			var url = '/insa/insa_card_print.asp';
			var pop_name = '인사 기록 카드';
			var param = '?emp_no='+id;
			var features = 'scrollbars=yes,width=750,height=600';

			url += param;

			pop_Window(url, pop_name, features);
		}
	</script>
</head>
<body>
	<div id="wrap">
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="insa_card00.asp" method="post" name="frm">
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableWrite">
					<colgroup>
						<col width="9%" >
						<col width="1%" >
						<col width="9%" >
						<col width="9%" >
						<col width="9%" >
						<col width="9%" >
						<col width="9%" >
						<col width="9%" >
						<col width="9%" >
						<col width="9%" >
						<col width="9%" >
						<col width="9%" >
					</colgroup>
					<tbody>
					<% If Not rs.EOF Or Not rs.BOF Then %>
						<tr>
							<%
							emp_email = rs("emp_email") & "@k-one.co.kr"
							%>
							<td colspan="2" rowspan="4" class="first">
							<img src="<%=photo_image%>" width="110" height="120" alt="">
							</td>
							<th>사원&nbsp;&nbsp;번호</th>
							<td class="left"><%=rs("emp_no")%></td>
							<th>소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
							<td colspan="2" class="left"><%=rs("emp_org_code")%>)&nbsp;<%'=rs("emp_org_name")%><%=rs("org_name")%>&nbsp;</td>
							<th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;무</th>
							<td class="left"><%=rs("emp_jikgun")%>-<%=rs("emp_jikmu")%>&nbsp;</td>
							<th>주민번호</th>
							<td colspan="2" class="left"><%=rs("emp_person1")%>-<%=rs("emp_person2")%>&nbsp;&nbsp;(<%=emp_sex%>)</td>
						</tr>
						<tr>
							<th>성명(한글)</th>
							<td class="left"><%=rs("emp_name")%>&nbsp;</td>
							<th>성명(영문)</th>
							<td colspan="2" class="left"><%=rs("emp_ename")%>&nbsp;</td>
							<th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;책</th>
							<td class="left"><%=rs("emp_position")%>&nbsp;</td>
							<th>직급(위)/승진일</th>
							<td colspan="2" class="left">(<%=rs("emp_grade")%>)&nbsp;<%=rs("emp_job")%>&nbsp;/&nbsp;<%=rs("emp_grade_date")%></td>
						</tr>
						<tr>
							<th>최초입사일</th>
							<td class="left"><%=rs("emp_first_date")%></td>
							<th>입&nbsp;&nbsp;&nbsp;사&nbsp;&nbsp;&nbsp;일</th>
							<td class="left"><%=rs("emp_in_date")%>&nbsp;</td>
							<th>전화번호</th>
							<td class="left"><%=rs("emp_tel_ddd")%>-<%=rs("emp_tel_no1")%>-<%=rs("emp_tel_no2")%>&nbsp;</td>
							<th>주소(현)</th>
							<td colspan="3" class="left"><%=rs("emp_sido")%>&nbsp;<%=rs("emp_gugun")%>&nbsp;<%=rs("emp_dong")%>&nbsp;<%=rs("emp_addr")%></td>
						</tr>
						<tr>
							<th>근속기산일</th>
							<td class="left"><%=rs("emp_gunsok_date")%>&nbsp;</td>
							<th>퇴직기산일</th>
							<td class="left"><%=rs("emp_end_gisan")%>&nbsp;</td>
							<th>휴대폰번호</th>
							<td class="left"><%=rs("emp_hp_ddd")%>-<%=rs("emp_hp_no1")%>-<%=rs("emp_hp_no2")%>&nbsp;</td>
							<th>이메일 주소</th>
							<td colspan="3" class="left"><%=emp_email%>&nbsp;</td>
						</tr>
						<tr>
							<th colspan="10" class="left">■ 학력 사항 ■</th>
							<td colspan="2" class="right">&nbsp;
							<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_school_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>','schoolview','scrollbars=yes,width=800,height=400')">☞ 학력 더보기</a>
							</td>
						</tr>
						<tr>
							<th colspan="3">기간</th>
							<th colspan="2">학교명</th>
							<th colspan="2">학과</th>
							<th colspan="2">전공</th>
							<th>부전공</th>
							<th>학위</th>
							<th>졸업</th>
						</tr>
							<td colspan="3" class="left"><%=sch_tab(1, 1)%>&nbsp;~&nbsp;<%=sch_tab(1, 2)%></td>
							<td colspan="2" class="left"><%=sch_tab(1, 3)%>&nbsp;</td>
							<td colspan="2" class="left"><%=sch_tab(1, 4)%>&nbsp;</td>
							<td colspan="2" class="left"><%=sch_tab(1, 5)%>&nbsp;</td>
							<td class="left"><%=sch_tab(1, 6)%>&nbsp;</td>
							<td class="left"><%=sch_tab(1, 7)%>&nbsp;</td>
							<td class="left"><%=sch_tab(1, 8)%>&nbsp;</td>
						 </tr>
						</tr>
							<td colspan="3" class="left"><%=sch_tab(2, 1)%>&nbsp;~&nbsp;<%=sch_tab(2, 2)%></td>
							<td colspan="2" class="left"><%=sch_tab(2, 3)%>&nbsp;</td>
							<td colspan="2" class="left"><%=sch_tab(2, 4)%>&nbsp;</td>
							<td colspan="2" class="left"><%=sch_tab(2, 5)%>&nbsp;</td>
							<td class="left"><%=sch_tab(2, 6)%>&nbsp;</td>
							<td class="left"><%=sch_tab(2, 7)%>&nbsp;</td>
							<td class="left"><%=sch_tab(2, 8)%>&nbsp;</td>
						 </tr>
						<tr>
							<th colspan="10" class="left">■ 이전 경력 사항 ■</th>
							<td colspan="2" class="right">&nbsp;
							<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_career_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>','careerview','scrollbars=yes,width=800,height=400')">☞ 경력 더보기</a>
							</td>
						</tr>
						<tr>
							<th colspan="3">재직기간</th>
							<th colspan="2">회사명</th>
							<th colspan="2">부  서</th>
							<th>직위</th>
							<th colspan="4">담당업무</th>
						</tr>
						<tr>
							<td colspan="3" class="left"><%=car_tab(1, 1)%>&nbsp;~&nbsp;<%=car_tab(1, 2)%></td>
							<td colspan="2" class="left"><%=car_tab(1, 3)%>&nbsp;</td>
							<td colspan="2" class="left"><%=car_tab(1, 4)%>&nbsp;</td>
							<td colspan="1" class="left"><%=car_tab(1, 5)%>&nbsp;</td>
							<td colspan="4" class="left"><%=car_tab(1, 6)%>&nbsp;</td>
						 </tr>
						<tr>
							<td colspan="3" class="left"><%=car_tab(2, 1)%>&nbsp;~&nbsp;<%=car_tab(2, 2)%></td>
							<td colspan="2" class="left"><%=car_tab(2, 3)%>&nbsp;</td>
							<td colspan="2" class="left"><%=car_tab(2, 4)%>&nbsp;</td>
							<td colspan="1" class="left"><%=car_tab(2, 5)%>&nbsp;</td>
							<td colspan="4" class="left"><%=car_tab(2, 6)%>&nbsp;</td>
						 </tr>
						 <tr>
							<th colspan="10" class="left">■ 자격증 사항 ■</th>
							<td colspan="2" class="right">&nbsp;
							<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_qual_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>','qualview','scrollbars=yes,width=800,height=400')">☞ 자격 더보기</a>
							</td>
						</tr>
						<tr>
							<th colspan="3">자격증 종목</th>
							<th>등급</th>
							<th colspan="2">합격년월일</th>
							<th colspan="2">발급 기관명</th>
							<th colspan="4">자격 등록번호</th>
						</tr>
						<tr>
							<td colspan="3" class="left"><%=qul_tab(1, 1)%>&nbsp;</td>
							<td class="left"><%=qul_tab(1, 2)%>&nbsp;</td>
							<td colspan="2" class="left"><%=qul_tab(1, 3)%>&nbsp;</td>
							<td colspan="2" class="left"><%=qul_tab(1, 4)%>&nbsp;</td>
							<td colspan="4" class="left"><%=qul_tab(1, 5)%>&nbsp;</td>
						 </tr>
						<tr>
							<td colspan="3" class="left"><%=qul_tab(2, 1)%>&nbsp;</td>
							<td class="left"><%=qul_tab(2, 2)%>&nbsp;</td>
							<td colspan="2" class="left"><%=qul_tab(2, 3)%>&nbsp;</td>
							<td colspan="2" class="left"><%=qul_tab(2, 4)%>&nbsp;</td>
							<td colspan="4" class="left"><%=qul_tab(2, 5)%>&nbsp;</td>
						 </tr>
						<tr>
							<td colspan="3" class="left"><%=qul_tab(3, 1)%>&nbsp;</td>
							<td class="left"><%=qul_tab(3, 2)%>&nbsp;</td>
							<td colspan="2" class="left"><%=qul_tab(3, 3)%>&nbsp;</td>
							<td colspan="2" class="left"><%=qul_tab(3, 4)%>&nbsp;</td>
							<td colspan="4" class="left"><%=qul_tab(3, 5)%>&nbsp;</td>
						 </tr>
						<tr>
							<th>병역 복무기간</th>
							<td colspan="2" class="left"><%=Mid(emp_military_date1, 1, 7)%>~<%=Mid(emp_military_date2, 1, 7)%>&nbsp;</td>
							<th>병역유형/계급</th>
							<td class="left"><%=rs("emp_military_id")%> - <%=rs("emp_military_grade")%>&nbsp;</td>
							<th>면제사유</th>
							<td colspan="2" class="left"><%=rs("emp_military_comm")%>&nbsp;</td>
							<th>결혼기념일</th>
							<td class="left"><%=emp_marry_date%>&nbsp;</td>
							<th>종교</th>
							<td class="left"><%=rs("emp_faith")%>&nbsp;</td>
						</tr>
				  <% End If %>
				  </tbody>
				</table>
			</div>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			  <tr>
				<td width="40%">
					<div class="btnCenter">
						<a href="#" class="btnType04" onClick="insaCardPopView('<%=rs("emp_no")%>');">인사기록카드 출력</a>
					<% If SysAdminYn = "Y" Then'작업된 코드 페이지 없음[허정호_20220322] %>
						<a href="/insa_excel_card_print.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>" class="btnType04">엑셀다운로드(미작업)</a>
					<% End If %>
					</div>
				</td>
				<td>
					<div class="btnCenter">
						<span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
					</div>
				</td>
				<td width="20%">
					<div class="btnCenter">
						<a href="#" class="btnType04" onClick="pop_Window('/insa/insa_card01.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg1%>&acpt_user=<%=acpt_user%>','emp_card1_pop','scrollbars=yes,width=1250,height=750')">☞ 인사기록 기타정보</a>
					</div>
				</td>
			  </tr>
			  </table>
		</form>
	</div>
</div>
</body>
</html>
<%
rs.Close : Set rs = Nothing
DBConn.Close() : Set DBConn = Nothing
%>