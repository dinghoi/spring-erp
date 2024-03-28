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
Dim rs_emp, arrTemp, title_line

emp_no = f_Request("emp_no")

title_line = "개인 인사기록 카드"

objBuilder.Append "Call USP_PERSON_CARD_VIEW('"&emp_no&"')"
Set rs_emp = DBConn.Execute(objBuilder.ToString())
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
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			//인사 정보 팝업[허정호_20210818]
			function insaPopView(id, type){
				var url, win_name, features;
				var param = '?emp_no='+id+'&emp_name='+name;

				switch(type){
					case 'school':
						url = '/insa/insa_school_view.asp';
						win_name = '학력 사항';
						features = 'scrollbars=yes,width=800,height=400';
						break;
					case 'career':
						url = '/insa/insa_career_view.asp';
						win_name = '경력 사항';
						features = 'scrollbars=yes,width=800,height=400';
						break;
					case 'qual':
						url = '/insa/insa_qual_view.asp';
						win_name = '자격증 사항';
						features = 'scrollbars=yes,width=800,height=400';
						break;
					default :
						url = '/insa/insa_card01.asp';
						win_name = '인사기록 기타정보';
						features = 'scrollbars=yes,width=1300,height=750';
				}

				url += param;
				pop_Window(url, win_name, features);
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
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
                        <%
						Dim i, j, k

						Dim emp_name, emp_org_code, emp_jikgun, emp_jikmu, emp_person1
						Dim emp_person2, emp_position, emp_grade, emp_job, emp_image
						Dim emp_military_date1, emp_military_date2, emp_marry_date, emp_grade_date, emp_end_date
						Dim emp_org_baldate, emp_sawo_date, emp_first_date, emp_in_date, emp_tel_ddd
						Dim emp_tel_no1, emp_tel_no2, emp_hp_ddd, emp_hp_no1, emp_hp_no2
						Dim emp_ename, emp_sido, emp_gugun, emp_dong, emp_addr
						Dim emp_gunsok_date, emp_end_gisan, emp_email, emp_faith, emp_military_id
						Dim emp_military_grade, emp_military_comm

						Dim photo_image, sex_id, emp_sex

						'=== 개인 인사 기본 정보 =====================

						If IsArray(arrTemp) Then
							emp_name = arrTemp(0, 0)
							emp_org_code = arrTemp(1, 0)
							emp_jikgun = arrTemp(2, 0)
							emp_jikmu = arrTemp(3, 0)
							emp_person1 = arrTemp(4, 0)
							emp_person2 = arrTemp(5, 0)
							emp_position = arrTemp(6, 0)
							emp_grade = arrTemp(7, 0)
							emp_job = arrTemp(8, 0)
							emp_image = arrTemp(9, 0)
							emp_military_date1 = arrTemp(10, 0)
							emp_military_date2 = arrTemp(11, 0)
							emp_marry_date = arrTemp(12, 0)
							emp_grade_date = arrTemp(13, 0)
							emp_end_date = arrTemp(14, 0)
							emp_org_baldate = arrTemp(15, 0)
							emp_sawo_date = arrTemp(16, 0)
							emp_first_date = arrTemp(17, 0)
							emp_in_date = arrTemp(18, 0)
							emp_tel_ddd = arrTemp(19, 0)
							emp_tel_no1 = arrTemp(20, 0)
							emp_tel_no2 = arrTemp(21, 0)
							emp_hp_ddd = arrTemp(22, 0)
							emp_hp_no1 = arrTemp(23, 0)
							emp_hp_no2 = arrTemp(24, 0)
							emp_ename = arrTemp(25, 0)
							emp_sido = arrTemp(26, 0)
							emp_gugun = arrTemp(27, 0)
							emp_dong = arrTemp(28, 0)
							emp_addr = arrTemp(29, 0)
							emp_gunsok_date = arrTemp(30, 0)
							emp_end_gisan = arrTemp(31, 0)
							emp_email = arrTemp(32, 0)
							emp_faith = arrTemp(33, 0)
							emp_military_id = arrTemp(34, 0)
							emp_military_grade = arrTemp(35, 0)
							emp_military_comm = arrTemp(36, 0)

							If f_toString(emp_image, "") = "" Then
								photo_image = ""
							Else
								photo_image = "/emp_photo/" & emp_image
							End If

							If f_toString(emp_person2, "") <> "" Then
							   sex_id = Mid(CStr(emp_person2), 1, 1)

								If sex_id = "1" Then
									 emp_sex = "남"
								Else
									 emp_sex = "여"
								End If
							End If

							If emp_military_date1 = "1900-01-01" Then
								emp_military_date1 = ""
								emp_military_date2 = ""
							End If

							If emp_marry_date = "1900-01-01" Then
								emp_marry_date = ""
							End If

							If emp_grade_date = "1900-01-01" Then
								emp_grade_date = ""
							End If

							If emp_end_date = "1900-01-01" Then
								emp_end_date = ""
							End If

							If emp_org_baldate = "1900-01-01" Then
								emp_org_baldate = ""
							End If

							If emp_sawo_date = "1900-01-01" Then
								emp_sawo_date = ""
							End If
						End If
						%>
							<tr>
                                <td colspan="2" rowspan="4" class="first">
									<img src="<%=photo_image%>" width="110" height="120" alt="">
                                </td>
								<th>사원&nbsp;&nbsp;번호</th>
                                <td class="left"><%=emp_no%></td>
								<th>소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
								<td colspan="2" class="left"><%=emp_org_code%>)<%=org_name%>&nbsp;</td>
                                <th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;무</th>
								<td class="left"><%=emp_jikgun%>-<%=emp_jikmu%>&nbsp;</td>
                                <th>주민번호</th>
								<td colspan="2" class="left"><%=emp_person1%>-<%=emp_person2%>&nbsp;&nbsp;(<%=emp_sex%>)</td>
                 			</tr>
							<tr>
								<th>성명(한글)</th>
                                <td class="left"><%=emp_name%>&nbsp;</td>
								<th>성명(영문)</th>
								<td colspan="2" class="left"><%=emp_ename%>&nbsp;</td>
                                <th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;책</th>
                                <td class="left"><%=emp_position%>&nbsp;</td>
								<th>직급(위)/승진일</th>
								<td colspan="2" class="left">(<%=emp_grade%>)&nbsp;<%=emp_job%>&nbsp;/&nbsp;<%=emp_grade_date%></td>
                 			</tr>
							<tr>
                                <th>최초입사일</th>
                                <td class="left"><%=emp_first_date%></td>
                                <th>입&nbsp;&nbsp;&nbsp;사&nbsp;&nbsp;&nbsp;일</th>
                                <td class="left"><%=emp_in_date%>&nbsp;</td>
                                <th>전화번호</th>
								<td class="left"><%=emp_tel_ddd%>-<%=emp_tel_no1%>-<%=emp_tel_no2%>&nbsp;</td>
								<th>주소(현)</th>
								<td colspan="3" class="left"><%=emp_sido%>&nbsp;<%=emp_gugun%>&nbsp;<%=emp_dong%>&nbsp;<%=emp_addr%></td>
                            </tr>
                            <tr>
                                <th>근속기산일</th>
                                <td class="left"><%=emp_gunsok_date%>&nbsp;</td>
                                <th>퇴직기산일</th>
                                <td class="left"><%=emp_end_gisan%>&nbsp;</td>
                                <th>휴대폰번호</th>
								<td class="left"><%=emp_hp_ddd%>-<%=emp_hp_no1%>-<%=emp_hp_no2%>&nbsp;</td>
                                <th>이메일 주소</th>
								<td colspan="3" class="left"><%=emp_email%>@k-one.co.kr&nbsp;</td>
							</tr>
                            <tr>
                                <th colspan="10" class="left">■ 학력 사항 ■</th>
                                <td colspan="2" class="right">&nbsp;
                                <a href="#" class="btnType03" onClick="insaPopView('<%=emp_no%>', 'school');">☞ 학력 더보기</a>
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
							<%
							Dim arrSch, rs_sch
							Dim sch_start_date, sch_end_date, sch_school_name, sch_dept, sch_major
							Dim sch_degree, sch_finish, sch_sub_major

							objBuilder.Append "Call USP_PERSON_CARD_SCHOOL_INFO('"&emp_no&"')"
							Call Rs_Open(rs_sch, DBConn, objBuilder.ToString())
							objBuilder.Clear()

							If Not rs_sch.EOF Then
								arrSch = rs_sch.getRows()
							End If

							Call Rs_Close(rs_sch)

							If IsArray(arrSch) Then
								For i = 0 To UBound(arrSch, 2)
									sch_start_date = arrSch(0, i)
									sch_end_date = arrSch(1, i)
									sch_school_name = arrSch(2, i)
									sch_dept = arrSch(3, i)
									sch_major = arrSch(4, i)
									sch_sub_major = arrSch(5, i)
									sch_degree = arrSch(6, i)
									sch_finish = arrSch(7, i)
							%>
							<tr>
                				<td colspan="3" class="left"><%=sch_start_date%>&nbsp;~&nbsp;<%=sch_end_date%></td>
                                <td colspan="2" class="left"><%=sch_school_name%>&nbsp;</td>
                                <td colspan="2" class="left"><%=sch_dept%>&nbsp;</td>
                                <td colspan="2" class="left"><%=sch_major%>&nbsp;</td>
                                <td class="left"><%=sch_sub_major%>&nbsp;</td>
                                <td class="left"><%=sch_degree%>&nbsp;</td>
                                <td class="left"><%=sch_finish%>&nbsp;</td>
                            </tr>
							<%
								Next
							Else
							%>
							<tr>
                				<td colspan="12" style="font-weight:bold;">해당 내역이 없습니다.</td>
                            </tr>
							<%End If%>
                            <tr>
                                <th colspan="10" class="left">■ 이전 경력 사항 ■</th>
                                <td colspan="2" class="right">&nbsp;
                                <a href="#" class="btnType03" onClick="insaPopView('<%=emp_no%>', 'career');">☞ 경력 더보기</a>
                                </td>
                            </tr>
                            <tr>
                                <th colspan="3">재직기간</th>
                                <th colspan="2">회사명</th>
                                <th colspan="2">부  서</th>
                                <th>직위</th>
                                <th colspan="4">담당업무</th>
                            </tr>
							<%
							Dim rs_career, arrCareer
							Dim career_join_date, career_end_date, career_office, career_dept, career_position
							Dim career_task

							objBuilder.Append "Call USP_PERSON_CARD_CAREER_INFO('"&emp_no&"')"
							Call Rs_Open(rs_career, DBConn, objBuilder.ToString())
							objBuilder.Clear()

							If Not rs_career.EOF Then
								arrCareer = rs_career.getRows()
							End If

							Call Rs_Close(rs_career)

							If IsArray(arrCareer) Then
								For i = 0 To UBound(arrCareer, 2)
									career_join_date = arrCareer(0, i)
									career_end_date = arrCareer(1, i)
									career_office = arrCareer(2, i)
									career_dept = arrCareer(3, i)
									career_position = arrCareer(4, i)
									career_task = arrCareer(5, i)
							%>
                            <tr>
                                <td colspan="3" class="left"><%=career_join_date%>&nbsp;~&nbsp;<%=career_end_date%></td>
                                <td colspan="2" class="left"><%=career_office%>&nbsp;</td>
                                <td colspan="2" class="left"><%=career_dept%>&nbsp;</td>
                                <td colspan="1" class="left"><%=career_position%>&nbsp;</td>
                                <td colspan="4" class="left"><%=career_task%>&nbsp;</td>
                            </tr>
							<%
								Next
							Else
							%>
							<tr>
                				<td colspan="12" style="font-weight:bold;">해당 내역이 없습니다.</td>
                            </tr>
							<%End If%>
                            <tr>
                                <th colspan="10" class="left">■ 자격증 사항 ■</th>
                                <td colspan="2" class="right">&nbsp;
                                <a href="#" class="btnType03" onClick="insaPopView('<%=emp_no%>', 'qual');">☞ 자격 더보기</a>
                                </td>
                            </tr>
                            <tr>
                                <th colspan="3">자격증 종목</th>
                                <th>등급</th>
                                <th colspan="2">합격년월일</th>
                                <th colspan="2">발급 기관명</th>
                                <th colspan="4">자격 등록번호</th>
                            </tr>
							<%
							Dim rs_qual, arrQual
							Dim qual_type, qual_grade, qual_pass_date, qual_org, qual_no

							objBuilder.Append "Call USP_PERSON_CARD_QUAL_INFO('"&emp_no&"')"
							Call Rs_Open(rs_qual, DBConn, objBuilder.ToString())
							objBuilder.Clear()

							If Not rs_qual.EOF Then
								arrQual = rs_qual.getRows()
							End If

							Call Rs_Close(rs_qual)

							If IsArray(arrQual) Then
								For i = 0 To UBound(arrQual, 2)
									qual_type = arrQual(0, i)
									qual_grade = arrQual(1, i)
									qual_pass_date = arrQual(2, i)
									qual_org = arrQual(3, i)
									qual_no = arrQual(4, i)
							%>
                            <tr>
                                <td colspan="3" class="left"><%=qual_type%>&nbsp;</td>
                                <td class="left"><%=qual_grade%>&nbsp;</td>
                                <td colspan="2" class="left"><%=qual_pass_date%>&nbsp;</td>
                                <td colspan="2" class="left"><%=qual_org%>&nbsp;</td>
                                <td colspan="4" class="left"><%=qual_no%>&nbsp;</td>
                            </tr>
							<%
								Next
							Else
							%>
							<tr>
                				<td colspan="12" style="font-weight:bold;">해당 내역이 없습니다.</td>
                            </tr>
							<%End If%>
                            <tr>
                                <th>병역 복무기간</th>
                                <td colspan="2" class="left"><%=Mid(emp_military_date1, 1, 7)%>~<%=Mid(emp_military_date2, 1, 7)%>&nbsp;</td>
                                <th>병역유형/계급</th>
                                <td class="left"><%=emp_military_id%> - <%=emp_military_grade%>&nbsp;</td>
                                <th>면제사유</th>
								<td colspan="2" class="left"><%=emp_military_comm%>&nbsp;</td>
                                <th>결혼기념일</th>
                                <td class="left"><%=emp_marry_date%>&nbsp;</td>
                                <th>종교</th>
                                <td class="left"><%=emp_faith%>&nbsp;</td>
							</tr>

                      </tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<tr></tr>
					<tr>
						<td>&nbsp;</td>
						<td style="width:21%;">
							<div class="btnCenter" style="float:left;">
								<a href="#" class="btnType04" onClick="insaPopView('<%=emp_no%>', '');">☞ 인사기록 기타정보</a>
								<span class="btnType01"><input type="button" value="닫기" onclick="close_win();"></span>
							</div>
						</td>
					</tr>
				</table>
			</div>
		</div>
	</body>
</html>
<!--#include virtual="/common/inc_footer.asp"-->