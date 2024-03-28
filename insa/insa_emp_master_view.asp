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
Dim u_type, view_condi, title_line
Dim rsEmp
Dim emp_name, emp_ename, emp_type, emp_sex, emp_person1, emp_person2
Dim sex_id, emp_image, att_file, emp_first_date, emp_in_date, emp_gunsok_date
Dim emp_yuncha_date, emp_end_gisan, emp_end_date, emp_bonbu
Dim emp_saupbu, emp_team, emp_org_code, emp_org_name, emp_org_baldate, emp_stay_code
Dim emp_stay_name, emp_reside_place, emp_reside_company, emp_grade, emp_grade_date
Dim emp_job, emp_position, emp_jikgun, emp_jikmu, emp_birthday, emp_birthday_id
Dim emp_zipcode, emp_sido, emp_gugun, emp_dong, emp_addr, emp_tel_ddd
Dim emp_tel_no1, emp_tel_no2, emp_hp_ddd, emp_hp_no1, emp_hp_no2
Dim emp_email, emp_military_id, emp_military_date1, emp_military_date2
Dim emp_military_grade, emp_military_comm, emp_hobby, emp_faith, emp_last_edu
Dim emp_marry_date, emp_disabled, emp_disab_grade, emp_sawo_id, emp_sawo_date
Dim emp_emergency_tel, emp_nation_code, emp_extension_no, cost_group, cost_center
Dim emp_pay_id, emp_reg_date, emp_reg_user, emp_mod_date, emp_mod_user
Dim photo_image, rsMem, grade, dz_id

u_type = Request("u_type")
emp_no = Request("emp_no")
view_condi = Request("view_condi")

title_line = " 인사기본사항 조회 "

objBuilder.Append "SELECT emtt.emp_name, emtt.emp_ename, emtt.emp_type, emtt.emp_sex, emtt.emp_person1, emtt.emp_person2, "
objBuilder.Append "	emtt.emp_image, emtt.emp_first_date, emtt.emp_in_date, emtt.emp_gunsok_date, emtt.emp_yuncha_date, "
objBuilder.Append "	emtt.emp_end_gisan, emtt.emp_end_date, emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, "
objBuilder.Append "	emtt.emp_org_code, emtt.emp_org_name, emtt.emp_org_baldate, emtt.emp_stay_code, emtt.emp_stay_name, "
objBuilder.Append "	emtt.emp_reside_place, emtt.emp_reside_company, emtt.emp_grade, emtt.emp_grade_date, emtt.emp_job, "
objBuilder.Append "	emtt.emp_position, emtt.emp_jikgun, emtt.emp_jikmu, emtt.emp_birthday, emtt.emp_birthday_id, "
objBuilder.Append "	emtt.emp_zipcode, emtt.emp_sido, emtt.emp_gugun, emtt.emp_dong, emp_addr, "
objBuilder.Append "	emtt.emp_tel_ddd, emtt.emp_tel_no1, emtt.emp_tel_no2, emtt.emp_hp_ddd, emtt.emp_hp_no1, emtt.emp_hp_no2, "
objBuilder.Append "	emtt.emp_email, emtt.emp_military_id, emtt.emp_military_date1, emtt.emp_military_date2, "
objBuilder.Append "	emtt.emp_military_grade, emtt.emp_military_comm, emtt.emp_hobby, emtt.emp_faith, emtt.emp_last_edu, "
objBuilder.Append "	emtt.emp_marry_date, emtt.emp_disabled, emtt.emp_disab_grade, emtt.emp_sawo_id, emtt.emp_sawo_date, "
objBuilder.Append "	emtt.emp_emergency_tel, emtt.emp_nation_code, emtt.emp_extension_no, emtt.cost_group, emtt.cost_center, "
objBuilder.Append "	emtt.emp_pay_id, emtt.emp_reg_date, emtt.emp_reg_user, emtt.emp_mod_date, emtt.emp_mod_user, "
objBuilder.Append "	eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, eomt.org_reside_place, "
objBuilder.Append "	dpit.dz_id "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "LEFT OUTER JOIN dz_pay_info AS dpit ON emtt.emp_no = dpit.emp_no "
objBuilder.Append "WHERE emtt.emp_no = '"&emp_no&"' "

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

emp_name = rsEmp("emp_name")
emp_ename = rsEmp("emp_ename")
emp_type = rsEmp("emp_type")
emp_sex = rsEmp("emp_sex")
emp_person1 = rsEmp("emp_person1")

emp_person2 = rsEmp("emp_person2")
If emp_person2 <> "" Then
	sex_id = mid(cstr(emp_person2),1,1)

	If sex_id = "1" Then
		emp_sex = "남"
	Else
		emp_sex = "여"
   	End If
End If

emp_image = rsEmp("emp_image")
att_file = rsEmp("emp_image")
emp_first_date = rsEmp("emp_first_date")
emp_in_date = rsEmp("emp_in_date")
emp_gunsok_date = rsEmp("emp_gunsok_date")
emp_yuncha_date = rsEmp("emp_yuncha_date")
emp_end_gisan = rsEmp("emp_end_gisan")
emp_end_date = rsEmp("emp_end_date")

emp_company = rsEmp("org_company")
emp_bonbu = rsEmp("org_bonbu")
emp_saupbu = rsEmp("org_saupbu")
emp_team = rsEmp("org_team")
emp_org_code = rsEmp("emp_org_code")
emp_org_name = rsEmp("org_name")

emp_org_baldate = rsEmp("emp_org_baldate")
emp_stay_code = rsEmp("emp_stay_code")
emp_stay_name = rsEmp("emp_stay_name")
emp_reside_place = rsEmp("emp_reside_place")
emp_reside_company = rsEmp("emp_reside_company")
emp_grade = rsEmp("emp_grade")
emp_grade_date = rsEmp("emp_grade_date")
emp_job = rsEmp("emp_job")
emp_position = rsEmp("emp_position")
emp_jikgun = rsEmp("emp_jikgun")
emp_jikmu = rsEmp("emp_jikmu")
emp_birthday = rsEmp("emp_birthday")
emp_birthday_id = rsEmp("emp_birthday_id")
emp_zipcode = rsEmp("emp_zipcode")
emp_sido = rsEmp("emp_sido")
emp_gugun = rsEmp("emp_gugun")
emp_dong = rsEmp("emp_dong")
emp_addr = rsEmp("emp_addr")
emp_tel_ddd = rsEmp("emp_tel_ddd")
emp_tel_no1 = rsEmp("emp_tel_no1")
emp_tel_no2 = rsEmp("emp_tel_no2")
emp_hp_ddd = rsEmp("emp_hp_ddd")
emp_hp_no1 = rsEmp("emp_hp_no1")
emp_hp_no2 = rsEmp("emp_hp_no2")
emp_email = rsEmp("emp_email")
emp_military_id = rsEmp("emp_military_id")
emp_military_date1 = rsEmp("emp_military_date1")
emp_military_date2 = rsEmp("emp_military_date2")
emp_military_grade = rsEmp("emp_military_grade")
emp_military_comm = rsEmp("emp_military_comm")
emp_hobby = rsEmp("emp_hobby")
emp_faith = rsEmp("emp_faith")
emp_last_edu = rsEmp("emp_last_edu")
emp_marry_date = rsEmp("emp_marry_date")
emp_disabled = rsEmp("emp_disabled")
emp_disab_grade = rsEmp("emp_disab_grade")

emp_sawo_id = rsEmp("emp_sawo_id")
If f_toString(rsEmp("emp_sawo_id"), "") = "" Then
   emp_sawo_id = "N"
End If

emp_sawo_date = rsEmp("emp_sawo_date")
emp_emergency_tel = rsEmp("emp_emergency_tel")
emp_nation_code = rsEmp("emp_nation_code")
emp_extension_no = rsEmp("emp_extension_no")
cost_group = rsEmp("cost_group")
cost_center = rsEmp("cost_center")
emp_pay_id = rsEmp("emp_pay_id")
'   end_date = mid(cstr(now()),1,10)
emp_reg_date = rsEmp("emp_reg_date")
emp_reg_user = rsEmp("emp_reg_user")
emp_mod_date = rsEmp("emp_mod_date")
emp_mod_user = rsEmp("emp_mod_user")
photo_image = "/emp_photo/"&rsEmp("emp_image")
att_file = rsEmp("emp_image")
dz_id = rsEmp("dz_id")

If emp_pay_id = "5" Then
	emp_pay_id = "안함"
Else
	emp_pay_id = "지급"
End If

If rsEmp("emp_military_date1") = "1900-01-01" Then
	emp_military_date1 = ""
	emp_military_date2 = ""
End If

If rsEmp("emp_birthday") = "1900-01-01" Then
	emp_birthday = ""
End If

If rsEmp("emp_marry_date") = "1900-01-01" Then
	emp_marry_date = ""
End If

If rsEmp("emp_grade_date") = "1900-01-01" Then
	emp_grade_date = ""
End If

If rsEmp("emp_end_date") = "1900-01-01" Then
	emp_end_date = ""
End If

If rsEmp("emp_org_baldate") = "1900-01-01" Then
	emp_org_baldate = ""
End If

If rsEmp("emp_sawo_date") = "1900-01-01" Then
	emp_sawo_date = ""
End If

rsEmp.close() : Set rsEmp = Nothing

objBuilder.Append "SELECT mg_group, grade FROM memb WHERE user_id = '"&emp_no&"' "

Set rsMem = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsMem.EOF Then
	mg_group = rsMem("mg_group")
	grade    = rsMem("grade")
Else
	mg_group = "1"
	grade    = ""
End If
rsMem.close() : Set rsMem = Nothing

'Sql="select * from emp_org_mst where org_code = '"&owner_org&"'"
'Set rs_owner=DbConn.Execute(Sql)

'owner_orgname = rs_owner("org_name")
'rs_owner.close()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
				/*if(document.frm.emp_name.value == ""){
					alert('성명을 입력하세요');
					frm.emp_name.focus();
					return false;
				}*/

				var result = confirm('정말 등록하시겠습니까?');

				if(result == true){
					return true;
				}
				return false;
			}

			function file_browse(){
           		document.frm.att_file.click();
           		document.frm.text1.value=document.frm.att_file.value;
			}

			$(document).ready(function(){
				// select box 값이 변경될때 선택된 현재값
				$("#grade").change(function() {
					// alert($(this).val()); // 값
					// alert($(this).children("option:selected").text()); // 내부text

					var params = { "user_id" : '<%=emp_no%>'
								 , "grade" : $(this).val()
								 };
					$.ajax({
						 url: "/insa/insa_emp_master_view_ajax.asp"
						,async: false
						,type: 'post'
						,data: params
						,dataType: "json"
						,contentType: "application/x-www-form-urlencoded; charset=euc-kr"
						,beforeSend: function(jqXHR){
							jqXHR.overrideMimeType("application/x-www-form-urlencoded; charset=euc-kr");
						}
						,error: function(jqXHR, status, errorThrown){
							alert("에러가 발생하였습니다.\n상태코드 : " + jqXHR.responseText + " : " + status + " : " + errorThrown);
						}
						,success: function(data) {
							var result = data.result;

    						if ( result=="succ")
    						{
								alert("권한레벨이 정상적으로 변경되었습니다.");
							}
						}
					});
				});
			});
		</script>

	</head>
	<body>
    <%
    '<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="inview()">
	%>
		<div id="wrap">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/isna/insa_emp_master_view.asp" method="post" name="frm" enctype="multipart/form-data">
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
							<tr>
								<td colspan="2" rowspan="4" class="left">
									<img src="<%=photo_image%>" width="110" height="120" alt=""/>
                                </td>
								<th>사원&nbsp;&nbsp;번호</th>
                                <td class="left"><%=emp_no%>
                                    <input type="hidden" name="emp_no" value="<%=emp_no%>">&nbsp;</td>
                                <th>성명(한글)</th>
                                <td class="left"><%=emp_name%>
                                    <input type="hidden" name="emp_name" id="emp_name" size="13" value="<%=emp_name%>">&nbsp;
								</td>
								<th>성명(영문)</th>
								<td colspan="2" class="left">
									<%=emp_ename%>&nbsp;
								</td>
                                <th>생년월일</th>
                                <td colspan="2" class="left">
									<%=emp_birthday%>&nbsp;―&nbsp;
									<input type="radio" name="emp_birthday_id" value="양" <%If emp_birthday_id = "양" Then %>checked<%End If %> disabled/>양
              						<input type="radio" name="emp_birthday_id" value="음" <%If emp_birthday_id = "음" Then %>checked<%End If %> disabled/>음
                                </td>
                            </tr>
                                <th>소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
								<td colspan="3" class="left"><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;&nbsp;<%=emp_reside_company%></td>
                                <th>조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
		                        <td colspan="5" class="left">
								<%
								Call EmpOrgCodeSelect(emp_org_code)

								If f_toString(emp_reside_company, "") <> "" Then
									Response.Write "(상주처회사&nbsp;:&nbsp;"&emp_reside_company&")&nbsp;"
								End If
								%>
								</td>
                            </tr>
                                <th>직원구분</th>
                                <td class="left"><%=emp_type%>&nbsp;</td>
                               	<th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;급</th>
								<td class="left"><%=emp_grade%>&nbsp;</td>
                                <th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;위</th>
								<td class="left"><%=emp_job%>&nbsp;</td>
                                <th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;책</th>
                                <td class="left"><%=emp_position%>&nbsp;</td>
                                <th>직&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;무</th>
								<td class="left"><%=emp_jikmu%>&nbsp;</td>
                           </tr>
                           <tr>
                                <th>최초입사일</th>
                                <td class="left"><%=emp_first_date%>&nbsp;</td>
                                <th>입&nbsp;&nbsp;&nbsp;사&nbsp;&nbsp;&nbsp;일</th>
                                <td class="left"><%=emp_in_date%>&nbsp;</td>
                                <th>퇴직기산일</th>
                                <td class="left"><%=emp_end_gisan%>&nbsp;</td>
                                <th>근속기산일</th>
                                <td class="left"><%=emp_gunsok_date%>&nbsp;</td>
                                <th>연차기산일</th>
                                <td class="left"><%=emp_yuncha_date%>&nbsp;</td>
                            </tr>
                            <tr>
                                <th colspan="2">주민번호</th>
								<td colspan="2" class="left"><%=emp_person1%>―<%=emp_person2%>&nbsp;(<%=emp_sex%>)</td>
                                <th>전화번호</th>
								<td colspan="3" class="left"><%=emp_tel_ddd%>―<%=emp_tel_no1%>―<%=emp_tel_no2%>&nbsp;</td>
                                <th>핸드폰</th>
								<td colspan="3" class="left"><%=emp_hp_ddd%>―<%=emp_hp_no1%>―<%=emp_hp_no2%>&nbsp;</td>
                            </tr>
                            <tr>
								<th colspan="2">주소(현)</th>
								<td colspan="7" class="left">(<%=emp_zipcode%>)<%=emp_sido%>&nbsp;<%=emp_gugun%>&nbsp;<%=emp_dong%>&nbsp;<%=emp_addr%>&nbsp;</td>
                                </td>
                                <th>e-메일주소</th>
								<td colspan="2" class="left"><%=emp_email%>@k-one.co.kr&nbsp;</td>
                            </tr>
                         	<tr>
								<th colspan="2" class="first">경조가입여부</th>
                                <td colspan="3" class="left"><%=emp_sawo_date%>&nbsp;
								<input type="radio" name="emp_sawo_id" value="Y" <%If emp_sawo_id = "Y" Then %>checked<%End If %> disabled/>가입
              					<input name="emp_sawo_id" type="radio" value="N" <%If emp_sawo_id = "N" Then %>checked<%End If %> disabled/>안함
                                </td>
								<th>결혼기념일</th>
                                <td class="left"><%=emp_marry_date%>&nbsp;</td>
                               	<th>취미</th>
                                <td class="left"><%=emp_hobby%>&nbsp;</td>
                                <th>장애/등급</th>
								<td colspan="2" class="left"><%=emp_disabled%>(<%=emp_disab_grade%>)&nbsp;</td>
                 			</tr>
                            <tr>
                                <th colspan="2" >병역유형</th>
                                <td class="left"><%=emp_military_id%>&nbsp;<%=emp_military_grade%></td>
                                </td>
                                <th>병역 복무기간</th>
                                <td colspan="3" class="left"><%=emp_military_date1%>∼<%=emp_military_date2%>&nbsp;</td>
                                <th>면제사유</th>
								<td class="left"><%=emp_military_comm%>&nbsp;</td>
                                <th>종교</th>
                                <td colspan="2" class="left"><%=emp_faith%>&nbsp;</td>
							</tr>
                            <tr>
                        		<th colspan="2" class="first">실근무지/주소</th>
                              <%
							  	Dim stay_name, stay_sido, stay_gugun, stay_dong, stay_addr

								If emp_stay_code <> "" Then
								   objBuilder.Append "SELECT stay_name, stay_sido, stay_gugun, stay_dong, stay_addr "
								   objBuilder.Append "FROM emp_stay WHERE stay_code = '"&emp_stay_code&"'"

								   Set rsStay = DBConn.Execute(objBuilder.ToString())
								   objBuilder.Clear()

							       If Not rsStay.eof Then
								       stay_name = rsStay("stay_name")
								       stay_sido = rsStay("stay_sido")
								       stay_gugun = rsStay("stay_gugun")
								       stay_dong = rsStay("stay_dong")
								       stay_addr = rsStay("stay_addr")
								   End If
								    rsStay.Close() : Set rsStay = Nothing
								End If
								DBConn.Close() : Set DBConn = Nothing
							  %>
                                <td colspan="2" class="left"><%=emp_stay_code%>&nbsp;<%=stay_name%></td>
                                <td colspan="5" class="left"><%=stay_sido%>&nbsp;<%=stay_gugun%>&nbsp;<%=stay_dong%>&nbsp;<%=stay_addr%>&nbsp;</td>
                                <th>한진그룹여부</th>
                                <td colspan="2" class="left">
								<input type="radio" name="mg_group" value="1" <%If mg_group = "1" Then %>checked<%End If %> disabled/>일반그룹
              					<input name="mg_group" type="radio" value="2" <%If mg_group = "2" Then %>checked<%End If %> disabled/>한진그룹
                                </td>
                            </tr>
                            <tr>
                        		<th colspan="2" class="first">내선번호</th>
                                <td class="left"><%=emp_extension_no%>&nbsp;</td>
                                <th>최종학력</th>
                                <td class="left"><%=emp_last_edu%>&nbsp;</td>
                                <th>비용 그룹</th>
                                <td class="left"><%=cost_group%>&nbsp;</td>
                                <th>비용구분</th>
                                <td class="left"><%=cost_center%>&nbsp;</td>
								<th>급여대상</th>
								<td colspan="2" class="left"><%=emp_pay_id%>&nbsp;</td>
                            </tr>
                            <tr>
                        		<th colspan="2" class="first">입력자</th>
                                <td colspan="2" class="left"><%=emp_reg_date%>&nbsp;(<%=emp_reg_user%>)</td>
                                <th>수정자</th>
                                <td colspan="2" class="left"><%=emp_mod_date%>&nbsp;(<%=emp_mod_user%>)</td>
								<th>급여ID</th>
								<td class="left"><%=dz_id%>&nbsp;</td>
								<th>권한레벨</th>
								<td colspan="2" class="left">
									<%
									'if user_id = "101100" or user_id = "101063" Or user_id = "102592" then
									If SysAdminYn = "Y" Then
									%>
										<select name="grade" id="grade" style="width:50px">
											<option value=""  <%If grade = ""  Then %>selected<%End if %>></option>
											<option value="0" <%If grade = "0" Then %>selected<%End if %>>0</option>
											<option value="1" <%If grade = "1" Then %>selected<%End if %>>1</option>
											<option value="2" <%If grade = "2" Then %>selected<%End if %>>2</option>
											<option value="3" <%If grade = "3" Then %>selected<%End if %>>3</option>
											<option value="4" <%If grade = "4" Then %>selected<%End if %>>4</option>
											<option value="5" <%If grade = "5" Then %>selected<%End if %>>5</option>
											<option value="6" <%If grade = "6" Then %>selected<%End if %>>6</option>
										</select>
									<%
									Else
										Response.Write grade
									End If
									%>
								</td>
                            </tr>
						</tbody>
					</table>
				</div>
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
                        <div align=left>
                             <strong class="btnType01">
								<input type="button" value="닫기" onclick="javascript:goAction();">
							 </strong>
                             <a href="#" class="btnType04" onClick="pop_Window('/insa/insa_card_print.asp?emp_no=<%=emp_no%>','인사 카드 출력','scrollbars=yes,width=750,height=600')">인사카드 출력</a>
                        </div>
				    </td>
                    <td width="80%">
					    <div class="btnCenter">
                             <a href="#" onClick="pop_Window('/insa/insa_appoint_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','발령사항','scrollbars=yes,width=1200,height=600')" class="btnType04">☞발령사항</a>
                             <a href="#" onClick="pop_Window('/insa/insa_family_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','가족사항','scrollbars=yes,width=800,height=400')" class="btnType04">☞가족사항</a>
                             <a href="#" onClick="pop_Window('/insa/insa_school_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','학력사항','scrollbars=yes,width=800,height=400')" class="btnType04">☞학력사항</a>
                             <a href="#" onClick="pop_Window('/insa/insa_career_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','경력사항','scrollbars=yes,width=850,height=400')" class="btnType04">☞경력사항</a>
                             <a href="#" onClick="pop_Window('/insa/insa_qual_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','자격증사항','scrollbars=yes,width=800,height=400')" class="btnType04">☞자격증사항</a>
                             <a href="#" onClick="pop_Window('/insa/insa_edu_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','교육사항','scrollbars=yes,width=800,height=400')" class="btnType04">☞교육사항</a>
                             <a href="#" onClick="pop_Window('/insa/insa_language_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','어학능력','scrollbars=yes,width=800,height=400')" class="btnType04">☞어학능력</a>
                             <a href="#" onClick="pop_Window('/insa/insa_reward_punish_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','상벌사항','scrollbars=yes,width=900,height=400')" class="btnType04">☞상벌사항</a>
                             <a href="#" onClick="pop_Window('/insa/insa_comment_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','특이사항','scrollbars=yes,width=800,height=400')" class="btnType04">☞특이사항</a>
					    </div>
                    </td>
			      </tr>
				</table>
				<input type="hidden" name="u_type" value="<%=u_type%>"/>
				<input type="hidden" name="view_condi" value="<%=view_condi%>"/>
			</form>
		</div>
	</div>
	</body>
</html>