<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)
curr_hh = int(cstr(datepart("h",now)))
curr_mm = int(cstr(datepart("n",now)))

user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
pay_grade = request.cookies("nkpmg_user")("coo_pay_grade")

u_type = request("u_type")
emp_no = request("emp_no")
view_condi=Request("view_condi")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_owner = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect


title_line = "[ 인사기본사항 조회 ]"

	Sql="select * from emp_master where emp_no = '"&emp_no&"'"
	Set rs=DbConn.Execute(Sql)
'Response.write(sql)

	emp_name = rs("emp_name")
    emp_ename = rs("emp_ename")
    emp_type = rs("emp_type")
    emp_sex = rs("emp_sex")
    emp_person1 = rs("emp_person1")
    emp_person2 = rs("emp_person2")
	if emp_person2 <> "" then
	   sex_id = mid(cstr(emp_person2),1,1)
	   if sex_id = "1" then
	         emp_sex = "남"
		  else
		     emp_sex = "여"
	   end if
	end if
    emp_image = rs("emp_image")
	att_file = rs("emp_image")
    emp_first_date = rs("emp_first_date")
    emp_in_date = rs("emp_in_date")
    emp_gunsok_date = rs("emp_gunsok_date")
    emp_yuncha_date = rs("emp_yuncha_date")
    emp_end_gisan = rs("emp_end_gisan")
    emp_end_date = rs("emp_end_date")
    emp_company = rs("emp_company")
    emp_bonbu = rs("emp_bonbu")
    emp_saupbu = rs("emp_saupbu")
    emp_team = rs("emp_team")
    emp_org_code = rs("emp_org_code")
    emp_org_name = rs("emp_org_name")
    emp_org_baldate = rs("emp_org_baldate")
    emp_stay_code = rs("emp_stay_code")
	emp_stay_name = rs("emp_stay_name")
    emp_reside_place = rs("emp_reside_place")
	emp_reside_company = rs("emp_reside_company")
    emp_grade = rs("emp_grade")
    emp_grade_date = rs("emp_grade_date")
    emp_job = rs("emp_job")
    emp_position = rs("emp_position")
    emp_jikgun = rs("emp_jikgun")
    emp_jikmu = rs("emp_jikmu")
    emp_birthday = rs("emp_birthday")
    emp_birthday_id = rs("emp_birthday_id")
    emp_family_zip = rs("emp_family_zip")
    emp_family_sido = rs("emp_family_sido")
    emp_family_gugun = rs("emp_family_gugun")
    emp_family_dong = rs("emp_family_dong")
    emp_family_addr = rs("emp_family_addr")
    emp_zipcode = rs("emp_zipcode")
    emp_sido = rs("emp_sido")
    emp_gugun = rs("emp_gugun")
    emp_dong = rs("emp_dong")
    emp_addr = rs("emp_addr")
    emp_tel_ddd = rs("emp_tel_ddd")
    emp_tel_no1 = rs("emp_tel_no1")
    emp_tel_no2 = rs("emp_tel_no2")
    emp_hp_ddd = rs("emp_hp_ddd")
    emp_hp_no1 = rs("emp_hp_no1")
    emp_hp_no2 = rs("emp_hp_no2")
    emp_email = rs("emp_email")
    emp_military_id = rs("emp_military_id")
    emp_military_date1 = rs("emp_military_date1")
    emp_military_date2 = rs("emp_military_date2")
    emp_military_grade = rs("emp_military_grade")
    emp_military_comm = rs("emp_military_comm")
    emp_hobby = rs("emp_hobby")
    emp_faith = rs("emp_faith")
    emp_last_edu = rs("emp_last_edu")
    emp_marry_date = rs("emp_marry_date")
    emp_disabled = rs("emp_disabled")
    emp_disab_grade = rs("emp_disab_grade")
    emp_sawo_id = rs("emp_sawo_id")
	if rs("emp_sawo_id") = "" or isNull(emp_sawo_id) then
	   emp_sawo_id = "N"
	end if
    emp_sawo_date = rs("emp_sawo_date")
    emp_emergency_tel = rs("emp_emergency_tel")
    emp_nation_code = rs("emp_nation_code")
	emp_extension_no = rs("emp_extension_no")
	cost_group = rs("cost_group")
	cost_center = rs("cost_center")
	emp_pay_id = rs("emp_pay_id")
'   end_date = mid(cstr(now()),1,10)
    emp_reg_date = rs("emp_reg_date")
    emp_reg_user = rs("emp_reg_user")
	emp_mod_date = rs("emp_mod_date")
    emp_mod_user = rs("emp_mod_user")
	photo_image = "/emp_photo/" + rs("emp_image")
    att_file = rs("emp_image")

	if emp_pay_id = "5" then
	       emp_pay_id = "안함"
	   else
	       emp_pay_id = "지급"
	end if

	if rs("emp_military_date1") = "1900-01-01" then
           emp_military_date1 = ""
           emp_military_date2 = ""
    end if
	if rs("emp_birthday") = "1900-01-01" then
           emp_birthday = ""
    end if
    if rs("emp_marry_date") = "1900-01-01" then
           emp_marry_date = ""
    end if
	if rs("emp_grade_date") = "1900-01-01" then
           emp_grade_date = ""
    end if
	if rs("emp_end_date") = "1900-01-01" then
           emp_end_date = ""
    end if
	if rs("emp_org_baldate") = "1900-01-01" then
           emp_org_baldate = ""
    end if
	if rs("emp_sawo_date") = "1900-01-01" then
           emp_sawo_date = ""
    end if

	rs.close()

	sql="select * from memb where user_id='"&emp_no&"'"
	set rs_memb=dbconn.execute(sql)
	if not rs_memb.eof then
		   mg_group = rs_memb("mg_group")
		   grade    = rs_memb("grade")
	   else
		   mg_group = "1"
		   grade    = ""
    end if
	rs_memb.close()
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
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=emp_first_date%>" );
			});
			$(function() {    $( "#datepicker1" ).datepicker();
											$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker1" ).datepicker("setDate", "<%=emp_in_date%>" );
			});
			$(function() {    $( "#datepicker2" ).datepicker();
											$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker2" ).datepicker("setDate", "<%=emp_end_gisan%>" );
			});
			$(function() {    $( "#datepicker3" ).datepicker();
											$( "#datepicker3" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker3" ).datepicker("setDate", "<%=emp_gunsok_date%>" );
			});
			$(function() {    $( "#datepicker4" ).datepicker();
											$( "#datepicker4" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker4" ).datepicker("setDate", "<%=emp_yuncha_date%>" );
			});
			$(function() {    $( "#datepicker5" ).datepicker();
											$( "#datepicker5" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker5" ).datepicker("setDate", "<%=emp_birthday%>" );
			});
			$(function() {    $( "#datepicker6" ).datepicker();
											$( "#datepicker6" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker6" ).datepicker("setDate", "<%=emp_sawo_date%>" );
			});
			$(function() {    $( "#datepicker7" ).datepicker();
											$( "#datepicker7" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker7" ).datepicker("setDate", "<%=emp_marry_date%>" );
			});
			$(function() {    $( "#datepicker8" ).datepicker();
											$( "#datepicker8" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker8" ).datepicker("setDate", "<%=emp_military_date1%>" );
			});
			$(function() {    $( "#datepicker9" ).datepicker();
											$( "#datepicker9" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker9" ).datepicker("setDate", "<%=emp_military_date2%>" );
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
				if(document.frm.emp_name.value =="") {
					alert('성명을 입력하세요');
					frm.emp_name.focus();
					return false;}


				a=confirm('등록하시겠습니까?');
				if (a==true) {
					return true;
				}
				return false;
			}
			function file_browse()	{
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
						 url: "insa_emp_master_view_ajax.asp"
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
								alert("권한레벨이 변경되었습니다.")
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
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_emp_master_view.asp" method="post" name="frm" enctype="multipart/form-data">
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
                                <img src="<%=photo_image%>" width=110 height=120 alt="">
                                </td>
								<th>사원&nbsp;&nbsp;번호</th>
                                <td class="left"><%=emp_no%>
                                    <input name="emp_no" type="hidden" value="<%=emp_no%>">&nbsp;</td>
                                <th>성명(한글)</th>
                                <td class="left"><%=emp_name%>
                                    <input name="emp_name" type="hidden" id="emp_name" size="13" value="<%=emp_name%>">&nbsp;</td>
								<th>성명(영문)</th>
								<td colspan="2" class="left"><%=emp_ename%>&nbsp;</td>
                                <th>생년월일</th>
                                <td colspan="2" class="left"><%=emp_birthday%>&nbsp;―&nbsp;
								<input type="radio" name="emp_birthday_id" value="양" <% if emp_birthday_id = "양" then %>checked<% end if %>>양
              					<input name="emp_birthday_id" type="radio" value="음" <% if emp_birthday_id = "음" then %>checked<% end if %>>음
                                </td>
                            </tr>
                                <th>소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
								<td colspan="3" class="left"><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;&nbsp;<%=emp_reside_company%></td>
                                <th>조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                            <% if emp_reside_company = "" or isnull(emp_reside_company) then %>
                                <td colspan="5" class="left"><%=emp_company%>―<%=emp_bonbu%>―<%=emp_saupbu%>―<%=emp_team%>&nbsp;</td>
                            <%    else %>
                                <td colspan="5" class="left"><%=emp_company%>―<%=emp_bonbu%>―<%=emp_saupbu%>―<%=emp_team%>&nbsp;&nbsp;(상주처회사&nbsp;:&nbsp;<%=emp_reside_company%>)&nbsp;</td>
                            <%  end if %>
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
                                <th colspan="2" >본적(주소)</th>
								<td colspan="7" class="left">(<%=emp_family_zip%>)<%=emp_family_sido%>&nbsp;<%=emp_family_gugun%>&nbsp;<%=emp_family_dong%>&nbsp;<%=emp_family_addr%>&nbsp;</td>
                                <th>비상연락</th>
								<td colspan="2" class="left"><%=emp_emergency_tel%>&nbsp;</td>
                            </tr>
                            <tr>
								<th colspan="2">주소(현)</th>
								<td colspan="7" class="left">(<%=emp_zipcode%>)<%=emp_sido%>&nbsp;<%=emp_gugun%>&nbsp;<%=emp_dong%>&nbsp;<%=emp_addr%>&nbsp;</td>
                                </td>
                                <th>e-메일주소</th>
								<td colspan="2" class="left"><%=emp_email%>@k-won.co.kr&nbsp;</td>
                            </tr>
                         	<tr>
								<th colspan="2" class="first">경조가입여부</th>
                                <td colspan="3" class="left"><%=emp_sawo_date%>&nbsp;
								<input type="radio" name="emp_sawo_id" value="Y" <% if emp_sawo_id = "Y" then %>checked<% end if %>>가입
              					<input name="emp_sawo_id" type="radio" value="N" <% if emp_sawo_id = "N" then %>checked<% end if %>>안함
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
								if emp_stay_code <> "" then
								   Sql="select * from emp_stay where stay_code = '"&emp_stay_code&"'"
								   Rs_stay.Open Sql, Dbconn, 1
							       if not rs_stay.eof then
								       emp_stay_name = rs_stay("stay_name")
								       stay_sido = rs_stay("stay_sido")
								       stay_gugun = rs_stay("stay_gugun")
								       stay_dong = rs_stay("stay_dong")
								       stay_addr = rs_stay("stay_addr")
								    end if
								    rs_stay.Close()
								end if
							  %>
                                <td colspan="2" class="left"><%=emp_stay_code%>&nbsp;<%=emp_stay_name%></td>
                                <td colspan="5" class="left"><%=stay_sido%>&nbsp;<%=stay_gugun%>&nbsp;<%=stay_dong%>&nbsp;<%=stay_addr%>&nbsp;</td>
                                <th>한진그룹여부</th>
                                <td colspan="2" class="left">
								<input type="radio" name="mg_group" value="1" <% if mg_group = "1" then %>checked<% end if %>>일반그룹
              					<input name="mg_group" type="radio" value="2" <% if mg_group = "2" then %>checked<% end if %>>한진그룹
                                </td>
                            </tr>
                            <tr>
                        		<th colspan="2" class="first">내선번호</th>
                                <td colspan="2" class="left"><%=emp_extension_no%>&nbsp;</td>
                                <th>최종학력</th>
                                <td colspan="2" class="left"><%=emp_last_edu%>&nbsp;</td>
                                <th>Cost Group</th>
                                <td colspan="2" class="left"><%=cost_group%>&nbsp;</td>
                                <th>비용구분</th>
                                <td class="left"><%=cost_center%>&nbsp;</td>
                            </tr>
                            <tr>
                        		<th colspan="2" class="first">입력자</th>
                                <td colspan="2" class="left"><%=emp_reg_date%>&nbsp;(<%=emp_reg_user%>)</td>
                                <th>수정자</th>
                                <td colspan="3" class="left"><%=emp_mod_date%>&nbsp;(<%=emp_mod_user%>)</td>
                                <th>급여대상</th>
								<td class="left"><%=emp_pay_id%>&nbsp;</td>
								<th>권한레벨</th>
								<td class="left">
									<%
									' 차재명, 윤성희
									if user_id = "101100" or user_id = "101063" then
										%>
										<select name="grade" id="grade" style="width:50px">
											<option value=""  <% If grade = ""  then %>selected<% end if %>></option>
											<option value="0" <% If grade = "0" then %>selected<% end if %>>0</option>
											<option value="1" <% If grade = "1" then %>selected<% end if %>>1</option>
											<option value="2" <% If grade = "2" then %>selected<% end if %>>2</option>
											<option value="3" <% If grade = "3" then %>selected<% end if %>>3</option>
											<option value="4" <% If grade = "4" then %>selected<% end if %>>4</option>
											<option value="5" <% If grade = "5" then %>selected<% end if %>>5</option>
											<option value="6" <% If grade = "6" then %>selected<% end if %>>6</option>
										</select>
										<%
									else
										%><%=grade%>&nbsp;<%
									end if
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
                             <strong class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></strong>
                             <a href="#" class="btnType04" onClick="pop_Window('insa_card_print.asp?emp_no=<%=emp_no%>','emp_card_pop','scrollbars=yes,width=750,height=600')">인사카드 출력</a>
                        </div>
				    </td>
                    <td width="80%">
					    <div class="btnCenter">
                             <a href="#" onClick="pop_Window('insa_appoint_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','appointview','scrollbars=yes,width=1200,height=600')" class="btnType04">☞ 발령사항</a>
                             <a href="#" onClick="pop_Window('insa_family_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','familyview','scrollbars=yes,width=800,height=400')" class="btnType04">☞ 가족사항</a>
                             <a href="#" onClick="pop_Window('insa_school_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','schoolview','scrollbars=yes,width=800,height=400')" class="btnType04">☞ 학력사항</a>
                             <a href="#" onClick="pop_Window('insa_career_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','careerview','scrollbars=yes,width=850,height=400')" class="btnType04">☞ 경력사항</a>
                             <a href="#" onClick="pop_Window('insa_qual_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','qualview','scrollbars=yes,width=800,height=400')" class="btnType04">☞ 자격사항</a>
                             <a href="#" onClick="pop_Window('insa_edu_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','eduview','scrollbars=yes,width=800,height=400')" class="btnType04">☞ 교육사항</a>
                             <a href="#" onClick="pop_Window('insa_language_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','eduview','scrollbars=yes,width=800,height=400')" class="btnType04">☞ 어학능력</a>
                             <a href="#" onClick="pop_Window('insa_reward_punish_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','reward_punishview','scrollbars=yes,width=900,height=400')" class="btnType04">☞ 상벌사항</a>
                    <% if user_id <> "100001" then %>
                             <a href="#" onClick="pop_Window('insa_comment_view.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>','eduview','scrollbars=yes,width=800,height=400')" class="btnType04">☞ 특이사항</a>
                    <% end if %>
					    </div>
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="view_condi" value="<%=view_condi%>" ID="Hidden1">
				</form>
		</div>
	</div>
	</body>
</html>