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
Dim be_pg, rsIndi, title_line

be_pg = "/person/insa_open_emp_save.asp"
title_line = "인사기본사항 변경"

objBuilder.Append "CALL USP_PERSON_INDIVIDUAL_INFO('"&emp_no&"');"

'Call Rs_Open(rsIndi, DBConn, objBuilder.ToString())
Set rsIndi = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsIndi.EOF Then
	Dim arrIndi, emp_name, emp_ename, emp_type, emp_sex, emp_person1
	Dim emp_person2, emp_image, emp_first_date, emp_in_date, emp_gunsok_date
	Dim emp_yuncha_date, emp_end_gisan, emp_end_date, emp_bonbu, emp_saupbu
	Dim emp_team, emp_org_code, emp_org_name, emp_org_baldate
	Dim emp_stay_code, emp_reside_place, emp_reside_company, emp_grade, emp_grade_date
	Dim emp_job, emp_position, emp_jikgun, emp_jikmu, emp_birthday
	Dim emp_birthday_id, emp_family_zip, emp_family_sido, emp_family_gugun, emp_family_dong
	Dim emp_family_addr, emp_zipcode, emp_sido, emp_gugun, emp_dong
	Dim emp_addr, emp_tel_ddd, emp_tel_no1, emp_tel_no2, emp_hp_ddd
	Dim emp_hp_no1, emp_hp_no2, emp_email, emp_military_id, emp_military_date1, emp_military_date2
	Dim emp_military_grade, emp_military_comm, emp_hobby, emp_faith, emp_last_edu, emp_marry_date
	Dim emp_disabled, emp_disab_grade, emp_sawo_id, emp_sawo_date
	Dim emp_emergency_tel, emp_nation_code, emp_extension_no, emp_reg_user, emp_mod_user
	Dim photo_image, att_file

	arrIndi = rsIndi.getRows()

	emp_name = arrIndi(0, 0)
	emp_ename = arrIndi(1, 0)
	emp_type = arrIndi(2, 0)
	emp_sex = arrIndi(3, 0)
	emp_person1 = arrIndi(4, 0)
	emp_person2 = arrIndi(5, 0)
	emp_image = arrIndi(6, 0)
	emp_first_date = arrIndi(7, 0)
	emp_in_date = arrIndi(8, 0)
	emp_gunsok_date = arrIndi(9, 0)
	emp_yuncha_date = arrIndi(10, 0)
	emp_end_gisan = arrIndi(11, 0)
	emp_end_date = arrIndi(12, 0)
	emp_company = arrIndi(13, 0)
	emp_bonbu = arrIndi(14, 0)
	emp_saupbu = arrIndi(15, 0)
	emp_team = arrIndi(16, 0)
	emp_org_code = arrIndi(17, 0)
	emp_org_name = arrIndi(18, 0)
	emp_org_baldate = arrIndi(19, 0)
	emp_stay_code = arrIndi(20, 0)
	emp_reside_place = arrIndi(21, 0)
	emp_reside_company = arrIndi(22, 0)
	emp_grade = arrIndi(23, 0)
	emp_grade_date = arrIndi(24, 0)
	emp_job = arrIndi(25, 0)
	emp_position = arrIndi(26, 0)
	emp_jikgun = arrIndi(27, 0)
	emp_jikmu = arrIndi(28, 0)
	emp_birthday = arrIndi(29, 0)
	emp_birthday_id = arrIndi(30, 0)
	emp_family_zip = arrIndi(31, 0)
	emp_family_sido = arrIndi(32, 0)
	emp_family_gugun = arrIndi(33, 0)
	emp_family_dong = arrIndi(34, 0)
	emp_family_addr = arrIndi(35, 0)
	emp_zipcode = arrIndi(36, 0)
	emp_sido = arrIndi(37, 0)
	emp_gugun = arrIndi(38, 0)
	emp_dong = arrIndi(39, 0)
	emp_addr = arrIndi(40, 0)
	emp_tel_ddd = arrIndi(41, 0)
	emp_tel_no1 = arrIndi(42, 0)
	emp_tel_no2 = arrIndi(43, 0)
	emp_hp_ddd = arrIndi(44, 0)
	emp_hp_no1 = arrIndi(45, 0)
	emp_hp_no2 = arrIndi(46, 0)
	emp_email = arrIndi(47, 0)
	emp_military_id = arrIndi(48, 0)
	emp_military_date1 = arrIndi(49, 0)
	emp_military_date2 = arrIndi(50, 0)
	emp_military_grade = arrIndi(51, 0)
	emp_military_comm = arrIndi(52, 0)
	emp_hobby = arrIndi(53, 0)
	emp_faith = arrIndi(54, 0)
	emp_last_edu = arrIndi(55, 0)
	emp_marry_date = arrIndi(56, 0)
	emp_disabled = arrIndi(57, 0)
	emp_disab_grade = arrIndi(58, 0)
	emp_sawo_id = arrIndi(59, 0)
	emp_sawo_date = arrIndi(60, 0)
	emp_emergency_tel = arrIndi(61, 0)
	emp_nation_code = arrIndi(62, 0)
	emp_extension_no = arrIndi(63, 0)
	emp_reg_user = arrIndi(64, 0)
	emp_mod_user = arrIndi(65, 0)

	photo_image = "/emp_photo/" & emp_image
	att_file = photo_image
Else
	Response.Write "<script type='text/javascript'>"
	Response.write "	alert('등록된 사번이 아닙니다.\n\n다시 확인해주세요.');"
	Response.write "	location.replace('/person/insa_person_mg.asp');"
	Response.write "</script>"
	Response.End
End If

Call Rs_Close(rsIndi)
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
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}

			//생년월일
			$(function(){
				$("#datepicker5").datepicker();
				$("#datepicker5").datepicker("option", "dateFormat", "yy-mm-dd");
				$("#datepicker5").datepicker("setDate", "<%=emp_birthday%>");
			});

			//결혼기념일
			$(function(){
				$("#datepicker7").datepicker();
				$("#datepicker7").datepicker("option", "dateFormat", "yy-mm-dd");
				$("#datepicker7").datepicker("setDate", "<%=emp_marry_date%>");
			});

			//병영 복무 시작일
			$(function(){
				$("#datepicker8").datepicker();
				$("#datepicker8").datepicker("option", "dateFormat", "yy-mm-dd");
				$("#datepicker8").datepicker("setDate", "<%=emp_military_date1%>");
			});

			//병영 복무 종료일
			$(function(){
				$("#datepicker9").datepicker();
				$("#datepicker9").datepicker("option", "dateFormat", "yy-mm-dd");
				$("#datepicker9").datepicker("setDate", "<%=emp_military_date2%>");
			});

			function goBefore(){
			   history.go(-1);
			}
			//submit validate
			function chkfrm(){
				/*if(isEmpty($('#emp_ename').val())){
					alert('영문성명을 입력해주세요.');
					frm.emp_ename.focus();
					return false;
				}*/

				if(isEmpty($('#emp_hp_ddd').val())){
					alert('휴대폰번호를 입력해주세요.');
					return false;
				}

				if(isEmpty($('#emp_hp_no1').val())){
					alert('휴대폰번호를 입력해주세요.');
					return false;
				}

				if(isEmpty($('#emp_hp_no2').val())){
					alert('휴대폰번호를 입력하세요.');
					return false;
				}

				if(isEmpty($('#emp_sido').val())){
					alert('주소(현)를 조회해주세요.');
					return false;
				}

				/*if(isEmpty($('#emp_addr').val())){
					alert('현주소 번지를 입력해주세요.');
					frm.emp_addr.focus();
					return false;
				}

				if(isEmpty($('#emp_email').val())){
					alert('이메일을 입력해주세요.');
					frm.emp_email.focus();
					return false;
				}*/

				if($('#'))

				if(isEmpty($('#emp_emergency_tel').val())){
					alert('비상연락 전화번호를 입력해주세요.');
					frm.emp_emergency_tel.focus();
					return false;
				}

//				if(document.frm.emp_extension_no.value =="") {
//					alert('내선번호를 입력하세요');
//					frm.emp_extension_no.focus();
//					return false;}

				if(isEmpty($('#emp_last_edu').val())){
					alert('최종학력을 입력해주세요.');
					frm.emp_last_edu.focus();
					return false;
				}

				/*
  				if(isEmpty($('#att_file').val())){
					alert('사진을 등록 하세요');
					frm.att_file.focus();
					return false;
				}*/

				if(!confirm('등록 하시겠습니까?')) return false;
				else return true;
			}

			//form 전송
			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					//document.frm.submit();

					var form = $('frm')[0];
					var formData = new FormData(form);

					console.log(formData);
				}
			}

			/*function file_browse()	{
           		document.frm.att_file.click();
           		document.frm.text1.value=document.frm.att_file.value;
			}*/

			//opener관련 오류가 발생하는 경우 아래 주석을 해지하고, 사용자의 도메인정보를 입력합니다. ("팝업API 호출 소스"도 동일하게 적용시켜야 합니다.)
			//document.domain = "abc.go.kr";
			function jusoCallBack(roadFullAddr,roadAddrPart1,addrDetail,roadAddrPart2,engAddr,jibunAddr,zipNo,admCd,rnMgtSn,bdMgtSn,detBdNmList,bdNm,bdKdcd,siNm,sggNm,emdNm,liNm,rn,udrtYn,buldMnnm,buldSlno,mtYn,lnbrMnnm,lnbrSlno,emdNo,gubun){
				if(gubun === 'juso'){
					$('#emp_sido').val(siNm);
					$('#emp_gugun').val(sggNm);
					$('#emp_dong').val(rn + ' ' + buldMnnm);
					$('#emp_addr').val(roadAddrPart2 + ' ' + addrDetail);
					$('#emp_zipcode').val(zipNo);
				}else if(gubun === 'family'){
					$('#emp_family_sido').val(siNm);
					$('#emp_family_gugun').val(sggNm);
					$('#emp_family_dong').val(rn + ' ' + buldMnnm);
					$('#emp_family_addr').val(roadAddrPart2 + ' ' + addrDetail);
					$('#emp_family_zip').val(zipNo);
				}
			}
		</script>
		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
    <!--<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="inview()">-->
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
	<!--<body>-->
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>

				<form action="<%=be_pg%>" method="post" name="frm" enctype="multipart/form-data">
					<input type="hidden" name="emp_no" value="<%=emp_no%>"/>
					<input type="hidden" name="emp_name" value="<%=emp_name%>"/>
					<input type="hidden" name="emp_ename" id="emp_ename" value="<%=emp_ename%>"/>
					<input type="hidden" name="emp_email" id="emp_email" value="<%=emp_email%>"/>
					<input type="hidden" name="mg_group" value="<%=mg_group%>"/>
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
                                <td class="left"><%=emp_no%>&nbsp;</td>
                                <th>성명(한글)</th>
                                <td class="left"><%=emp_name%>&nbsp;</td>
								<th>성명(영문)</th>
								<td colspan="2" class="left"><%=emp_ename%></td>
                                <th>생년월일<span style="color:red;">*</span></th>
                                <td colspan="2" class="left">
									<input type="text" name="emp_birthday" size="10" id="datepicker5" style="width:70px;" value="<%=emp_birthday%>" readonly/>
									&nbsp;―&nbsp;
									<input type="radio" name="emp_birthday_id" value="양" <%If emp_birthday_id = "양" Then %>checked<%End If %> />양
              						<input type="radio" name="emp_birthday_id" value="음" <%If emp_birthday_id = "음" Then %>checked<%End If %> />음
                                </td>
                            </tr>
							<tr>
                                <th>소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
								<td colspan="3" class="left">(<%=emp_org_code%>)<%=emp_org_name%></td>
                                <th>조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                                <td colspan="5" class="left">
								<%
								Call EmpOrgCodeSelect(emp_org_code)

								If f_toString(emp_reside_company, "") <> "" Then
									Response.Write "(" & emp_reside_company & ")"
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
								<td colspan="2" class="left"><%=emp_person1%>-<%=emp_person2%>&nbsp;(<%=emp_sex%>)&nbsp;</td>
                                <th>전화번호</th>
								<td colspan="3" class="left">
									<input type="text" name="emp_tel_ddd" id="emp_tel_ddd" size="3" maxlength="3" value="<%=emp_tel_ddd%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									  -
									<input type="text" name="emp_tel_no1" id="emp_tel_no1" size="4" maxlength="4" value="<%=emp_tel_no1%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									  -
									<input type="text" name="emp_tel_no2" id="emp_tel_no2" size="4" maxlength="4" value="<%=emp_tel_no2%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
                                </td>
                                <th>휴대폰번호<span style="color:red;">*</span></th>
								<td colspan="3" class="left">
									<input type="text" name="emp_hp_ddd" id="emp_hp_ddd" size="3" maxlength="3" value="<%=emp_hp_ddd%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									  -
									<input type="text" name="emp_hp_no1" id="emp_hp_no1" size="4" maxlength="4" value="<%=emp_hp_no1%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									  -
									<input type="text" name="emp_hp_no2" id="emp_hp_no2" size="4" maxlength="4" value="<%=emp_hp_no2%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
                                </td>
                            </tr>
                            <tr>
								<th colspan="2">주소(현)<span style="color:red;">*</span></th>
								<td colspan="7" class="left">
									<input type="text" name="emp_zipcode" id="emp_zipcode" style="width:50px;" value="<%=emp_zipcode%>" readonly/>
									-
									<input type="text" name="emp_sido" id="emp_sido" style="width:80px;" value="<%=emp_sido%>" readonly/>
									<input type="text" name="emp_gugun" id="emp_gugun" style="width:80px;" value="<%=emp_gugun%>" readonly/>
									<input type="text" name="emp_dong" id="emp_dong" style="width:150px;" value="<%=emp_dong%>" readonly/>
									<input type="text" name="emp_addr" id="emp_addr" style="width:230px;" value="<%=emp_addr%>" notnull errname="번지" onKeyUp="checklength(this,50)" readonly/>
									<a href="#" class="btnType03" onClick="pop_Window('/insa/jusoPopup.asp?gubun=juso','family_zip_select','scrollbars=yes,width=600,height=400')">주소조회</a>
                                </td>
                                <th>이메일</th>
								<td colspan="2" class="left"><%=emp_email%>@k-one.co.kr</td>
                            </tr>
                         	<tr>
                                <th colspan="2" class="first">경조가입여부</th>
                                <td class="left">
								<%
								If emp_sawo_id = "Y" Then
									Response.Write "가입"
								Else
									Response.Write "안함"
								End If
								%>&nbsp;
								</td>
                                </td>
								<th>경조가입일</th>
                                <td class="left"><%=emp_sawo_date%>&nbsp;</td>
								<th>결혼기념일</th>
                                <td class="left">
									<input type="text" name="emp_marry_date" size="10" id="datepicker7" style="width:70px;" value="<%=emp_marry_date%>" readonly/>
								</td>
								<th>취미</th>
                                <td class="left">
									<input type="text" name="emp_hobby" id="emp_hobby" style="width:80px;" value="<%=emp_hobby%>" />
								</td>
                                <th>장애/등급</th>
								<td colspan="2" class="left">
								<%
								Response.Write emp_disabled
								If f_toString(emp_disab_grade, "") <> "" Then
									Response.Write "-"  & emp_disab_grade
								End If
								%>&nbsp;
								</td>
                 			</tr>
                            <tr>
                                <th colspan="2" >병역유형</th>
                                <td class="left">
								<%
								Call SelectEmpEtcCodeList("emp_military_id", "emp_military_id", "width:90px;", "06", emp_military_id)
								%>
                                </td>
                                <th>병역계급</th>
                                <td class="left">
								<%
								Call SelectEmpEtcCodeList("emp_military_grade", "emp_military_grade", "width:90px;", "07", emp_military_grade)

								DBConn.Close() : Set DBConn = Nothing
								%>
                                </td>
                                <th>병역 복무기간</th>
                                <td colspan="2" class="left">
									<input type="text" name="emp_military_date1" size="10" id="datepicker8" style="width:70px;" value="<%=emp_military_date1%>" readonly/>
									∼
									<input type="text" name="emp_military_date2" size="10" id="datepicker9" style="width:70px;" value="<%=emp_military_date2%>" readonly/>
                                </td>
                                <th>면제사유</th>
								<td class="left">
									<input type="text" name="emp_military_comm" id="emp_military_comm" style="width:80px;" value="<%=emp_military_comm%>"/>
								</td>
                                <th>종교</th>
                                <td class="left">
									<input type="text" name="emp_faith" id="emp_faith" style="width:80px;" value="<%=emp_faith%>"/>
								</td>
							</tr>
                            <tr>
								<th colspan="2">비상연락<span style="color:red;">*</span></th>
								<td class="left">
									<input type="text" name="emp_emergency_tel" id="emp_emergency_tel" style="width:80px;" value="<%=emp_emergency_tel%>" onKeyUp="checklength(this, 15)" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								</td>
								 <th>최종학력<span style="color:red;">*</span></th>
                                <td class="left">
                                <select name="emp_last_edu" id="emp_last_edu" value="<%=emp_last_edu%>" style="width:90px;">
			            	        <option value="" <%If emp_last_edu = "" Then %>selected<%End If %>>선택</option>
				                    <option value='고등학교' <%If emp_last_edu = "고등학교" Then %>selected<% End If %>>고등학교</option>
                                    <option value='전문대' <%If emp_last_edu = "전문대" Then %>selected<% End If %>>전문대</option>
                                    <option value='대학교' <%If emp_last_edu = "대학교" Then %>selected<% End If %>>대학교</option>
                                    <option value='대학원수료' <%If emp_last_edu = "대학원수료" Then %>selected<% End If %>>대학원수료</option>
                                    <option value='대학원' <%If emp_last_edu = "대학원" Then %>selected<% End If %>>대학원</option>
                                </select>
                        		<th class="first">내선번호</th>
                                <td colspan="2" class="left">
									<input name="emp_extension_no" type="text" id="emp_extension_no" size="16 " value="<%=emp_extension_no%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
                                </td>
                                </td>
                                <th>한진그룹여부</th>
                                <td colspan="3" class="left">
								<%
								Select Case mg_group
									Case "1"
										Response.Write "일반그룹"
									Case "2"
										Response.Write "한진그룹"
								End Select
								%>
                                </td>
                             </tr>
						</tbody>
					</table>
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="8%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th scope="row">사진등록</th>
								<td class="left">
									<input type="file" name= "att_file" size="70" accept="image/gif" /> * 첨부파일은 1개만 가능하며 최대용량은 2MB
                                </td>
							</tr>
						</tbody>
                    </table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    <div class="btnCenter">
                         <span class="btnType01"><input type="button" value="수정" onclick="javascript:frmcheck();" /></span>
                         <span class="btnType01"><input type="button" value="이전" onclick="javascript:goBefore();" /></span>
                    </div>
                    </td>
				    <td width="52%">
					<div class="btnCenter">
                    <a class="btnType04">☞ 가족사항 ☞ 학력사항 ☞ 경력사항 ☞ 자격사항 ☞ 교육사항 ☞ 어학능력을 등록하시기 바랍니다</a>
					</div>
                    </td>
			      </tr>
				</table>
			</form>
		</div>
	</div>
	</body>
</html>