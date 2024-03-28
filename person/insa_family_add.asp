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
Dim u_type, family_empno, family_seq, emp_name
Dim family_rel, family_name, family_birthday, family_birthday_id
Dim family_job, family_live, family_person1, family_person2
Dim family_tel_ddd, family_tel_no1, family_tel_no2, family_support_yn
Dim family_national, family_disab, family_merit, family_serius
Dim family_pensioner, family_witak, family_holt, family_holt_date, family_children
Dim curr_date, title_line, rsFamily

u_type = Request.QueryString("u_type")
family_empno = Request.QueryString("family_empno")
family_seq = Request.QueryString("family_seq")
emp_name = Request.QueryString("emp_name")

family_rel = ""
family_name = ""
family_birthday = ""
family_birthday_id = "음"
family_job = ""
family_live = "안함"
family_person1 = ""
family_person2 = ""
family_tel_ddd = ""
family_tel_no1 = ""
family_tel_no2 = ""
family_support_yn = "N"
family_national = "내국인"
family_disab = ""
family_merit = ""
family_serius = ""
family_pensioner = ""
family_witak = ""
family_holt = ""
family_holt_date = ""
family_children = ""

curr_date = Mid(CStr(Now()), 1, 10)
title_line = "가족사항 등록"

If u_type = "U" Then
	objBuilder.Append "SELECT family_rel, family_name, family_birthday, family_birthday_id, family_job, "
	objBuilder.Append "	family_live, family_person1, family_person2, family_tel_ddd, family_tel_no1, "
	objBuilder.Append "	family_tel_no2, family_support_yn, family_national, family_disab, family_merit, "
	objBuilder.Append "	family_serius, family_pensioner, family_witak, family_holt, family_holt_date, "
	objBuilder.Append "	family_children "
	objBuilder.Append "FROM emp_family "
	objBuilder.Append "WHERE family_empno = '"&family_empno&"' AND family_seq = '"&family_seq&"';"

	Set rsFamily = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	family_rel = rsFamily("family_rel")
    family_name = rsFamily("family_name")
    family_birthday = rsFamily("family_birthday")
    family_birthday_id = rsFamily("family_birthday_id")
    family_job = rsFamily("family_job")
    family_live = rsFamily("family_live")
    family_person1 = rsFamily("family_person1")
    family_person2 = rsFamily("family_person2")
	family_tel_ddd = rsFamily("family_tel_ddd")
    family_tel_no1 = rsFamily("family_tel_no1")
    family_tel_no2 = rsFamily("family_tel_no2")
	family_support_yn = rsFamily("family_support_yn")
	family_national = rsFamily("family_national")
    family_disab = rsFamily("family_disab")
	family_merit = rsFamily("family_merit")
    family_serius = rsFamily("family_serius")
    family_pensioner = rsFamily("family_pensioner")
    family_witak = rsFamily("family_witak")
    family_holt = rsFamily("family_holt")
    family_holt_date = rsFamily("family_holt_date")
	family_children = rsFamily("family_children")

	If family_birthday = "1900-01-01"  Then
	   family_birthday = ""
	end If

	If family_holt_date = "1900-01-01"  Then
	   family_holt_date = ""
	End If

	rsFamily.Close() : Set rsFamily = Nothing

	title_line = "가족사항 변경"
End If
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
			//생년월일
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=family_birthday%>" );
			});

			//입양일자
			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=family_holt_date%>" );
			});

			function goAction(){
			   window.close();
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.family_rel == ""){
					alert('관계를 선택해주세요.');
					frm.family_rel.focus();
					return false;
				}

				if(document.frm.family_name.value == ""){
					alert('성명을 입력해주세요.');
					frm.family_name.focus();
					return false;
				}

				if(document.frm.family_birthday.value == ""){
					alert('생년월일을 입력해주세요.');
					frm.family_birthday.focus();
					return false;
				}

				if(document.frm.family_tel_ddd.value == ""){
					alert('휴대폰번호를 입력해주세요.');
					frm.family_tel_ddd.focus();
					return false;
				}

				if(document.frm.family_tel_no1.value == ""){
					alert('휴대폰번호를 입력해주세요.');
					frm.family_tel_no1.focus();
					return false;
				}

				if(document.frm.family_tel_no2.value ==""){
					alert('휴대폰번호를 입력해주세요.');
					frm.family_tel_no2.focus();
					return false;
				}

				/*if(document.frm.family_support_yn.value == ""){
					alert('부양가족 여부를 선택해주세요.');
					frm.family_support_yn.focus();
					return false;
				}*/

				var result = confirm('등록 하시겠습니까?');

				if(result){
					return true;
				}else{
					return false
				};
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
	<body>
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/person/insa_family_add_save.asp" method="post" name="frm">
			<div class="gView">
			  <table cellpadding="0" cellspacing="0" class="tableWrite">
				<colgroup>
					<col width="15%" >
					<col width="18%" >
					<col width="15%" >
					<col width="18%" >
					<col width="15%" >
					<col width="*" >
				</colgroup>
				<tbody>
				<tr>
					<th style="background:#FFFFE6">사번</th>
					<td class="left" bgcolor="#FFFFE6">
						<input type="text" name="family_empno" id="family_empno" size="14" value="<%=family_empno%>" readonly class="no-input"/>
						<input type="hidden" name="family_seq" value="<%=family_seq%>"/>
					</td>
					<th style="background:#FFFFE6">성명</th>
					<td colspan="3" class="left" bgcolor="#FFFFE6">
						<input type="text" name="emp_name" id="emp_name" size="14" value="<%=emp_name%>" readonly class="no-input"/>
					</td>
				</tr>
				<tr>
				  <th>관계<span style="color:red;">*</span></th>
				  <td colspan="5" class="left">
					  <select name="family_rel" id="family_rel" value="<%=family_rel%>" style="width:100px;">
						  <option value="" <%If family_rel = "" Then %>selected<%End If %>>선택</option>
						  <option value='부' <%If family_rel = "부" Then %>selected<%End If %>>부</option>
						  <option value='모' <%If family_rel = "모" Then %>selected<%End If %>>모</option>
						  <option value='남편' <%If family_rel = "남편" Then %>selected<%End If %>>남편</option>
						  <option value='아내' <%If family_rel = "아내" Then %>selected<%End If %>>아내</option>
						  <option value='아들' <%If family_rel = "아들" Then %>selected<%End If %>>아들</option>
						  <option value='딸' <%If family_rel = "딸" Then %>selected<%End If %>>딸</option>
						  <option value='조부' <%If family_rel = "조부" Then %>selected<%End If %>>조부</option>
						  <option value='조모' <%If family_rel = "조모" Then %>selected<%End If %>>조모</option>
						  <option value='외조부' <%If family_rel = "외조부" Then %>selected<%End If %>>외조부</option>
						  <option value='외조모' <%If family_rel = "외조모" Then %>selected<%End If %>>외조모</option>
						  <option value='시부' <%If family_rel = "시부" Then %>selected<%End If %>>시부</option>
						  <option value='시모' <%If family_rel = "시모" Then %>selected<%End If %>>시모</option>
						  <option value='장인' <%If family_rel = "장인" Then %>selected<%End If %>>장인</option>
						  <option value='장모' <%If family_rel = "장모" Then %>selected<%End If %>>장모</option>
						  <option value='형(형제자매)' <%If family_rel = "형(형제자매)" Then %>selected<%End If %>>형(형제자매)</option>
						  <option value='매(형제자매)' <%If family_rel = "매(형제자매)" Then %>selected<%End If %>>매(형제자매)</option>
						  <option value='제(형제자매)' <%If family_rel = "제(형제자매)" Then %>selected<%End If %>>제(형제자매)</option>
						  <option value='올케' <%If family_rel = "올케" Then %>selected<%End If %>>올케</option>
						  <option value='자(형제자매)' <%If family_rel = "자(형제자매)" Then %>selected<%End If %>>자(형제자매)</option>
						  <option value='손자' <%If family_rel = "손자" Then %>selected<%End If %>>손자</option>
						  <option value='손녀' <%If family_rel = "손녀" Then %>selected<%End If %>>손녀</option>
						  <option value='자부' <%If family_rel = "자부" Then %>selected<%End If %>>자부</option>
						  <option value='손부' <%If family_rel = "손부" Then %>selected<%End If %>>손부</option>
						  <option value='기타관계' <%If family_rel = "기타관계" Then %>selected<%End If %>>기타관계</option>
					  </select>
					  &nbsp;
					  (<span style="color:red;font-size:11px;">위탁아동인경우는 기타관계를 선택하십시요.</span>)
				  </td>
				</tr>
				<tr>
				  <th>성명<span style="color:red;">*</span></th>
				  <td colspan="2" class="left">
					<input type="text" name="family_name" id="family_name" size="14" value="<%=family_name%>"/></td>
				  <th>출생지</th>
				  <td colspan="2" class="left">
					  <select name="family_national" id="family_national" value="<%=family_national%>" style="width:90px">
						  <option value="" <%If family_rel = "" Then %>selected<%End If %>>선택</option>
						  <option value='내국인' <%If family_national = "내국인" Then %>selected<%End If %>>내국인</option>
						  <option value='외국인' <%If family_national = "외국인" Then %>selected<%End If %>>외국인</option>
					  </select>
				  </td>
				</tr>
				<tr>
					<th>생년월일<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
						<input type="text" name="family_birthday" value="<%=family_birthday%>" style="width:70px;text-align:center" id="datepicker" readonly="true"/>
						&nbsp;&nbsp;
						<input type="radio" name="family_birthday_id" value="양" <%If family_birthday_id = "양" Then %>checked<%End If %>/>양
						<input type="radio" name="family_birthday_id" value="음" <%If family_birthday_id = "음" Then %>checked<%End If %>/>음
					</td>
					<th>주민등록번호</th>
					<td colspan="2" class="left">
						<input type="text" name="family_person1" id="family_person1" style="width:40px;" maxlength="6" value="<%=family_person1%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
						-
						<input type="text" name="family_person2" id="family_person2" style="width:50px;"  maxlength="7" value="<%=family_person2%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
						&nbsp;
						(<span style="color:red;font-size:11px;">연말정산 시 필수</span>)
					</td>
			   </tr>
			   <tr>
					<th>직업</th>
					<td colspan="2" class="left">
						<input name="family_job" type="text" id="family_job" style="width:160px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=family_job%>"/>
					</td>
					<th>휴대폰번호<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
						<input type="text" name="family_tel_ddd" id="family_tel_ddd" size="3" maxlength="3" value="<%=family_tel_ddd%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								  -
						<input type="text" name="family_tel_no1" id="family_tel_no1" size="4" maxlength="4" value="<%=family_tel_no1%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								  -
						<input type="text" name="family_tel_no2" id="family_tel_no2" size="4" maxlength="4" value="<%=family_tel_no2%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
					</td>
				</tr>
				<tr>
					<th>동거여부</th>
					<td colspan="2" class="left">
						<input type="radio" name="family_live" value="동거" <%If family_live = "동거" Then %>checked<%End If %>/>동거
						<input type="radio" name="family_live" value="안함" <%If family_live = "안함" Then %>checked<%End If %>/>안함
					<th>부양가족</th>
					<td colspan="2" class="left">
						<input type="radio" name="family_support_yn" value="Y" <%If family_support_yn = "Y" Then %>checked<%End If %>/>부양
						<input type="radio" name="family_support_yn" value="N" <%If family_support_yn = "N" Then %>checked<%End If %>/>안함
				  </td>
				</tr>
				<tr>
					<th>장애인</th>
					<td colspan="2" class="left">
						<input type="checkbox" name="disab_check" value="Y" <%If family_disab = "Y" Then %>checked<%End If %> id="disab_check"/>장애인
						<input type="checkbox" name="merit_check" value="Y" <%If family_merit = "Y" Then %>checked<%End If %> id="merit_check"/>국가유공자
						<input type="checkbox" name="serius_check" value="Y" <%If family_serius = "Y" Then %>checked<%End If %> id="serius_check"/>중증환자
					</td>
					<th>국민기초생활수급</th>
					<td colspan="2" class="left">
						<input type="checkbox" name="pensioner_check" value="Y" <%If family_pensioner = "Y" Then %>checked<%End If %> id="pensioner_check"/>예
					</td>
				</tr>
				<tr>
					<th>입양여부</th>
					<td class="left">
						<input type="checkbox" name="holt_check" value="Y" <%If family_holt = "Y" Then %>checked<%End If %> id="holt_check"/>예
					</td>
					<th>입양일자</th>
					<td class="left">
						<input type="text" name="family_holt_date" value="<%=family_holt_date%>" style="width:70px;text-align:center" id="datepicker1" readonly/>
					</td>
					<th>위탁아동</th>
					<td class="left">
						<input type="checkbox" name="witak_check" value="Y" <%If family_witak = "Y" Then %>checked<%End If %> id="witak_check"/>예
					</td>
				</tr>
				<tr>
					<th>자녀양육</th>
					<td colspan="5" class="left">
						<input type="checkbox" name="children_check" value="Y" <%If family_children = "Y" Then %>checked<%End If %> id="children_check"/>예&nbsp;&nbsp;
						(<span style="color:red;font-size:11px;">6세미만 자녀의경우 연말정산 추가공제 체크</span>)
					</td>
				</tr>
				</tbody>
			  </table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="<%If u_type = "U" Then%>수정<%Else%>등록<%End If%>" onclick="javascript:frmcheck();"/></span>
				<span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"/></span>
			</div>
			<input type="hidden" name="u_type" value="<%=u_type%>"/>
			</form>
		</div>
	</body>
</html>