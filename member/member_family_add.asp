<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
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
Dim f_seq, f_birthday_id, f_live, f_support_yn, f_national
Dim curr_date, title_line

f_birthday_id = "음"
f_live = "안함"
f_support_yn = "N"
f_national = "내국인"

curr_date = Mid(CStr(Now()), 1, 10)
title_line = "가족사항 등록"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>회원관리</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "" );
			});

			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "" );
			});

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
				if(document.frm.f_rel == ""){
					alert('관계를 선택하세요');
					frm.f_rel.focus();
					return false;
				}

				if(document.frm.f_name.value == ""){
					alert('성명을 입력하세요');
					frm.f_name.focus();
					return false;
				}

				if(document.frm.f_birthday.value == ""){
					alert('생년월일을 입력하세요');
					frm.f_birthday.focus();
					return false;
				}

				if(document.frm.f_tel_ddd.value == ""){
					alert('전화번호를 입력하세요');
					frm.family_tel_no1.focus();
					return false;
				}

				if(document.frm.f_tel_no1.value == ""){
					alert('전화번호를 입력하세요');
					frm.family_tel_no1.focus();
					return false;
				}

				if(document.frm.f_tel_no2.value ==""){
					alert('전화번호를 입력하세요');
					frm.family_tel_no2.focus();
					return false;
				}

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
			<form action="/member/member_family_proc.asp" method="post" name="frm">
			<div class="gView">
			  <table cellpadding="0" cellspacing="0" class="tableWrite">
				<colgroup>
					<col width="15%" >
					<col width="17%" >
					<col width="15%" >
					<col width="18%" >
					<col width="15%" >
					<col width="*" >
				</colgroup>
				<tbody>
				<tr>
					<th style="background:#FFFFE6">성명</th>
					<td colspan="5" class="left" bgcolor="#FFFFE6">
						<input type="text" name="m_name" id="m_name" size="14" value="<%=m_name%>" class="no-input" readonly="true"/>
					</td>
				</tr>
				<tr>
				  <th>관계<span style="color:red;">*</span></th>
				  <td colspan="5" class="left">
					  <select name="f_rel" id="f_rel" style="width:100px">
						  <option value="">선택</option>
						  <option value='부'>부</option>
						  <option value='모'>모</option>
						  <option value='남편'>남편</option>
						  <option value='아내'>아내</option>
						  <option value='아들'>아들</option>
						  <option value='딸'>딸</option>
						  <option value='조부'>조부</option>
						  <option value='조모'>조모</option>
						  <option value='외조부'>외조부</option>
						  <option value='외조모'>외조모</option>
						  <option value='시부'>시부</option>
						  <option value='시모'>시모</option>
						  <option value='장인'>장인</option>
						  <option value='장모'>장모</option>
						  <option value='형(형제자매)'>형(형제자매)</option>
						  <option value='매(형제자매)'>매(형제자매)</option>
						  <option value='제(형제자매)'>제(형제자매)</option>
						  <option value='올케'>올케</option>
						  <option value='자(형제자매)'>자(형제자매)</option>
						  <option value='손자'>손자</option>
						  <option value='손녀'>손녀</option>
						  <option value='자부'>자부</option>
						  <option value='손부'>손부</option>
						  <option value='기타관계'>기타관계</option>
					  </select>
					  &nbsp;위탁아동인경우는 기타관계를 선택하세요.
				  </td>
				</tr>
				<tr>
				  <th>성명<span style="color:red;">*</span></th>
				  <td colspan="2" class="left">
					<input type="text" name="f_name" id="f_name" size="14"/></td>
				  <th>출생지</th>
				  <td colspan="2" class="left">
					  <select name="f_national" id="f_national" style="width:90px">
						  <option value="">선택</option>
						  <option value='내국인'>내국인</option>
						  <option value='외국인'>외국인</option>
					  </select>
				  </td>
				</tr>
				<tr>
					<th>생년월일<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
						<input type="text" name="f_birthday" id="datepicker" style="width:70px;text-align:center" readonly="true"/>
						&nbsp;&nbsp;
						<input type="radio" name="f_birthday_id" id="f_birthday_id" value="양" checked/>양
						<input type="radio" name="f_birthday_id" id="f_birthday_id" value="음"/>음
					</td>
					<th>주민등록번호</th>
					<td colspan="2" class="left">
						<input type="text" name="f_person1" id="f_person1" size="6" maxlength="6" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
						-
						<input type="text" name="f_person2" id="f_person2" size="7" maxlength="7" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
						(연말정산필수)
					</td>
			   </tr>
			   <tr>
					<th>직업</th>
					<td colspan="2" class="left">
						<input type="text" name="f_job" id="f_job" style="width:160px; ime-mode:active" onKeyUp="checklength(this,20);"/>
					</td>
					<th>전화번호<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
						<input type="text" name="f_tel_ddd" id="f_tel_ddd" size="3" maxlength="3" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								  -
						<input type="text" name="f_tel_no1" id="f_tel_no1" size="4" maxlength="4" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								  -
						<input type="text" name="f_tel_no2" id="f_tel_no2" size="4" maxlength="4" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
					</td>
				</tr>
				<tr>
					<th>동거여부</th>
					<td colspan="2" class="left">
						<input type="radio" name="f_live" id="f_live" value="동거" checked/>동거
						<input type="radio" name="f_live" id="f_live" value="안함"/>안함
					<th>부양가족</th>
					<td colspan="2" class="left">
						<input type="radio" name="f_support_yn" id="f_support_yn" value="Y" checked/>부양
						<input type="radio" name="f_support_yn" id="f_support_yn" value="N"/>안함
				  </td>
				</tr>
				<tr>
					<th>장애인</th>
					<td colspan="2" class="left">
						<input type="checkbox" name="disab_check" id="disab_check" value="Y"/>장애인
						<input type="checkbox" name="merit_check" id="merit_check" value="Y"/>국가유공자
						<input type="checkbox" name="serius_check" id="serius_check" value="Y"/>중증환자
					</td>
					<th>국민기초생활수급</th>
					<td colspan="2" class="left">
						<input type="checkbox" name="pensioner_check" id="pensioner_check" value="Y"/>예
					</td>
				</tr>
				<tr>
					<th>입양여부</th>
					<td class="left">
						<input type="checkbox" name="holt_check" id="holt_check" value="Y"/>예
					</td>
					<th>입양일자</th>
					<td class="left">
						<input name="f_holt_date" type="text" id="datepicker1" style="width:70px;text-align:center" readonly="true"/>
					</td>
					<th>위탁아동</th>
					<td class="left">
						<input type="checkbox" name="witak_check" id="witak_check" value="Y"/>예
					</td>
				</tr>
				<tr>
					<th>자녀양육</th>
					<td colspan="5" class="left">
						<input type="checkbox" name="children_check" id="children_check" value="Y"/>예&nbsp;&nbsp;(6세미만 자녀의경우 연말정산 추가공제 체크)
					</td>
				</tr>
				</tbody>
			  </table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();"/></span>
				<span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"/></span>
			</div>
			</form>
		</div>
	</body>
</html>