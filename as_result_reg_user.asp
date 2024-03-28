<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)

acpt_no = request("acpt_no")
be_pg = request("be_pg")

page = request("page")
from_date = request("from_date")
to_date = request("to_date")
date_sw = request("date_sw")
process_sw = request("process_sw")
field_check = request("field_check")
field_view = request("field_view")
view_sort = request("view_sort")
page_cnt = request("page_cnt")
condi_com = request("company")
view_c = request("view_c")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_into = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect


'if rs("as_process") = "입고" then
	sql = "select max(in_seq) as max_seq from as_into where acpt_no = "&int(acpt_no)
	set rs_into = dbconn.execute(sql)
	max_seq = rs_into("max_seq")

	if isnull(max_seq) then	
		in_process = "없음"
		in_place = "없음"
	  else
		in_seq = max_seq
		sql = "select in_process,in_place from as_into where acpt_no = "&int(acpt_no)&" and in_seq = "&int(in_seq)
		Set rs_into = DbConn.Execute(SQL)
		if rs_into.eof then
			in_process = "없음"
			in_place = "없음"
		  else
			in_process = rs_into("in_process")
			in_place = rs_into("in_place")
		end if
	end if
'end if

Sql = "select * from as_acpt where acpt_no = "&int(acpt_no)
Set rs = DbConn.Execute(SQL)

acpt_date = mid(cstr(rs("acpt_date")),1,10)
acpt_hh = int(datepart("h",rs("acpt_date")))
acpt_mm = int(datepart("n",rs("acpt_date")))
acpt_ss = datepart("s",rs("acpt_date"))

if acpt_hh < 10 then
	acpt_hh = "0" + cstr(acpt_hh)
end if

if acpt_mm < 10 then
	acpt_mm = "0" + cstr(acpt_mm)
end if

if arrival_time = "0000" or arrival_time = null or arrival_time = "" then
	arrival_date = curr_date
	arrival_time = "0000"
end if

if isnull(rs("dev_inst_cnt")) or rs("dev_inst_cnt") = "" then
	dev_inst_cnt = "1"
  else
  	dev_inst_cnt = rs("dev_inst_cnt")
end if

as_type = rs("as_type")
if rs("sms") = "Y" then
	sms_view = "발송"
  else
  	sms_view = "발송안함"
end if
new_sms = "N"

title_line = "A/S 결과 등록"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
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
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=rs("request_date")%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
											$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker1" ).datepicker("setDate", "<%=rs("visit_date")%>" );
			});	  
			$(function() {    $( "#datepicker2" ).datepicker();
											$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker2" ).datepicker("setDate", "<%=rs("in_date")%>" );
			});	  
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.as_process.value == "입고" && document.frm.as_process_old.value == "입고") {
					alert('입고 상태에서는 수정이 불가 합니다 !!!');
					frm.as_process.focus();
					return false;}
				if(document.frm.c_grade.value >"4") {
					alert('수정 또는 등록 권한이 없습니다 !!!');
					frm.addr.focus();
					return false;}
				if(document.frm.acpt_user.value == "") {
					alert('사용자를 입력하세요 !!!');
					frm.acpt_user.focus();
					return false;}
				if(document.frm.addr.value =="") {
					alert('나머지 주소를 입력하세요');
					frm.addr.focus();
					return false;}
				if(document.frm.as_memo.value =="") {
					alert('장애내용을 입력하세요');
					frm.as_memo.focus();
					return false;}			
				if(document.frm.request_date.value =="") {
					alert('요청일을 입력하세요');
					frm.request_date.focus();
					return false;}
				if(document.frm.request_date.value < document.frm.acpt_date.value) {
					alert('요청일이 접수일보다 빠름니다');
					frm.request_date.focus();
					return false;}
				if(document.frm.request_hh.value >"23"||document.frm.request_hh.value <"00") {
					alert('요청시간이 잘못되었습니다');
					frm.request_hh.focus();
					return false;}
				if(document.frm.request_mm.value >"59"||document.frm.request_mm.value <"00") {
					alert('요청분이 잘못되었습니다');
					frm.request_mm.focus();
					return false;}
				if(document.frm.request_date.value == document.frm.acpt_date.value) {
					if(document.frm.request_hh.value < document.frm.acpt_hh.value) {
						alert('요청시간이 접수시간 보다 빠름니다');
						frm.request_hh.focus();
						return false;}}
				if(document.frm.request_date.value == document.frm.acpt_date.value) {
					if(document.frm.request_hh.value == document.frm.acpt_hh.value) {
						if(document.frm.request_mm.value <= document.frm.acpt_mm.value) {
							alert('요청분이 접수분 보다 빠름니다');
							frm.request_mm.focus();
							return false;}}}
			
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
					if(document.frm.visit_date.value =="") {
						alert('완료일을 입력하세요');
						frm.visit_date.focus();
						return false;}
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
					if(document.frm.visit_date.value < document.frm.acpt_date.value) {
						alert('완료일이 접수일보다 빠름니다');
						frm.visit_date.focus();
						return false;}
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
					if(document.frm.visit_date.value > document.frm.curr_date.value) {
						alert('완료일이 현재일보다 빠름니다');
						frm.visit_date.focus();
						return false;}
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
					if(document.frm.visit_hh.value >"23"||document.frm.visit_hh.value <"00") {
						alert('완료시간이 잘못되었습니다');
						frm.visit_hh.focus();
						return false;}
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
					if(document.frm.visit_mm.value >"59"||document.frm.visit_mm.value <"00") {
						alert('완료분이 잘못되었습니다');
						frm.visit_mm.focus();
						return false;}
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
					if(document.frm.visit_date.value == document.frm.acpt_date.value) {
						if(document.frm.visit_hh.value < document.frm.acpt_hh.value) {
							alert('완료시간이 접수시간 보다 빠름니다');
							frm.visit_hh.focus();
							return false;}}
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
					if(document.frm.visit_date.value == document.frm.acpt_date.value) {
						if(document.frm.visit_hh.value == document.frm.acpt_hh.value) {
							if(document.frm.visit_mm.value <= document.frm.acpt_mm.value) {
								alert('완료분이 접수분 보다 빠름니다');
								frm.visit_mm.focus();
								return false;}}}
			
				if(document.frm.as_process_old.value =="입고" || document.frm.as_process_old.value =="대체입고") 
					if(document.frm.as_process.value =="접수" || document.frm.as_process.value =="연기") {
						document.frm.as_process.value = document.frm.as_process_old.value
						alert('입고를 접수나 연기로 변경 불가');
						frm.as_process.focus();
						return false;}
				if(document.frm.as_process_old.value =="입고" || document.frm.as_process_old.value =="대체입고") 
					if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="취소") {
						if(document.frm.in_process.value !="수리완료") {
							if(document.frm.in_process.value !="입고취소") {
								alert('수리완료 또는 입고취소 되지 않아 완료 또는 취소 등록 할수 없습니다');
								frm.as_process.focus();
								return false;}}}
				if(document.frm.as_process.value =="입고" || document.frm.as_process.value =="대체입고" || document.frm.as_process.value =="대체") 
					if(document.frm.as_type.value !="방문처리") {
						alert('입고,대체 및 대체입고는 반드시 방문처리이어야 함');
						frm.as_type.focus();
					return false;}
				if(document.frm.as_process.value =="입고" || document.frm.as_process.value =="대체입고") 
					if(document.frm.into_reason.value =="") {
						alert('입고 및 연기 사유를 입력하세요');
						frm.into_reason.focus();
					return false;}
				if(document.frm.as_process.value =="입고" || document.frm.as_process.value =="대체입고") 
					if(document.frm.in_date.value < document.frm.acpt_date.value) {
						alert('입고일자가 접수일자보다 작습니다');
						frm.in_date.focus();
					return false;}
				if(document.frm.as_process.value =="입고" || document.frm.as_process.value =="대체입고") 
					if(document.frm.in_date.value =="") {
						alert('입고일자를 입력하세요');
						frm.in_date.focus();
					return false;}
				if(document.frm.as_process.value =="입고") 
					if(document.frm.in_date.value > document.frm.curr_date.value) {
						alert('입고일이 현재일보다 빠름니다');
						frm.in_date.focus();
						return false;}
				if(document.frm.as_process.value =="입고" || document.frm.as_process.value =="대체입고") 
					if(document.frm.in_place.value =="없음") {
						alert('입고처를 입력하세요');
						frm.in_place.focus();
					return false;}
				if(document.frm.as_process.value =="입고" || document.frm.as_process.value =="대체입고") 
					if(document.frm.in_replace.value =="") {
						alert('대체여부를 선택하여야 합니다');
						frm.in_replace.focus();
					return false;}
			
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
					if(document.frm.as_history.value =="") {
						alert('처리 내역을 입력하세요');
						frm.as_history.focus();
					return false;}
				if(document.frm.as_process.value =="완료") 
					if(document.frm.as_type.value =="신규설치" || document.frm.as_type.value =="신규설치공사" || document.frm.as_type.value =="이전설치" || document.frm.as_type.value =="이전설치공사" || document.frm.as_type.value =="랜공사" || document.frm.as_type.value =="이전랜공사" || document.frm.as_type.value =="장비회수" || document.frm.as_type.value =="예방점검") {
						if(document.frm.dev_inst_cnt.value < 0 || document.frm.dev_inst_cnt.value > 999 || document.frm.dev_inst_cnt.value == "") {
							alert('설치대수가 999보다 크거나 잘못되었습니다');
							frm.dev_inst_cnt.focus();
					return false;}}
				if(document.frm.as_process.value =="완료") 
					if(document.frm.as_type.value =="신규설치" || document.frm.as_type.value =="신규설치공사" || document.frm.as_type.value =="이전설치" || document.frm.as_type.value =="이전설치공사" || document.frm.as_type.value =="랜공사" || document.frm.as_type.value =="이전랜공사" || document.frm.as_type.value =="장비회수" || document.frm.as_type.value =="예방점검") {
						if(document.frm.ran_cnt.value < 0 || document.frm.ran_cnt.value > 999 || document.frm.ran_cnt.value == "") {
							alert('공사대수가 999보다 크거나 잘못되었습니다');
							frm.ran_cnt.focus();
					return false;}}
				if(document.frm.as_process.value =="완료") 
					if(document.frm.as_type.value =="신규설치" || document.frm.as_type.value =="신규설치공사" || document.frm.as_type.value =="이전설치" || document.frm.as_type.value =="이전설치공사" || document.frm.as_type.value =="랜공사" || document.frm.as_type.value =="이전랜공사") {
						if(document.frm.work_man_cnt.value < 1 || document.frm.work_man_cnt.value > 30 || document.frm.work_man_cnt.value == "") {
							alert('작업 인원수 30보다 크거나 잘못되었습니다');
							frm.work_man_cnt.focus();
					return false;}}
				if(document.frm.as_process.value =="완료") 
					if(document.frm.as_type.value =="신규설치" || document.frm.as_type.value =="신규설치공사" || document.frm.as_type.value =="이전설치" || document.frm.as_type.value =="이전설치공사" || document.frm.as_type.value =="랜공사" || document.frm.as_type.value =="이전랜공사") {
						if(document.frm.alba_cnt.value < 0 || document.frm.alba_cnt.value > 30 || document.frm.alba_cnt.value == "") {
							alert('알바 인원수 30보다 크거나 잘못되었습니다');
							frm.alba_cnt.focus();
					return false;}}
			
				j=0;
				   for(i=0;i<document.frm.err01.length;i++){  
					if (document.frm.err01[i].checked==true){   
					 j++;
					}
				   }
				k=0;
				   for(i=0;i<document.frm.err02.length;i++){  
					if (document.frm.err02[i].checked==true){   
					 k++;
					}
				   }
			
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
				 if(document.frm.as_type.value =="원격처리" || document.frm.as_type.value =="방문처리" || document.frm.as_type.value =="기타") 
					if(document.frm.as_device.value =="데스크탑" || document.frm.as_device.value =="노트북" || document.frm.as_device.value =="DTO" || document.frm.as_device.value =="DTS") 
						if(j == 0 && k == 0) {
							alert('장애처리를 CHECK 하세요');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err03.length;i++){  
					if (document.frm.err03[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
				 if(document.frm.as_type.value =="원격처리" || document.frm.as_type.value =="방문처리" || document.frm.as_type.value =="기타") 
					if(document.frm.as_device.value =="모니터") 
						if(j == 0) {
							alert('모니터 장애처리를 CHECK 하세요');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err04.length;i++){  
					if (document.frm.err04[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
				 if(document.frm.as_type.value =="원격처리" || document.frm.as_type.value =="방문처리" || document.frm.as_type.value =="기타") 
					if(document.frm.as_device.value =="프린터" || document.frm.as_device.value =="스케너" || document.frm.as_device.value =="플로터") 
						if(j == 0) {
							alert('장애처리를 CHECK 하세요');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err05.length;i++){  
					if (document.frm.err05[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
				 if(document.frm.as_type.value =="원격처리" || document.frm.as_type.value =="방문처리" || document.frm.as_type.value =="기타") 
					if(document.frm.as_device.value =="통신장비" || document.frm.as_device.value =="AP" || document.frm.as_device.value =="허브" || document.frm.as_device.value =="라우터" || document.frm.as_device.value =="TA" || document.frm.as_device.value =="네트웍장비" || document.frm.as_device.value =="회선") 
						if(j == 0) {
							alert('통신장비 장애처리를 CHECK 하세요');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err06.length;i++){  
					if (document.frm.err06[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
				 if(document.frm.as_type.value =="원격처리" || document.frm.as_type.value =="방문처리" || document.frm.as_type.value =="기타") 
					if(document.frm.as_device.value =="서버" || document.frm.as_device.value =="워크스테이션") 
						if(j == 0) {
							alert('장애처리를 CHECK 하세요');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err07.length;i++){  
					if (document.frm.err07[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
				 if(document.frm.as_type.value =="원격처리" || document.frm.as_type.value =="방문처리" || document.frm.as_type.value =="기타") 
					if(document.frm.as_device.value =="아답터") 
						if(j == 0) {
							alert('장애처리를 CHECK 하세요');
							frm.as_history.focus();
						return false;}
				j=0;
				   for(i=0;i<document.frm.err09.length;i++){  
					if (document.frm.err09[i].checked==true){   
					 j++;
					}
				   }
				if(document.frm.as_process.value =="완료" || document.frm.as_process.value =="대체" || document.frm.as_process.value =="취소"  || document.frm.as_process.value =="대체입고") 
				 if(document.frm.as_type.value =="원격처리" || document.frm.as_type.value =="방문처리" || document.frm.as_type.value =="기타") 
					if(document.frm.as_device.value =="기타") 
						if(j == 0) {
							alert('기타 내역을 CHECK 하세요');
							frm.as_history.focus();
						return false;}
					
				if(document.frm.as_process.value =="완료"){
					if (document.frm.as_type.value == '신규설치' || document.frm.as_type.value == '신규설치공사' || document.frm.as_type.value == '이전설치' || document.frm.as_type.value == '이전설치공사' || document.frm.as_type.value == '랜공사' || document.frm.as_type.value == '이전랜공사' || document.frm.as_type.value == '장비회수' || document.frm.as_type.value == '예방점검') {
						if(document.frm.att_file1.value =="" && document.frm.att_file2.value =="" && document.frm.att_file3.value =="" && document.frm.att_file4.value =="" && document.frm.att_file5.value =="") {
							alert('사진 첨부가 되지 않았습니다');
							frm.att_file1.focus();
							return false;}}}
					
				if(document.frm.as_process.value =="완료"){
					if (document.frm.as_type.value == '신규설치' || document.frm.as_type.value == '신규설치공사' || document.frm.as_type.value == '이전설치' || document.frm.as_type.value == '이전설치공사' || document.frm.as_type.value == '랜공사' || document.frm.as_type.value == '이전랜공사' || document.frm.as_type.value == '장비회수' || document.frm.as_type.value == '예방점검') {
					{
					b=confirm('작업인원이 ' + document.frm.work_man_cnt.value +'명 맞습니까?')
					if (b==false) {
						return false;
						}
					}
				}}
				{
				a=confirm('변경하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function ce_mod_view() 
			{
				if (document.frm.ce_mod_ck.checked == true) {
					document.getElementById('s_ce').style.display = ''; 
					document.getElementById('s_ce_id').style.display = ''; 
					document.getElementById('ce_mod').style.display = ''; }
				if (document.frm.ce_mod_ck.checked == false) {
					document.getElementById('ce_mod').style.display = 'none'; 
					document.getElementById('s_ce').style.display = 'none'; 
					document.getElementById('s_ce_id').style.display = 'none'; }
			}
			function inview() {
			var c = document.frm.as_process.options[document.frm.as_process.selectedIndex].value;
				if (c == '입고' || c == '대체입고') 
				{
					document.getElementById('in_menu').style.display = '';
				}
			}
			function menu1() {

				var c = document.frm.as_process.options[document.frm.as_process.selectedIndex].value;
				var d = document.frm.as_device.options[document.frm.as_device.selectedIndex].value;
				var e = document.frm.as_type.options[document.frm.as_type.selectedIndex].value;
				var f = document.frm.company.value;
					 {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = 'none';
						document.getElementById('end_keyin2').style.display = 'none';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					}
					if (c == '입고') 
					{
						document.getElementById('in_menu').style.display = '';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = 'none';
						document.getElementById('end_keyin2').style.display = 'none';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					}
					if (c == '완료' || c == '대체' || c == '취소') 
					  if (e == '원격처리' || e == '방문처리' || e == '기타') {
						if (d == '데스크탑' || d == '노트북' || d == 'DTO' || d == 'DTS') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = '';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '완료') 
						if (e == '신규설치' || e == '신규설치공사' || e == '이전설치' || e == '이전설치공사' || e == '랜공사' || e == '이전랜공사' || e == '장비회수' || e == '예방점검') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = '';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = '';		
					}
					if (c == '취소') 
						if (e == '신규설치' || e == '신규설치공사' || e == '이전설치' || e == '이전설치공사' || e == '랜공사' || e == '이전랜공사' || e == '장비회수' || e == '예방점검') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					}
					if (c == '완료' || c == '대체' || c == '취소') 
					  if (e == '원격처리' || e == '방문처리' || e == '기타') {
						if (d == '모니터') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = '';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '완료' || c == '대체' || c == '취소') 
					  if (e == '원격처리' || e == '방문처리' || e == '기타') {
						if (d == '프린터' || d == '스케너' || d == '플로터') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = '';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '완료' || c == '대체' || c == '취소') 
					  if (e == '원격처리' || e == '방문처리' || e == '기타') {
						if (d == '통신장비' || d == 'AP' || d == '허브' || d == '라우터' || d == 'TA' || d == '네트웍장비' || d == '회선') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = '';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '완료' || c == '대체' || c == '취소') 
					  if (e == '원격처리' || e == '방문처리' || e == '기타') {
						if (d == '서버' || d == '워크스테이션') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = '';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '완료' || c == '대체' || c == '취소') 
					  if (e == '원격처리' || e == '방문처리' || e == '기타') {
						if (d == '아답터') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = '';
						document.getElementById('end_menu7').style.display = 'none';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}
					if (c == '완료' || c == '대체' || c == '취소')
					  if (e == '원격처리' || e == '방문처리' || e == '기타') {
						if (d == '기타') {
						document.getElementById('in_menu').style.display = 'none';
						document.getElementById('inst_menu').style.display = 'none';		
						document.getElementById('end_keyin1').style.display = '';
						document.getElementById('end_keyin2').style.display = '';
						document.getElementById('end_menu1').style.display = 'none';
						document.getElementById('end_menu2').style.display = 'none';
						document.getElementById('end_menu3').style.display = 'none';
						document.getElementById('end_menu4').style.display = 'none';
						document.getElementById('end_menu5').style.display = 'none';
						document.getElementById('end_menu6').style.display = 'none';
						document.getElementById('end_menu7').style.display = '';		
						document.getElementById('att_menu').style.display = 'none';		
					  }
					}				
				}
		</script>

	</head>
	<body onLoad="inview()">
		<div id="wrap">			
			<!--#include virtual = "/include/user_header.asp" -->
			<!--#include virtual = "/include/as_sub_menu_user.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="as_result_reg_user_ok.asp" method="post" enctype="multipart/form-data" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">접수번호</th>
								<td class="left"><%=rs("acpt_no")%>
                                <input name="acpt_no" type="hidden" id="acpt_no" value="<%=rs("acpt_no")%>">
                				<input name="c_grade" type="hidden" id="c_grade" value="<%=c_grade%>">
                                </td>
								<th>접수일</th>
								<td class="left"><%=rs("acpt_date")%>
								<input name="acpt_date" type="hidden" id="acpt_date" value="<%=acpt_date%>">
                				<input name="acpt_hh" type="hidden" id="acpt_hh" value="<%=acpt_hh%>">
                				<input name="acpt_mm" type="hidden" id="acpt_mm" value="<%=acpt_mm%>">
                                </td>
								<th>접수자</th>
								<td class="left"><%=rs("acpt_man")%>
                				<input name="curr_date" type="hidden" id="curr_date" value="<%=curr_date%>">
            					</td>
								<th>회사</th>
								<td class="left"><%=rs("company")%>
                				<input name="company" type="hidden" id="company" value="<%=rs("company")%>">
                                </td>
							</tr>
							<tr>
								<th class="first">조직명</th>
								<td class="left"><%=rs("dept")%><input name="dept" type="hidden" id="dept" value="<%=rs("dept")%>"></td>
								<th>전화번호1</th>
								<td class="left"><%=rs("tel_ddd")%>-<%=rs("tel_no1")%>-<%=rs("tel_no2")%></td>
								<th>사용자</th>
								<td class="left">
                                <input name="acpt_user" type="text" size="10" onKeyUp="checklength(this,20)" value="<%=rs("acpt_user")%>">
								  &nbsp;<strong>직급</strong>
                                <input name="user_grade" type="text" size="8" onKeyUp="checklength(this,20)" value="<%=rs("user_grade")%>"></td>
								<th>전화번호2</th>
								<td class="left">
								<select name="hp_ddd" id="hp_ddd">
									<option>선택</option>
									<option value="02" <%If rs("hp_ddd") = "02" then %>selected<% end if %>>02</option>
									<option value="010" <%If rs("hp_ddd") = "010" then %>selected<% end if %>>010</option>
				  					<option value="011" <%If rs("hp_ddd") = "011" then %>selected<% end if %>>011</option>
				  					<option value="016" <%If rs("hp_ddd") = "016" then %>selected<% end if %>>016</option>
				  					<option value="017" <%If rs("hp_ddd") = "017" then %>selected<% end if %>>017</option>
				  					<option value="018" <%If rs("hp_ddd") = "018" then %>selected<% end if %>>018</option>
				  					<option value="019" <%If rs("hp_ddd") = "019" then %>selected<% end if %>>019</option>
								</select>-              	
								<input name="hp_no1" type="text" id="hp_no1" size="4" maxlength="4" value="<%=rs("hp_no1")%>">-
                            	<input name="hp_no2" type="text" id="hp_no2" size="4" maxlength="4" value="<%=rs("hp_no2")%>">
                              </td>
							</tr>
							<tr>
								<th class="first">주소</th>
								<td colspan="5" class="left"><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%>
                                <input name="sido" type="hidden" id="sido" value="<%=rs("sido")%>">
                                <input name="gugun" type="hidden" id="gugun" value="<%=rs("gugun")%>">
                                <input name="dong" type="hidden" id="dong2" value="<%=rs("dong")%>">
              					<input name="addr" type="text" id="addr" style="width:250px" onKeyUp="checklength(this,50)" value="<%=rs("addr")%>">
              					<input name="view_ok" type="hidden" id="view_ok" value="">
                                </td>
								<th>종전문자</th>
								<td class="left"><%=sms_view%></td>
							</tr>
							<tr>
								<th class="first">기존CE</th>
								<td class="left"><%=rs("mg_ce")%>&nbsp;(&nbsp;<%=rs("mg_ce_id")%>&nbsp;)
                                <input name="mg_ce_id" type="hidden" id="mg_ce_id2" value="<%=rs("mg_ce_id")%>">
                				<input name="mg_ce" type="hidden" id="mg_ce" value="<%=rs("mg_ce")%>">                           					
                                </td>
								<th>변경CE</th>
								<td class="left" colspan="3"><strong>변경을 원하면 선택하세요</strong>
								<input name="ce_mod_ck" type="checkbox" id="ce_mod_ck" value="1"  onClick="ce_mod_view()">
                				<input name="s_ce" id="s_ce" type="text" value="<%=user_name%>" size="10" readonly="true" style="display:none">
                				<input name="s_ce_id" id="s_ce_id" type="text" value="<%=user_id%>" size="10" readonly="true" style="display:none">
                				<input name="s_reside_place" type="hidden" id="s_reside_place">
                				<input name="s_team" type="hidden" id="s_team">
                                <a href="#" class="btnType03" onClick="pop_Window('ce_select.asp?gubun=<%="수정"%>&mg_group=<%=mg_group%>','ceselect','scrollbars=yes,width=500,height=400')" id="ce_mod" style="display:none">CE변경</a>
                                </td>
								<th>문자재발송</th>
								<td class="left">
                                <input type="radio" name="new_sms" value="Y" <% if new_sms = "Y" then %>checked<% end if %>>발송 
              					<input name="new_sms" type="radio" value="N" <% if new_sms = "N" then %>checked<% end if %>>발송안함
                                </td>
							</tr>
							<tr>
								<th class="first">장애내용</th>
								<td class="left" colspan="7">
                                <textarea name="as_memo" rows="5" id="textarea"><%=rs("as_memo")%></textarea>
                                </td>
							</tr>
							<tr>
								<th class="first">요청일/시간</th>
								<td class="left">
                                <input name="request_date" type="text" size="10" readonly="true" id="datepicker" style="width:70px;">&nbsp;
                                <input name="request_hh" type="text" id="request_hh" value="<%=mid(rs("request_time"),1,2)%>" size="2" maxlength="2">
                                <strong>시</strong>
                                <input name="request_mm" type="text" id="request_mm" value="<%=mid(rs("request_time"),3,2)%>" size="2" maxlength="2"><strong>분</strong>
							  	</td>
								<th>완료일/시간</th>
								<td class="left">
                                <input name="visit_date" type="text" size="10" readonly="true" id="datepicker1" style="width:70px;">&nbsp;
                                <input name="visit_hh" type="text" id="visit_hh" value="<%=mid(rs("visit_time"),1,2)%>" size="2" maxlength="2">
                                <strong>시</strong>
                                <input name="visit_mm" type="text" id="visit_mm" value="<%=mid(rs("visit_time"),3,2)%>" size="2" maxlength="2"><strong>분</strong>
                                </td>
							  <th>처리유형</th>
								<td class="left">
								<% if (as_type = "신규설치" or as_type = "신규설치공사" or as_type = "이전설치" or as_type = "이전설치공사" or as_type = "랜공사" or as_type = "이전랜공사" or as_type = "장비회수" or as_type = "예방점검") then %>
                                <select name="as_type" id="as_type" style="width:150px" onChange="menu1()">
                                  <option value="<%=as_type%>" <%If as_type = as_type then %>selected<% end if %>><%=as_type%></option>
                                </select>
                                <%   else %>
                                <select name="as_type" id="as_type" style="width:150px" onChange="menu1()">
								  <option value="방문처리" <%If as_type = "방문처리" then %>selected<% end if %>>방문처리</option>
								  <option value="원격처리" <%If as_type = "원격처리" then %>selected<% end if %>>원격처리</option>
								  <option value="신규설치" <%If as_type = "신규설치" then %>selected<% end if %>>신규설치</option>
								  <option value="신규설치공사" <%If as_type = "신규설치공사" then %>selected<% end if %>>신규설치공사</option>
								  <option value="이전설치" <%If as_type = "이전설치" then %>selected<% end if %>>이전설치</option>
								  <option value="이전설치공사" <%If as_type = "이전설치공사" then %>selected<% end if %>>이전설치공사</option>
								  <option value="랜공사" <%If as_type = "랜공사" then %>selected<% end if %>>랜공사</option>
								  <option value="이전랜공사" <%If as_type = "이전랜공사" then %>selected<% end if %>>이전랜공사</option>
								  <option value="장비회수" <%If as_type = "장비회수" then %>selected<% end if %>>장비회수</option>
								  <option value="예방점검" <%If as_type = "예방점검" then %>selected<% end if %>>예방점검</option>
								  <option value="기타" <%If as_type = "기타" then %>selected<% end if %>>기타</option>
							    </select>
			 					<% end if %>
             					<input name="as_type_old" type="hidden" id="as_type_old" value="<%=as_type%>">
                                </td>
								<th>처리현황</th>
								<td class="left">
                                <select name="as_process" style="width:150px" onChange="menu1()">
                                  <option value="접수"  <%If rs("as_process") = "접수" then %>selected<% end if %>>접수</option>
                                  <option value="완료"  <%If rs("as_process") = "완료" then %>selected<% end if %>>완료</option>
                                  <option value="입고"  <%If rs("as_process") = "입고" then %>selected<% end if %>>입고</option>
                                  <option value="연기"  <%If rs("as_process") = "연기" then %>selected<% end if %>>연기</option>
                                  <option value="취소"  <%If rs("as_process") = "취소" then %>selected<% end if %>>취소</option>
                                </select>                
                                <input name="as_process_old" type="hidden" id="as_process_old" value="<%=rs("as_process")%>">
                                </td>
							</tr>
							<tr>
								<th class="first">장애장비</th>
								<td class="left">
                            <%
								Sql="select * from etc_code where etc_type = '31' order by etc_code asc"
								Rs_etc.Open Sql, Dbconn, 1
							%>
								<select name="as_device" id="select" style="width:150px" onChange="menu1()">
                			<% 
								do until rs_etc.eof 
			  				%>
                					<option value='<%=rs_etc("etc_name")%>' <%If rs("as_device") = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			<%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							%>
            					</select>
            					</td>
								<th>제조사</th>
								<td class="left">
                            <%
								Sql="select * from etc_code where etc_type = '21' order by etc_code asc"
								Rs_etc.Open Sql, Dbconn, 1
							%>
              					<select name="maker" id="maker" style="width:150px">
                			<% 
								do until rs_etc.eof 
			  				%>
                					<option value='<%=rs_etc("etc_name")%>' <%If rs("maker") = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			<%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							%>
            					</select>
            					</td>
								<th>모델명</th>
								<td class="left">
                                <input name="model_no" type="text" id="model_no" style="width:150px" onKeyUp="checklength(this,20)" value="<%=rs("model_no")%>">
                            </td>
								<th>시리얼번호</th>
								<td class="left"><input name="serial_no" type="text" id="serial_no" style="width:150px"  onKeyUp="checklength(this,20)" value="<%=rs("serial_no")%>"></td>
							</tr>
							<tr>
								<th class="first">자산번호</th>
								<td class="left"><input name="asets_no" type="text" id="asets_no" style="width:150px" onKeyUp="checklength(this,20)" value="<%=rs("asets_no")%>"></td>
								<th>지연/입고사유</th>
								<td class="left" colspan="5"><textarea name="into_reason"><%=rs("into_reason")%></textarea></td>
							</tr>
							<tr style="display:none" id="in_menu">
								<th class="first" style="background:#FCF">입고일자</th>
								<td class="left"><input name="in_date" type="text" id="datepicker2" size="10" readonly="true"></td>
								<td style="background:#FCF">입고처</td>
								<td class="left">
							  <% if rs("as_process") = "입고" then	%>
                              	<%=in_place%>
                              <%	  else	%>
                              	<select name="in_place" class="style12" id="select2">
                                	<option value="없음">없음</option>
                                	<option value="자체입고">자체입고</option>
                                	<option value="본사입고">본사입고</option>
                                	<option value="Repair Shop">Repair Shop</option>
                              	</select>
                              <% end if %>
                                </td>
								<td style="background:#FCF">대체</td>
								<td class="left">
                				<select name="in_replace" id="in_replace">
                					<option></option>
                					<option value="않함" <%If rs("in_replace") = "않함" then %>selected<% end if %>>않함</option>
                					<option value="대체" <%If rs("in_replace") = "대체" then %>selected<% end if %>>대체</option>
              					</select>
            					</td>
								<td style="background:#FCF">입고진행</td>
								<td class="left"><%=in_process%>
                				<input name="in_process" type="hidden" id="in_process" value="<%= in_process%>">
                                </td>
							</tr>
							<tr style="display:none" id="end_keyin1">
								<th class="first" style="background:#FFC">사용부품</th>
								<td class="left" colspan="7"><input name="as_parts" type="text" id="as_parts" onKeyUp="checklength(this,50)" value="<%=rs("as_parts")%>" size="50"></td>
							</tr>
							<tr style="display:none" id="end_keyin2">
								<th class="first" style="background:#FFC">처리내역</th>
								<td class="left" colspan="7"><textarea name="as_history" rows="2" id="textarea"></textarea></td>
							</tr>
							<tr id="inst_menu" style="display:none">
								<th class="first" colspan="2" style="background:#E1FFE1">설치,이전,공사,회수,예방점검 수량</th>
								<td class="left" colspan="6" bgcolor="#E1FFE1">
								설치대수
                                <input name="dev_inst_cnt" type="text" id="dev_inst_cnt" style="width:30px;text-align:right" onKeyUp="checkNum(this);"  maxlength="3" value="<%=dev_inst_cnt%>">
                                대&nbsp; 
                                공사대수
                                <input name="ran_cnt" type="text" id="ran_cnt" style="width:30px;text-align:right" onKeyUp="checkNum(this);" value="0" maxlength="3">대&nbsp;
                                작업인력
                                <input name="work_man_cnt" type="text" id="work_man_cnt" style="width:30px;text-align:right" value="1" maxlength="2" readonly="true">명&nbsp;
                                알바인원
                                <input name="alba_cnt" type="text" id="alba_cnt" style="width:30px;text-align:right" onKeyUp="checkNum(this);" value="0" maxlength="2">명&nbsp;
								<a href="#" id="work_ce" class="btnType03" onClick="pop_Window('work_ce_add.asp?acpt_no=<%=rs("acpt_no")%>','work_ce_add_pop','scrollbars=yes,width=700,height=500')">2명이상작업인력등록</a>
                                <br><strong>작업인력이 1명인 경우는 설치,공사대수 및 알바인력을 입력 하고, 만약 2명이상인 경우는 2명이상버튼을 눌러 세부사항을 입력한다.</strong>
                                </td>
            				</tr>
						</tbody>
					</table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu1" style="display:none">
<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '01' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '01' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>데스크탑<br>노트북<br>S/W 장애</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err01" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						  <%
                                sql = "select count(*) from etc_code where etc_type = '02' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '02' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#E1FFE1">
                                <p>데스크탑<br>노트북<br>H/W 장애</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#E1FFE1">
                                <input type="checkbox" name="err02" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
                    	</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu2" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '03' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '03' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>모니터 장애</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err03" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu3" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '04' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '04' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>프린터<br>스케너<br>플로터 장애</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err04" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu4" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '05' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '05' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>통신장비<br>네트웍 장애</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err05" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu5" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '06' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '06' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>워크스테이션<br>서버 장애</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err06" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu6" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '07' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '07' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>아답터 장애</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err07" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
                    </table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="end_menu7" style="display:none">
						<colgroup>
							<col width="*" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
						</colgroup>
						<tbody>
							<%
                                sql = "select count(*) from etc_code where etc_type = '09' and used_sw = 'Y'"
                                Set RsCount = Dbconn.Execute (sql)			
                                total_record = cint(RsCount(0)) 'Result.RecordCount
                                rscount.close()
                                
                                SQL="select * from etc_code where etc_type = '09' and used_sw = 'Y'"
                                rs_etc.Open Sql, Dbconn, 1
                                row_span = (total_record -1) / 6 + 1
                            %>
							<tr>
								<th class="first" rowspan="<%=row_span%>" valign="middle" style="background:#FFFFE6">
                                <p>기타</p>
                                </th>
							<%
							row_cnt = 1
							record_cnt = 1
							do until rs_etc.EOF
							%>
								<td class="left" bgcolor="#FFFFE6">
                                <input type="checkbox" name="err09" value="<%=rs_etc("etc_code")%>"><%=rs_etc("etc_name")%>
                                </td>
              				<% 
								if row_cnt = 6 then
									if record_cnt <> total_record then 
							%>
							</tr>
            				<tr>
              				<% 	   
									end if
								end if
 								row_cnt = row_cnt + 1
								record_cnt = record_cnt + 1
								if row_cnt = 7 then
									row_cnt = 1
									end if
								rs_etc.MoveNext()
							loop
							rs_etc.Close()
							%>
            				</tr>
						</tbody>
					</table>
					<table cellpadding="0" cellspacing="0" class="tableWrite" id="att_menu" style="display:none">
						<colgroup>
							<col width="8%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first" valign="middle" style="background:#FFFFE6">파일첨부1</th>
								<td class="left" bgcolor="#FFFFE6"><input name="att_file1" type="file" id="att_file1" size="100"></td>
            				</tr>
							<tr>
								<th class="first" valign="middle" style="background:#FFFFE6">파일첨부2</th>
								<td class="left" bgcolor="#FFFFE6"><input name="att_file2" type="file" id="att_file2" size="100"></td>
            				</tr>
							<tr>
								<th class="first" valign="middle" style="background:#FFFFE6">파일첨부3</th>
								<td class="left" bgcolor="#FFFFE6"><input name="att_file3" type="file" id="att_file3" size="100"></td>
            				</tr>
							<tr>
								<th class="first" valign="middle" style="background:#FFFFE6">파일첨부4</th>
								<td class="left" bgcolor="#FFFFE6"><input name="att_file4" type="file" id="att_file4" size="100"></td>
            				</tr>
							<tr>
								<th class="first" valign="middle" style="background:#FFFFE6">파일첨부5</th>
								<td class="left" bgcolor="#FFFFE6"><input name="att_file5" type="file" id="att_file5" size="100"></td>
            				</tr>
						</tbody>
                    </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="이전" onclick="javascript:goBefore();"></span>
                </div>
                <br>
				<input name="reside_place" type="hidden" id="reside_place" value="<%=rs("reside_place")%>">
                <input name="team" type="hidden" id="team" value="<%=rs("team")%>">
                <input name="sms_old" type="hidden" id="sms_old" value="<%=rs("sms")%>" size="1">
                <input name="be_pg" type="hidden" id="be_pg" value="<%=be_pg%>">
                <input name="write_date" type="hidden" id="write_date" value="<%=rs("write_date")%>">
                <input name="write_cnt" type="hidden" id="write_cnt" value="<%=rs("write_cnt")%>">
                <input name="page" type="hidden" id="page" value="<%=page%>">
                <input name="from_date" type="hidden" id="from_date" value="<%=from_date%>">
                <input name="to_date" type="hidden" id="to_date" value="<%=to_date%>">
                <input name="date_sw" type="hidden" id="date_sw" value="<%=date_sw%>">
                <input name="process_sw" type="hidden" id="process_sw" value="<%=process_sw%>">
                <input name="field_check" type="hidden" id="field_check" value="<%=field_check%>">
                <input name="field_view" type="hidden" id="field_view" value="<%=field_view%>">
                <input name="view_sort" type="hidden" id="view_sort" value="<%=view_sort%>">
                <input name="condi_com" type="hidden" id="condi_com" value="<%=condi_com%>">
                <input name="view_c" type="hidden" id="view_c" value="<%=view_c%>">
                <input name="tel_ddd" type="hidden" id="tel_ddd" value="<%=rs("tel_ddd")%>">
                <input name="tel_no1" type="hidden" id="tel_no1" value="<%=rs("tel_no1")%>">
                <input name="tel_no2" type="hidden" id="tel_no2" value="<%=rs("tel_no2")%>">
                <input name="mg_group" type="hidden" id="mg_group" value="<%=rs("mg_group")%>">
        	</form>
		</div>				
	</div>        				
	</body>
</html>

