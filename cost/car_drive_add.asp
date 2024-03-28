<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->

<!--기존 include 파일 하단 사용으로 수정[허정호_20220302]-->
<!--include virtual="/include/db_create.asp" -->
<!--include virtual="/include/end_check.asp" -->

<!--#include virtual="/common/func.asp" --><!--사용자 정의 함수 : 허정호_20201202-->
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

'누락된 변수 선언 추가(include file : end_check.asp)[허정호_20201202]
'Dim rs_end
'Dim end_saupbu, new_date, end_date

'=========================================================
'미사용 코드 수정(상단 include file)[허정호_20220302]
'Dim sql_trade
'sql_trade="select * from trade where use_sw = 'Y' and ( trade_id = '매출' or trade_id = '공용' ) order by trade_name asc"

Dim end_saupbu, rs_end, end_date, new_date

If saupbu = "" Then
	end_saupbu = "사업부외나머지"
Else
  	end_saupbu = saupbu
End If

objBuilder.Append "SELECT MAX(end_month) AS 'max_month' FROM cost_end "
objBuilder.Append "WHERE saupbu='"&end_saupbu&"' AND end_yn='Y' "

Set rs_end = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rs_end("max_month")) Then
	end_date = "2014-08-31"
Else
	new_date = DateAdd("m",1,DateValue(Mid(rs_end("max_month"),1,4)&"-"&Mid(rs_end("max_month"),5,2)&"-01"))
	end_date = DateAdd("d",-1,new_date)
End If

rs_end.Close():Set rs_end=Nothing

'=========================================================
Dim tRunSQL, tRunRs, rs, rs_memb, rs_next, rs_etc
Dim run_seq
Dim transSQL, rs_tran
Dim rs_car

Dim u_type, mg_ce_id, mg_ce, start_company, start_point
Dim start_hh, start_mm, end_company, end_point, end_km
Dim end_hh, end_mm, far, run_memo, repair_cost
Dim oil_amt, oil_price, parking, toll, end_yn, cancel_yn
Dim curr_date, run_date, strNowWeek, week, company
Dim car_no, car_name, car_owner, oil_kind, last_km
Dim max_km, start_km
Dim end_view, cancel_view
Dim repair_pay, oil_pay, parking_pay, toll_pay
Dim reg_id, reg_date, reg_user, mod_id, mod_date, mod_user
Dim next_km, pre_km
Dim rs_next2
Dim title_line

u_type = f_Request("u_type")

mg_ce_id = user_id
mg_ce = user_name
start_company = ""
start_point = ""
start_hh = ""
start_mm = ""
end_company = ""
end_point = ""
end_km = 0
end_hh = ""
end_mm = ""
far = 0
run_memo = ""
'payment = "현금"
repair_cost = 0
oil_amt = 0
oil_price = 0
parking = 0
toll = 0
end_yn = "N"
cancel_yn = "N"

curr_date = Mid(CStr(Now()), 1, 10)
run_date = Mid(CStr(Now()), 1, 10)

strNowWeek = Weekday(run_date)
Select Case (strNowWeek)
   Case 1
       week = "일요일"
   Case 2
       week = "월요일"
   Case 3
       week = "화요일"
   Case 4
       week = "수요일"
   Case 5
       week = "목요일"
   Case 6
       week = "금요일"
   Case 7
       week = "토요일"
End Select

company = "없음"

If u_type <> "U" Then
	'sql = "select * from car_info where owner_emp_no ='"&emp_no&"' ORDER BY car_owner DESC, car_no ASC"
	'set rs_car=dbconn.execute(sql)
	objBuilder.Append "SELECT car_no, car_name, car_owner, oil_kind, last_km "
	objBuilder.Append "FROM car_info "
	objBuilder.Append "WHERE owner_emp_no ='"&emp_no&"' "
	'처분일자 조건 추가[허정호_20220224]
	objBuilder.Append "	AND (end_date = '' OR end_date IS NULL OR end_date = '1900-01-01')"
	objBuilder.Append "ORDER BY car_owner DESC, car_no ASC "

	Set rs_car = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'코드 수정[허정호_20201202]
	'car_no가 "미등록"인 데이터 DB에서 확인 안됨(입력 페이지 없음)
	'미등록으로 아래 쿼리 실행 시 쿼리 조회 시간이 1분 정보 소요됨(운영에서 쿼리 속도 느림)
	'차량 정보(car_no) 없을 경우 아래 쿼리 실행 안하는 것으로 주석 처리
	If rs_car.EOF Or rs_car.BOF Then
		'car_no = "미등록"
		car_no = ""
		car_name = ""
		car_owner = ""
		oil_kind = ""
		last_km = 0
	Else
		car_no = rs_car("car_no")
		car_name = rs_car("car_name")
		car_owner = rs_car("car_owner")
		oil_kind = rs_car("oil_kind")
		last_km = rs_car("last_km")
	End If

	rs_car.Close():Set rs_car = Nothing

	' 차량 변경시 도착KM,주행거리를 새롭게 입력할것
	'sql = "select car_no, max(end_km) as max_km from transit_cost where car_no = '"&car_no&"'"
	'set rs_tran=dbconn.execute(sql)
	If f_toString(car_no, "") = "" Or IsNull(car_no) Then
		max_km = ""
		start_point = ""
		start_company = ""
	Else
		'objBuilder.Append "SELECT car_no, MAX(end_km) AS max_km "
		'objBuilder.Append "FROM transit_cost "
		'objBuilder.Append "WHERE car_no = '"&car_no&"' "
		objBuilder.Append "SELECT car_no, end_km AS max_km, end_point, end_company "
		objBuilder.Append "FROM transit_cost "
		objBuilder.Append "WHERE mg_ce_id = '"&emp_no&"' "
		objBuilder.Append "	AND car_no = '"&car_no&"' "
		objBuilder.Append "ORDER BY reg_date DESC "
		objBuilder.Append "LIMIT 1 "

		Set rs_tran = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_tran.BOF Or rs_tran.EOF Then
			max_km = ""
		Else
			max_km = rs_tran("max_km")
			start_point = rs_tran("end_point")
			start_company = rs_tran("end_company")
		End If

		rs_tran.close():Set rs_tran = Nothing
	End If

	If max_km = "" Or IsNull(max_km) Then
		last_km = last_km
	Else
		last_km = max_km
	End If

	start_km = last_km
	end_km = last_km

	title_line = "차량 운행일지 등록"
'If u_type = "U" Then
Else
	run_date = f_Request("run_date")
	mg_ce_id = f_Request("mg_ce_id")
	run_seq = f_Request("run_seq")

	'sql = "select * from transit_cost where run_date ='"&run_date&"' and mg_ce_id ='"&mg_ce_id&"' and run_seq ='"&run_seq&"'"
	'set rs = dbconn.execute(sql)

	objBuilder.Append "SELECT car_no, car_name, car_owner, oil_kind, start_company, "
	objBuilder.Append "start_point, start_time, start_km, end_company, end_point, "
	objBuilder.Append "end_time, end_km, far, repair_pay, repair_cost, "
	objBuilder.Append "run_memo, oil_amt, oil_pay, oil_price, parking_pay, "
	objBuilder.Append "parking, toll_pay, toll, cancel_yn, end_yn, "
	objBuilder.Append "reg_id, reg_date, reg_user, mod_id, mod_date, "
	objBuilder.Append "mod_user "
	objBuilder.Append "FROM transit_cost "
	objBuilder.Append "WHERE run_date ='"&run_date&"' "
	objBuilder.Append "AND mg_ce_id ='"&mg_ce_id&"' "
	objBuilder.Append "AND run_seq ='"&run_seq&"' "

	Set rs = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	car_no = rs("car_no")
	car_name = rs("car_name")
	car_owner = rs("car_owner")
	oil_kind = rs("oil_kind")

	start_company = rs("start_company")
	start_point = rs("start_point")
	start_hh = Mid(rs("start_time"), 1, 2)
	start_mm = Mid (rs("start_time"), 3, 2)
	start_km = Int(rs("start_km"))
	end_company = rs("end_company")
	end_point = rs("end_point")
	end_hh = Mid(rs("end_time"), 1, 2)
	end_mm = Mid(rs("end_time"), 3, 2)
	end_km = Int(rs("end_km"))
	far = Int(rs("far"))
'	payment = rs("payment")
	repair_pay = rs("repair_pay")
	repair_cost = Int(rs("repair_cost"))
	run_memo = rs("run_memo")
	oil_amt = Int(rs("oil_amt"))
	oil_pay = rs("oil_pay")
	oil_price = Int(rs("oil_price"))
	parking_pay = rs("parking_pay")
	parking = Int(rs("parking"))
	toll_pay = rs("toll_pay")
	toll = Int(rs("toll"))
	cancel_yn = rs("cancel_yn")
	end_yn = rs("end_yn")
	reg_id = rs("reg_id")
	reg_date = rs("reg_date")
	reg_user = rs("reg_user")
	mod_id = rs("mod_id")
	mod_date = rs("mod_date")
	mod_user = rs("mod_user")

	rs.close() : Set rs = Nothing

	'sql = "select * from memb where user_id = '"&rs("mg_ce_id")&"'"
	'set rs_memb=dbconn.execute(sql)
	objBuilder.Append "SELECT user_name "
	objBuilder.Append "FROM memb "
	objBuilder.Append "WHERE user_id = '"&mg_ce_id&"' "

	Set rs_memb = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_memb.EOF Or rs_memb.BOF Then
		mg_ce = "ERROR"
	Else
		mg_ce = rs_memb("user_name")
	End If

 	rs_memb.close():Set rs_memb = Nothing

	' 차량 운행자가 바뀌는 경우  max(end_km)가 다르다고 문의할 수 있으니 이때는 상관하자말고 출발KM를 시작KM로 새로 등록하면 된다고 안내하면됨..(문의 : 2019-01-04 정구일)
	'sql = "select car_no, max(end_km) as max_km from transit_cost where car_no = '"&car_no&"'"
	'set rs_tran=dbconn.execute(sql)
	objBuilder.Append "SELECT car_no, MAX(end_km) as max_km "
	objBuilder.Append "FROM transit_cost "
	'objBuilder.Append "WHERE car_no = '"& car_no &"' "
	objBuilder.Append "WHERE run_date ='"&run_date&"' "
	objBuilder.Append "	AND mg_ce_id ='"&mg_ce_id&"' "
	objBuilder.Append "	AND run_seq ='"&run_seq&"' "

	Set rs_tran = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	max_km = rs_tran("max_km")

	If max_km = "" Or IsNull(max_km) Then
		last_km = last_km
	Else
		last_km = max_km
	End If

	rs_tran.close():Set rs_tran = Nothing

	'sql = "select * from transit_cost where mg_ce_id ='"&mg_ce_id&"' and start_km >= "&int(end_km)
	'rs_next.Open sql, Dbconn, 1
	objBuilder.Append "SELECT start_km "
	objBuilder.Append "FROM transit_cost "
	'objBuilder.Append "WHERE mg_ce_id ='"&mg_ce_id&"' "
	objBuilder.Append "WHERE run_date ='"&run_date&"' "
	objBuilder.Append "	AND mg_ce_id ='"&mg_ce_id&"' "
	objBuilder.Append "	AND run_seq ='"&run_seq&"' "
	objBuilder.Append "	AND start_km >= "&Int(end_km)

	'rs_next.Open objBuilder.ToString(), DBConn, 1
	Set rs_next=DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_next.EOF Then
		next_km = 999999
	Else
		next_km = rs_next("start_km")
	End If

	rs_next.Close():Set rs_next = Nothing

	'Set rs_next2 = Server.CreateObject("ADODB.RecordSet")

	'sql = "select * from transit_cost where mg_ce_id ='"&mg_ce_id&"' and end_km <= "&int(start_km)&" order by end_km desc"
	'rs_next.Open sql, Dbconn, 1
	objBuilder.Append "SELECT end_km "
	objBuilder.Append "FROM transit_cost "
	'objBuilder.Append "WHERE mg_ce_id ='" & mg_ce_id&"' "
	objBuilder.Append "WHERE run_date ='"&run_date&"' "
	objBuilder.Append "	AND mg_ce_id ='"&mg_ce_id&"' "
	objBuilder.Append "	AND run_seq ='"&run_seq&"' "
	objBuilder.Append "	AND end_km <= " & Int(start_km) & " "
	objBuilder.Append "ORDER BY end_km DESC "

	Set rs_next2 = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_next2.EOF Then
		pre_km = 0
	Else
		pre_km = rs_next2("end_km")
	End If

	rs_next2.Close():Set rs_next2 = Nothing

	title_line = "차량 운행일지 변경"
End If

If end_yn = "Y" Then
	end_view = "마감"
Else
  	end_view = "진행"
End If

If cancel_yn = "Y" Then
	cancel_view = "취소"
Else
  	cancel_view = "지급"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
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
				$( "#datepicker" ).datepicker("setDate", "<%=run_date%>" );
			});

			function goAction(){
			   window.close () ;
			}

			function goBefore(){
			   history.back() ;
			}

			function frmcheck(){
				if (chkfrm()) {
					document.frm.submit();
				}
			}

			function chkfrm(){
				start_km=parseInt(document.frm.start_km.value.replace(/,/g,""));
				end_km=parseInt(document.frm.end_km.value.replace(/,/g,""));
				old_start_km=parseInt(document.frm.old_start_km.value.replace(/,/g,""));
				old_end_km=parseInt(document.frm.old_end_km.value.replace(/,/g,""));
				last_km=parseInt(document.frm.last_km.value.replace(/,/g,""));

				//차량 정보 없을 경우 기본값 공백으로 수정[허정호_20201202]
				//if(document.frm.car_no.value == "미등록"){
				if(document.frm.car_no.value == "미등록" || document.frm.car_no.value == ""){
					alert('등록차량이 없습니다');
					frm.car_no.focus();
					return false;
				}

				if(document.frm.last_km.value == ""){
					alert('최종KM가 없습니다, 차량정보를 변경하시길 바랍니다');
					frm.last_km.focus();
					return false;
				}

				if(document.frm.run_date.value <= document.frm.end_date.value){
					alert('이용일자가 마감이 되어 있는 날자입니다');
					frm.run_date.focus();
					return false;
				}

				if(document.frm.run_date.value > document.frm.curr_date.value){
					alert('이용일자가 현재일보다 클수가 없습니다.');
					frm.run_date.focus();
					return false;
				}

				if(document.frm.start_company.value =="" ){
					alert('출발회사를 선택하세요');
					frm.start_company.focus();
					return false;
				}

				if(document.frm.start_point.value =="" ){
					alert('출발주소을 입력하세요');
					frm.start_point.focus();
					return false;
				}

				if(document.frm.u_type.value !="U" ){
					if(start_km < last_km) {
						alert('출발KM가 최종KM보다 작습니다.');
						frm.start_km.focus();
						return false;
					}
				}

				if(document.frm.u_type.value =="U" ){
					if(start_km < document.frm.pre_km.value){
						alert('출발KM가 이전의 도착KM 작습니다.');
						frm.start_km.focus();
						return false;
					}
				}

				if(document.frm.start_hh.value >"23"||document.frm.start_hh.value <"00"){
					alert('출발시간이 잘못되었습니다');
					frm.start_hh.focus();
					return false;
				}

				if(document.frm.start_mm.value >"59"||document.frm.start_mm.value <"00"){
					alert('출발분이 잘못되었습니다');
					frm.start_mm.focus();
					return false;
				}

				if(document.frm.end_company.value =="" ){
					alert('도착회사를 선택하세요');
					frm.end_company.focus();
					return false;
				}

				if(document.frm.end_point.value =="" ){
					alert('도착주소을 입력하세요');
					frm.end_point.focus();
					return false;
				}

				if(start_km >= end_km) {
					alert('도착KM가 출발KM보다 작습니다.');
					frm.end_km.focus();
					return false;
				}

				if(document.frm.u_type.value =="U" ){
					if(end_km > document.frm.next_km.value){
						alert('도착KM가 다음의 출발KM보다 큽니다');
						frm.end_km.focus();
						return false;
					}
				}

				if(document.frm.end_hh.value >"23"||document.frm.end_hh.value <"00"){
					alert('도착시간이 잘못되었습니다');
					frm.end_hh.focus();
					return false;
				}

				if(document.frm.end_mm.value >"59"||document.frm.end_mm.value <"00"){
					alert('도착분이 잘못되었습니다');
					frm.end_mm.focus();
					return false;
				}

				if(document.frm.start_hh.value > document.frm.end_hh.value){
					alert('도착시간이 출발시간 보다 빠름니다');
					frm.end_hh.focus();
					return false;
				}

				if(document.frm.start_hh.value == document.frm.end_hh.value){
					if(document.frm.start_mm.value > document.frm.end_mm.value){
						alert('도착시간이 출발시간 보다 빠름니다');
						frm.end_mm.focus();
						return false;
					}
				}

				if(document.frm.run_memo.value =="" ){
					alert('운행목적을 선택하세요');
					frm.run_memo.focus();
					return false;
				}

				if(document.frm.oil_amt.value == 0){
					if(document.frm.oil_price.value > 0) {
						alert('주유량이 없는데 주유금액이 있습니다.');
						frm.oil_amt.focus();
						return false;
					}
				}

				if(document.frm.oil_amt.value > 0){
					if(document.frm.oil_price.value == 0){
						alert('주유량이 있는데 주유금액이 없습니다.');
						frm.oil_price.focus();
						return false;
					}
				}

				{
					a = confirm('입력하시겠습니까?');
					if (a==true){
						return true;
					}
					return false;
				}
			}

			function week_check(){
				a = document.frm.run_date.value.substring(0,4);
				b = document.frm.run_date.value.substring(5,7);
				c = document.frm.run_date.value.substring(8,10);

				var newDate = new Date(a,b-1,c);
				var s = newDate.getDay();

				switch(s) {
					case 0: str = "일요일" ; break;
					case 1: str = "월요일" ; break;
					case 2: str = "화요일" ; break;
					case 3: str = "수요일" ; break;
					case 4: str = "목요일" ; break;
					case 5: str = "금요일" ; break;
					case 6: str = "토요일" ; break;
				}

				document.frm.week.value = str;
			}

			function payment_view(){
				var c = document.frm.oil_pay.value;

				if (c == '현금'){
					document.getElementById("oil_price").readOnly = true;
					document.frm.oil_price.value = 0;
				}

				if (c == '법인카드'){
					document.getElementById("oil_price").readOnly = "";
				}
			}

			function km_cal(txtObj){
				if (txtObj.value.length<5){
					txtObj.value=txtObj.value.replace(/,/g,"");
					txtObj.value=txtObj.value.replace(/\D/g,"");
					start_km=parseInt(document.frm.start_km.value.replace(/,/g,""));
					end_km=parseInt(document.frm.end_km.value.replace(/,/g,""));
					document.frm.far.value = end_km - start_km;
				}

				var num = txtObj.value;

				if (num == "--" ||  num == "." ) num = "";

				if (num != "" ){
					temp=new String(num);
					if(temp.length<1) return "";

					// 음수처리
					if(temp.substr(0,1)=="-") minus="-";
						else minus="";

					// 소수점이하처리
					dpoint=temp.search(/\./);

					if(dpoint>0)
					{
					// 첫번째 만나는 .을 기준으로 자르고 숫자제외한 문자 삭제
					dpointVa="."+temp.substr(dpoint).replace(/\D/g,"");
					temp=temp.substr(0,dpoint);
					}else dpointVa="";

					// 숫자이외문자 삭제
					temp=temp.replace(/\D/g,"");
					zero=temp.search(/[1-9]/);

					if(zero==-1) return "";
					else if(zero!=0) temp=temp.substr(zero);

					if(temp.length<4) return minus+temp+dpointVa;
					buf="";

					while (true){
						if(temp.length<3) { buf=temp+buf; break; }

						buf=","+temp.substr(temp.length-3)+buf;
						temp=temp.substr(0, temp.length-3);
					}

					if(buf.substr(0,1)==",") buf=buf.substr(1);

					//return minus+buf+dpointVa;
					txtObj.value = minus+buf+dpointVa;

					start_km=parseInt(document.frm.start_km.value.replace(/,/g,""));
					end_km=parseInt(document.frm.end_km.value.replace(/,/g,""));
					document.frm.far.value = end_km - start_km;

				}else txtObj.value = "0";
			}

			function update_view(){
				var c = document.frm.u_type.value;

				if (c == 'U')
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}

			function delcheck(){
				a = confirm('정말 삭제하시겠습니까?');

				if (a==true){
					document.frm.action = "/cost/car_drive_del_ok.asp";
					document.frm.submit();
					return true;
				}
				return false;
			}
        </script>
	</head>
	<body onLoad="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/cost/car_drive_add_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="13%" >
							<col width="37%" >
							<col width="13%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">운행일</th>
								<td class="left">
                                <input name="run_date" type="text" id="datepicker" style="width:70px" value="<%=run_date%>" readonly="true">&nbsp;
                                마감일자 : <%=end_date%>
							<%  If u_type = "U" Then	%>
                                <input name="old_date" type="hidden" value="<%=run_date%>">
                            <%	End If	%>
                                </td>
								<th>운행자</th>
								<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)
                                <input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=mg_ce_id%>">
                                </td>
							</tr>
							<tr>
								<th class="first">차량정보</th>
								<td colspan="3" class="left">
                                <strong>소유 :</strong><input name="car_owner" type="text" id="car_owner" style="width:30px" value="<%=car_owner%>" readonly="true">&nbsp;
                                <strong>차량번호 :</strong><input name="car_no" type="text" id="car_no" style="width:70px" value="<%=car_no%>" readonly="true">&nbsp;
                                <strong>차종 :</strong><input name="car_name" type="text" id="car_name" style="width:90px" value="<%=car_name%>" readonly="true">&nbsp;
                                <strong>유종 :</strong><input name="oil_kind" type="text" id="oil_kind" style="width:50px" value="<%=oil_kind%>" readonly="true">&nbsp;
                                <strong>최종KM :</strong><input name="last_km" type="text" id="last_km" style="width:50px" value="<%=FormatNumber(last_km, 0)%>" readonly="true"><a href="#" class="btnType03" onClick="pop_Window('car_search.asp','car_search_pop','scrollbars=yes,width=600,height=300')">차량조회</a><br><br><strong>* 차량 조회시 정보가 없는 경우는 회사차량 배정이 안되어 있어 인사총무팀 차량 담당자에 문의 바랍니다.</strong>
                                </td>
						    </tr>
							<tr>
								<th class="first">출발회사</th>
								<td class="left">
								  <%
								  	Dim rsStartCompany
									'Sql="select * from trade where (trade_id = '매출' or trade_id = '공용')  and use_sw = 'Y' order by trade_name asc"
									'Rs_etc.Open Sql, Dbconn, 1
									objBuilder.Append "SELECT trade_name "
									objBuilder.Append "FROM trade "
									objBuilder.Append "WHERE (trade_id='매출' OR trade_id='공용') "
									objBuilder.Append "	AND use_sw = 'Y' "
									objBuilder.Append "ORDER BY trade_name ASC "

									'Rs_etc.Open objBuilder.ToString(), DBConn, 1
									Set rsStartCompany=DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()
								  %>
                                  <select name="start_company" id="select" style="width:150px">
                                    <option value="">선택</option>
                                    <option value='집' <%If start_company = "집" Then %>selected<% End If %>>집</option>
                                    <option value='본사(회사)' <%If start_company = "본사(회사)" Then %>selected<% End If %>>본사(회사)</option>
                                    <%
                                        Do Until rsStartCompany.EOF
                                    %>
                                    <option value='<%=rsStartCompany("trade_name")%>' <%If rsStartCompany("trade_name") = start_company Then %>selected<% End If %>><%=rsStartCompany("trade_name")%></option>
                                    <%
                                        	rsStartCompany.MoveNext()
                                        Loop

                                        rsStartCompany.Close():Set rsStartCompany = Nothing
                                    %>
                                  </select>
                                </td>
								<th>출발주소</th>
								<td class="left"><input name="start_point" type="text" id="start_point" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50)" value="<%=start_point%>"></td>
							</tr>
							<tr>
								<th class="first">출발KM</th>
								<td class="left"><input name="start_km" type="text" id="start_km" style="width:55px;text-align:right" value="<%=formatnumber(start_km,0)%>" onKeyUp="km_cal(this);"></td>
								<th>출발시간</th>
								<td class="left">
                                <input name="start_hh" type="text" id="start_hh" size="2" maxlength="2" value="<%=start_hh%>">시
								<input name="start_mm" type="text" id="start_mm" size="2" maxlength="2" value="<%=start_mm%>">분
								</td>
							</tr>
							<tr>
								<th class="first">도착회사</th>
								<td class="left">
								<%
								Dim rsEndCompany
								'Sql="select * from trade where (trade_id = '매출' or trade_id = '공용')  and use_sw = 'Y' order by trade_name asc"									'Rs_etc.Open Sql, Dbconn, 1
								objBuilder.Append "SELECT trade_name "
								objBuilder.Append "FROM trade "
								objBuilder.Append "WHERE (trade_id='매출' OR trade_id='공용') "
								objBuilder.Append "	AND use_sw = 'Y' "
								'소속 본부와 관리사업부가 동일한 거래처만 노출 처리[허정호_20220224]
								objBuilder.Append "	AND saupbu='"&bonbu&"' "
								objBuilder.Append "ORDER BY trade_name ASC "

								'RS_etc.Open objBuilder.ToString(), DBConn, 1
								Set rsEndCompany = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
                                %>
									<select name="end_company" id="select" style="width:150px">
										<option value="">선택</option>
										<option value='본사(회사)' <%If end_company = "본사(회사)" Then %>selected<% End If %>>본사(회사)</option>
										<option value='집' <%If end_company = "집" Then %>selected<% End If %>>집</option>
									<%
									Do Until rsEndCompany.EOF
									%>
											<option value='<%=rsEndCompany("trade_name")%>' <%If rsEndCompany("trade_name") = end_company Then %>selected<% End If %>><%=rsEndCompany("trade_name")%></option>
									<%
										rsEndCompany.MoveNext()
									Loop

									rsEndCompany.Close():Set rsEndCompany = Nothing
									%>
									</select>
                                </td>
								<th>도착주소</th>
								<td class="left"><input name="end_point" type="text" id="end_point" style="width:200px; ime-mode:active" onKeyUp="checklength(this,50)" value="<%=end_point%>"></td>
							</tr>
							<tr>
								<th class="first">도착KM</th>
								<td class="left"><input name="end_km" type="text" id="end_km" style="width:55px;text-align:right" value="<%=formatnumber(end_km,0)%>" onKeyUp="km_cal(this);"></td>
								<th>도착시간</th>
								<td class="left">
                                <input name="end_hh" type="text" id="end_hh" size="2" maxlength="2" value="<%=end_hh%>">시
								<input name="end_mm" type="text" id="end_mm" size="2" maxlength="2" value="<%=end_mm%>">분
								</td>
							</tr>
					    	<tr>
								<th class="first">주행거리</th>
								<td class="left"><input name="far" type="text" id="far" style="width:50px;text-align:right" value="<%=FormatNumber(far, 0)%>" readonly="true"></td>
								<th>운행목적</th>
								<td class="left">
								<%
									'Sql="select * from etc_code where etc_type = '42' and used_sw = 'Y' order by etc_code asc"
									'Rs_etc.Open Sql, Dbconn, 1
									objBuilder.Append "SELECT etc_name "
									objBuilder.Append "FROM etc_code "
									objBuilder.Append "WHERE etc_type='42' "
									objBuilder.Append "AND used_sw='Y' "
									objBuilder.Append "ORDER BY etc_code ASC "

									Set rs_etc = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()
                                %>
                                  <select name="run_memo" id="run_memo" style="width:150px">
                                    <option value="">선택</option>
                                    <%
                                        Do Until rs_etc.EOF
                                    %>
                                    <option value='<%=rs_etc("etc_name")%>' <%If rs_etc("etc_name") = run_memo Then %>selected<% End If %>><%=rs_etc("etc_name")%></option>
                                    <%
                                        	rs_etc.MoveNext()
                                        Loop
                                        rs_etc.Close():Set rs_etc = Nothing
										DBConn.Close():Set DBConn = Nothing
                                    %>
                                </select></td>
							</tr>
							<tr>
								<th class="first">주유량(L)</th>
								<td class="left">
							<% If u_type = "U" Then	%>
                                <input name="oil_amt" type="text" id="oil_amt" style="width:80px;text-align:right" value="<%=formatnumber(oil_amt,0)%>" onKeyUp="plusComma(this);" >
							<%   Else	%>
                                <input name="oil_amt" type="text" id="oil_amt" style="width:80px;text-align:right" onKeyUp="plusComma(this);" >
							<% End If	%>
                                </td>
                                <th>회사차량<br>주유금액</th>
								<td class="left">현금 또는 개인카드
								  <select name="oil_pay" id="select" style="width:80px" onChange="payment_view()">
                                    <option value='현금' <%If oil_pay= "현금" Then %>selected<% End If %>>현금</option>
                                </select>
							<% If u_type = "U" Then	%>
                                <input name="oil_price" type="text" id="oil_price" style="width:80px;text-align:right" value="<%=formatnumber(oil_price,0)%>" onKeyUp="plusComma(this);">
							<%   Else	%>
                                <input name="oil_price" type="text" id="oil_price" style="width:80px;text-align:right" onKeyUp="plusComma(this);">
							<% End If	%>
                                </td>
							</tr>
							<tr>
								<th class="first">주차비</th>
								<td class="left">지불방법
                                  <select name="parking_pay" id="parking_pay" style="width:80px">
                                    <option value='현금' <%If parking_pay= "현금" Then %>selected<% End If %>>현금</option>
                            	</select>
							<% If u_type = "U" Then	%>
                            	<input name="parking" type="text" id="parking" style="width:80px;text-align:right" value="<%=formatnumber(parking,0)%>" onKeyUp="plusComma(this);" >
							<%   Else	%>
                            	<input name="parking" type="text" id="parking" style="width:80px;text-align:right" onKeyUp="plusComma(this);" >
							<% End If	%>
                                </td>
                                <th>통행료</th>
								<td class="left">지불방법
                                <select name="toll_pay" id="toll_pay" style="width:80px">
                                    <option value='현금' <%If toll_pay= "현금" Then %>selected<% End If %>>현금</option>
                              	</select>
							<% If u_type = "U" Then	%>
                                <input name="toll" type="text" id="toll" style="width:80px;text-align:right" value="<%=FormatNumber(toll, 0)%>" onKeyUp="plusComma(this);" >
							<%   Else	%>
                                <input name="toll" type="text" id="toll" style="width:80px;text-align:right" onKeyUp="plusComma(this);" >
							<% End If	%>
                                </td>
							</tr>
    				  <tr id="cancel_col" style="display:none">
						<th class="first">취소여부</th>
						<td class="left"><%=cancel_view%></td>
                        <th>마감여부</th>
						<td class="left"><%=end_view%></td>
					</tr>
					<tr id="info_col" style="display:none">
						<th class="first">등록정보</th>
						<td class="left"><%=reg_user%>&nbsp;<%=reg_id%>(<%=reg_date%>)</td>
                    	<th>변경정보</th>
						<td class="left"><%=mod_user%>&nbsp;<%=mod_id%>(<%=mod_date%>)</td>
					</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();"/></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"/></span>
				<%
					If u_type = "U" And user_id = mg_ce_id Then
						If end_yn = "N" Or end_yn = "C" Then
				%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();"/></span>
        		<%
						End If
					End If
				%>
                </div>
				<br>
				<input type="hidden" name="u_type" value="<%=u_type%>"/>
                <input type="hidden" name="old_start_km" value="<%=start_km%>"/>
                <input type="hidden" name="old_end_km" value="<%=end_km%>"/>
                <input type="hidden" name="curr_date" value="<%=curr_date%>"/>
                <input type="hidden" name="end_date" value="<%=end_date%>"/>
                <input type="hidden" name="end_yn" value="<%=end_yn%>"/>
				<input type="hidden" name="run_seq" value="<%=run_seq%>"/>
				<input type="hidden" name="cancel_yn" value="<%=cancel_yn%>"/>
                <input type="hidden" name="mod_id" value="<%=mod_id%>"/>
                <input type="hidden" name="mod_user" value="<%=mod_user%>"/>
                <input type="hidden" name="mod_date" value="<%=mod_date%>"/>
                <input type="hidden" name="next_km" value="<%=next_km%>"/>
                <input type="hidden" name="pre_km" value="<%=pre_km%>"/>
			</form>
		</div>
	</body>
</html>