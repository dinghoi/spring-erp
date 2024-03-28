<!--#include virtual="/common/inc_top.asp" -->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'DB Connection
'===================================================
'db_create.asp(include)에 정의됨

'===================================================
'StringBuilder Object
'===================================================
Dim objBuilder

'StringBuffer 형식 사용[허정호_20201123]
Set objBuilder = New StringBuilder

'===================================================
'Request & Param
'===================================================

'end_check.asp(include) 변수 선언[허정호_20201124]
Dim rs, Rs_acc, rs_trade, rs_reside, Rs_type
Dim rs_org, rs_etc, rs_memb, rs_emp, Rs_ddd
Dim RsCount, Rs_in, rs_hol, rs_next, rs_pre
Dim sql_trade
Dim end_saupbu

Dim tRunSQL, tRunRs
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
Dim title_line

u_type = Request("u_type")

'차량운행일지등록
If toStrings(u_type, "") = "" Then
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
	repair_cost = 0
	oil_amt = 0
	oil_price = 0
	parking = 0
	toll = 0
	end_yn = "N"
	cancel_yn = "N"

	curr_date = Mid(CStr(Now()), 1, 10)
	run_date = Mid(CStr(Now()), 1, 10)
	strNowWeek = WeekDay(run_date)

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

	'sql = "select * from car_info where owner_emp_no ='"&emp_no&"' ORDER BY car_owner DESC, car_no ASC"
	objBuilder.Append "SELECT car_no, car_name, car_owner, oil_kind, last_km, "
	objBuilder.Append "(SELECT MAX(end_km) FROM transit_cost WHERE car_no=ci.car_no) AS max_km "
	objBuilder.Append "FROM car_info AS ci "
	objBuilder.Append "WHERE owner_emp_no = '" & emp_no & "' "
	objBuilder.Append "ORDER BY car_owner DESC, car_no ASC "
	objBuilder.Append "LIMIT 1; "

	'차량 변경시 도착KM,주행거리를 새롭게 입력할것 -> 기존 사용 주석[허정호_20201123]
	'기존 쿼리 조회 시 여러개일 경우 첫번째 row만 변수에 담음 -> row 1개만 조회토록 쿼리 수정[허정호_20201123]
	Set rs_car = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	car_no = rs_car("car_no")
	car_name = rs_car("car_name")
	car_owner = rs_car("car_owner")
	oil_kind = rs_car("oil_kind")
	last_km = rs_car("last_km")
	max_km = rs_car("max_km")

	rs_car.Close()
	Set rs_car = Nothing

	If max_km = "" Or IsNull(max_km) Then
		last_km = last_km
	Else
		last_km = max_km
	End If

	start_km = last_km
	end_km = last_km

	objBuilder.Append "SELECT end_point, end_company "
	objBuilder.Append "FROM transit_cost "
	objBuilder.Append "WHERE car_no = '" & car_no & "' AND end_km = '" & end_km & "' "

	Set rs_tran = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rs_tran.EOF Or rs_tran.BOF Then
		start_point = ""
		start_company = ""
	Else
		start_point = rs_tran("end_point")
		start_company = rs_tran("end_company")
	End If

	rs_tran.Close()
	Set rs_tran = Nothing

	title_line = "차량 운행일지 등록"

'If u_type = "U" Then
Else '차량 운행일지 수정
	run_date = Request("run_date")
	mg_ce_id = Request("mg_ce_id")
	run_seq = Request("run_seq")

	'sql = "select * from transit_cost where run_date ='"&run_date&"' and mg_ce_id ='"&mg_ce_id&"' and run_seq ='"&run_seq&"'"
	'set rs = dbconn.execute(sql)

	'sql = "select * from memb where user_id = '"&rs("mg_ce_id")&"'"
	'set rs_memb=dbconn.execute(sql)

	'if	rs_memb.eof or rs_memb.bof then
	'	mg_ce = "ERROR"
	' else
	'	mg_ce = rs_memb("user_name")
	'end if
	'rs_memb.close()
	objBuilder.Append "SELECT tc.car_no, tc.car_name, tc.car_owner, tc.oil_kind, "
	objBuilder.Append "tc.start_company, tc.start_point, tc.start_time, tc.start_km, tc.end_company, "
	objBuilder.Append "tc.end_point, tc.end_time, tc.end_km, tc.far, tc.repair_pay, "
	objBuilder.Append "tc.repair_cost, tc.run_memo, tc.oil_amt, tc.oil_pay, tc.oil_price, "
	objBuilder.Append "tc.parking_pay, tc.parking, tc.toll_pay, tc.toll, tc.cancel_yn, "
	objBuilder.Append "tc.end_yn, tc.reg_id, tc.reg_date, tc.reg_user, tc.mod_id, "
	objBuilder.Append "tc.mod_date, tc.mod_user, "

	objBuilder.Append "mb.user_name "
	objBuilder.Append "FROM transit_cost AS tc "
	objBuilder.Append "INNER JOIN memb AS mb ON tc.mg_ce_id = mb.user_id "
	objBuilder.Append "WHERE tc.run_date ='" & run_date & "' "
	objBuilder.Append "AND tc.mg_ce_id ='" & mg_ce_id & "' "
	objBuilder.Append "AND tc.run_seq ='" & run_seq & "' "

	Set tRunRs = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	mg_ce = tRunRs("user_name")
	If mg_ce = "" Or IsNull(mg_ce) Then
		mg_ce = "ERROR"
	End If

	car_no = tRunRs("car_no")
	car_name = tRunRs("car_name")
	car_owner = tRunRs("car_owner")
	oil_kind = tRunRs("oil_kind")

	start_company = tRunRs("start_company")
	start_point = tRunRs("start_point")
	start_hh = Mid(tRunRs("start_time"), 1, 2)
	start_mm = Mid(tRunRs("start_time"), 3, 2)
	start_km = Int(tRunRs("start_km"))
	end_company = tRunRs("end_company")
	end_point = tRunRs("end_point")
	end_hh = Mid(tRunRs("end_time"), 1, 2)
	end_mm = Mid(tRunRs("end_time"), 3, 2)
	end_km = Int(tRunRs("end_km"))
	far = Int(tRunRs("far"))
	repair_pay = tRunRs("repair_pay")
	repair_cost = Int(tRunRs("repair_cost"))
	run_memo = tRunRs("run_memo")
	oil_amt = int(tRunRs("oil_amt"))
	oil_pay = tRunRs("oil_pay")
	oil_price = Int(tRunRs("oil_price"))
	parking_pay = tRunRs("parking_pay")
	parking = Int(tRunRs("parking"))
	toll_pay = tRunRs("toll_pay")
	toll = Int(tRunRs("toll"))
	cancel_yn = tRunRs("cancel_yn")
	end_yn = tRunRs("end_yn")
	reg_id = tRunRs("reg_id")
	reg_date = tRunRs("reg_date")
	reg_user = tRunRs("reg_user")
	mod_id = tRunRs("mod_id")
	mod_date = tRunRs("mod_date")
	mod_user = tRunRs("mod_user")

	tRunRs.Close()
	Set tRunRs = Nothing

	' 차량 운행자가 바뀌는 경우  max(end_km)가 다르다고 문의할 수 있으니 이때는 상관하자말고 출발KM를 시작KM로 새로 등록하면 된다고 안내하면됨..(문의 : 2019-01-04 정구일)
	'sql = "select car_no, max(end_km) as max_km from transit_cost where car_no = '"&car_no&"'"
	'set rs_tran=dbconn.execute(sql)
	'max_km = rs_tran("max_km")
	max_km = tRunRs("max_km")

	If max_km = "" Or IsNull(max_km) Then
		last_km = last_km
	Else
		last_km = max_km
	End If
	'rs_tran.close()
	response.end

	sql = "select * from transit_cost where mg_ce_id ='"&mg_ce_id&"' and start_km >= "&int(end_km)
	rs_next.Open sql, Dbconn, 1
	if rs_next.eof then
		next_km = 999999
	  else
		next_km = rs_next("start_km")
	end if
	rs_next.close()

	sql = "select * from transit_cost where mg_ce_id ='"&mg_ce_id&"' and end_km <= "&int(start_km)&" order by end_km desc"
	rs_next.Open sql, Dbconn, 1
	if rs_next.eof then
		pre_km = 0
	  else
		pre_km = rs_next("end_km")
	end if
	rs_next.close()

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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=run_date%>" );
			});

			function goAction () {
			   window.close () ;
			}

			function goBefore () {
			   history.back() ;
			}

			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				start_km=parseInt(document.frm.start_km.value.replace(/,/g,""));
				end_km=parseInt(document.frm.end_km.value.replace(/,/g,""));
				old_start_km=parseInt(document.frm.old_start_km.value.replace(/,/g,""));
				old_end_km=parseInt(document.frm.old_end_km.value.replace(/,/g,""));
				last_km=parseInt(document.frm.last_km.value.replace(/,/g,""));

				if(document.frm.car_no.value == "미등록") {
					alert('등록차량이 없습니다');
					frm.car_no.focus();
					return false;}
				if(document.frm.last_km.value == "") {
					alert('최종KM가 없습니다, 차량정보를 변경하시길 바랍니다');
					frm.last_km.focus();
					return false;}
				if(document.frm.run_date.value <= document.frm.end_date.value) {
					alert('이용일자가 마감이 되어 있는 날자입니다');
					frm.run_date.focus();
					return false;}
				if(document.frm.run_date.value > document.frm.curr_date.value) {
					alert('이용일자가 현재일보다 클수가 없습니다.');
					frm.run_date.focus();
					return false;}
				if(document.frm.start_company.value =="" ) {
					alert('출발회사를 선택하세요');
					frm.start_company.focus();
					return false;}
				if(document.frm.start_point.value =="" ) {
					alert('출발주소을 입력하세요');
					frm.start_point.focus();
					return false;}
				if(document.frm.u_type.value !="U" ) {
					if(start_km < last_km) {
						alert('출발KM가 최종KM보다 작습니다.');
						frm.start_km.focus();
						return false;}}
				if(document.frm.u_type.value =="U" ) {
					if(start_km < document.frm.pre_km.value) {
						alert('출발KM가 이전의 도착KM 작습니다.');
						frm.start_km.focus();
						return false;}}
				if(document.frm.start_hh.value >"23"||document.frm.start_hh.value <"00") {
					alert('출발시간이 잘못되었습니다');
					frm.start_hh.focus();
					return false;}
				if(document.frm.start_mm.value >"59"||document.frm.start_mm.value <"00") {
					alert('출발분이 잘못되었습니다');
					frm.start_mm.focus();
					return false;}
				if(document.frm.end_company.value =="" ) {
					alert('도착회사를 선택하세요');
					frm.end_company.focus();
					return false;}
				if(document.frm.end_point.value =="" ) {
					alert('도착주소을 입력하세요');
					frm.end_point.focus();
					return false;}
				if(start_km >= end_km) {
					alert('도착KM가 출발KM보다 작습니다.');
					frm.end_km.focus();
					return false;}
				if(document.frm.u_type.value =="U" ) {
					if(end_km > document.frm.next_km.value) {
						alert('도착KM가 다음의 출발KM보다 큽니다');
						frm.end_km.focus();
						return false;}}
				if(document.frm.end_hh.value >"23"||document.frm.end_hh.value <"00") {
					alert('도착시간이 잘못되었습니다');
					frm.end_hh.focus();
					return false;}
				if(document.frm.end_mm.value >"59"||document.frm.end_mm.value <"00") {
					alert('도착분이 잘못되었습니다');
					frm.end_mm.focus();
					return false;}
				if(document.frm.start_hh.value > document.frm.end_hh.value) {
					alert('도착시간이 출발시간 보다 빠름니다');
					frm.end_hh.focus();
					return false;}
				if(document.frm.start_hh.value == document.frm.end_hh.value) {
					if(document.frm.start_mm.value > document.frm.end_mm.value) {
						alert('도착시간이 출발시간 보다 빠름니다');
						frm.end_mm.focus();
						return false;}}
				if(document.frm.run_memo.value =="" ) {
					alert('운행목적을 선택하세요');
					frm.run_memo.focus();
					return false;}
				if(document.frm.oil_amt.value == 0) {
					if(document.frm.oil_price.value > 0) {
						alert('주유량이 없는데 주유금액이 있습니다.');
						frm.oil_amt.focus();
						return false;}}
				if(document.frm.oil_amt.value > 0) {
					if(document.frm.oil_price.value == 0) {
						alert('주유량이 있는데 주유금액이 없습니다.');
						frm.oil_price.focus();
						return false;}}
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}

			function week_check() {

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

			function payment_view() {
			var c = document.frm.oil_pay.value;
				if (c == '현금')
				{
					document.getElementById("oil_price").readOnly = true;
					document.frm.oil_price.value = 0;
				}
				if (c == '법인카드')
				{
					document.getElementById("oil_price").readOnly = "";
				}
			}

			function km_cal(txtObj){
				if (txtObj.value.length<5) {
					txtObj.value=txtObj.value.replace(/,/g,"");
					txtObj.value=txtObj.value.replace(/\D/g,"");
					start_km=parseInt(document.frm.start_km.value.replace(/,/g,""));
					end_km=parseInt(document.frm.end_km.value.replace(/,/g,""));
					document.frm.far.value = end_km - start_km;
				}
				var num = txtObj.value;
				if (num == "--" ||  num == "." ) num = "";
				if (num != "" ) {
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
					while (true)
					{
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

			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U')
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}

			function delcheck(){
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.action = "car_drive_del_ok.asp";
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
				<form action="car_drive_add_save.asp" method="post" name="frm">
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
							<%If u_type = "U" Then%>
                                <input name="old_date" type="hidden" value="<%=run_date%>">
                            <%End If%>
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
                                <strong>최종KM :</strong><input name="last_km" type="text" id="last_km" style="width:50px" value="<%=formatnumber(last_km,0)%>" readonly="true"><a href="#" class="btnType03" onClick="pop_Window('car_search.asp','car_search_pop','scrollbars=yes,width=600,height=300')">차량조회</a><br><br><strong>* 차량 조회시 정보가 없는 경우는 회사차량 배정이 안되어 있어 인사총무팀 차량 담당자에 문의 바랍니다.</strong>
                                </td>
						    </tr>
							<tr>
								<th class="first">출발회사</th>
								<td class="left">
								  <%
                                        'Sql="select * from trade where (trade_id = '매출' or trade_id = '공용')  and use_sw = 'Y' order by trade_name asc"
										objBuilder.Append "SELECT trade_name "
										objBuilder.Append "FROM trade "
										objBuilder.Append "WHERE (trade_id = '매출' OR trade_id = '공용') AND use_sw = 'Y' "
										objBuilder.Append "ORDER BY trade_name ASC "

                                        'Rs_etc.Open Sql, Dbconn, 1
										Rs_etc.Open objBuilder.ToString(), DBConn, 1
										objBuilder.Clear()
                                    %>
                                  <select name="start_company" id="select" style="width:150px">
                                    <option value="">선택</option>
                                    <option value='집' <%If start_company = "집" Then %>selected<% End If %>>집</option>
                                    <option value='본사(회사)' <%If start_company = "본사(회사)" Then %>selected<% End If %>>본사(회사)</option>
                                    <%
                                        Do Until rs_etc.EOF
                                    %>
                                    <option value='<%=rs_etc("trade_name")%>' <%If rs_etc("trade_name") = start_company Then %>selected<% End If %>><%=rs_etc("trade_name")%></option>
                                    <%
                                        	rs_etc.MoveNext()
                                        Loop

                                        rs_etc.Close()
										Set rs_etc = Nothing
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
                                        Sql="select * from trade where (trade_id = '매출' or trade_id = '공용')  and use_sw = 'Y' order by trade_name asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                    %>
                                  <select name="end_company" id="select" style="width:150px">
                                    <option value="">선택</option>
                                    <option value='본사(회사)' <%If end_company = "본사(회사)" then %>selected<% end if %>>본사(회사)</option>
                                    <option value='집' <%If end_company = "집" then %>selected<% end if %>>집</option>
                                    <%
                                        do until rs_etc.eof
                                    %>
                                    <option value='<%=rs_etc("trade_name")%>' <%If rs_etc("trade_name") = end_company then %>selected<% end if %>><%=rs_etc("trade_name")%></option>
                                    <%
                                        	rs_etc.movenext()
                                        loop
                                        rs_etc.Close()
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
								<td class="left"><input name="far" type="text" id="far" style="width:50px;text-align:right" value="<%=formatnumber(far,0)%>" readonly="true"></td>
								<th>운행목적</th>
								<td class="left"><%
                                        Sql="select * from etc_code where etc_type = '42' and used_sw = 'Y' order by etc_code asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                    %>
                                  <select name="run_memo" id="run_memo" style="width:150px">
                                    <option value="">선택</option>
                                    <%
                                        do until rs_etc.eof
                                    %>
                                    <option value='<%=rs_etc("etc_name")%>' <%If rs_etc("etc_name") = run_memo then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                                    <%
                                        	rs_etc.movenext()
                                        loop
                                        rs_etc.Close()
                                    %>
                                </select></td>
							</tr>
							<tr>
								<th class="first">주유량(L)</th>
								<td class="left">
							<% if u_type = "U" then	%>
                                <input name="oil_amt" type="text" id="oil_amt" style="width:80px;text-align:right" value="<%=formatnumber(oil_amt,0)%>" onKeyUp="plusComma(this);" >
							<%   else	%>
                                <input name="oil_amt" type="text" id="oil_amt" style="width:80px;text-align:right" onKeyUp="plusComma(this);" >
							<% end if	%>
                                </td>
                                <th>회사차량<br>주유금액</th>
								<td class="left">현금 또는 개인카드
								  <select name="oil_pay" id="select" style="width:80px" onChange="payment_view()">
                                    <option value='현금' <%If oil_pay= "현금" then %>selected<% end if %>>현금</option>
                                </select>
							<% if u_type = "U" then	%>
                                <input name="oil_price" type="text" id="oil_price" style="width:80px;text-align:right" value="<%=formatnumber(oil_price,0)%>" onKeyUp="plusComma(this);">
							<%   else	%>
                                <input name="oil_price" type="text" id="oil_price" style="width:80px;text-align:right" onKeyUp="plusComma(this);">
							<% end if	%>
                                </td>
							</tr>
							<tr>
								<th class="first">주차비</th>
								<td class="left">지불방법
                                  <select name="parking_pay" id="parking_pay" style="width:80px">
                                    <option value='현금' <%If parking_pay= "현금" then %>selected<% end if %>>현금</option>
                            	</select>
							<% if u_type = "U" then	%>
                            	<input name="parking" type="text" id="parking" style="width:80px;text-align:right" value="<%=formatnumber(parking,0)%>" onKeyUp="plusComma(this);" >
							<%   else	%>
                            	<input name="parking" type="text" id="parking" style="width:80px;text-align:right" onKeyUp="plusComma(this);" >
							<% end if	%>
                                </td>
                                <th>통행료</th>
								<td class="left">지불방법
                                <select name="toll_pay" id="toll_pay" style="width:80px">
                                    <option value='현금' <%If toll_pay= "현금" then %>selected<% end if %>>현금</option>
                              	</select>
							<% if u_type = "U" then	%>
                                <input name="toll" type="text" id="toll" style="width:80px;text-align:right" value="<%=formatnumber(toll,0)%>" onKeyUp="plusComma(this);" >
							<%   else	%>
                                <input name="toll" type="text" id="toll" style="width:80px;text-align:right" onKeyUp="plusComma(this);" >
							<% end if	%>
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
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
				<%
					if u_type = "U" and user_id = mg_ce_id then
						if end_yn = "N" or end_yn = "C" then
				%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();" ID="Button1" NAME="Button1"></span>
        		<%
						end if
					end if
				%>
                </div>
				<br>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="old_start_km" value="<%=start_km%>" ID="Hidden1">
                <input type="hidden" name="old_end_km" value="<%=end_km%>" ID="Hidden1">
                <input type="hidden" name="curr_date" value="<%=curr_date%>" ID="Hidden1">
                <input type="hidden" name="end_date" value="<%=end_date%>" ID="Hidden1">
                <input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
				<input type="hidden" name="run_seq" value="<%=run_seq%>" ID="Hidden1">
				<input type="hidden" name="cancel_yn" value="<%=cancel_yn%>" ID="Hidden1">
                <input type="hidden" name="mod_id" value="<%=mod_id%>" ID="Hidden1">
                <input type="hidden" name="mod_user" value="<%=mod_user%>" ID="Hidden1">
                <input type="hidden" name="mod_date" value="<%=mod_date%>" ID="Hidden1">
                <input type="hidden" name="next_km" value="<%=next_km%>" ID="Hidden1">
                <input type="hidden" name="pre_km" value="<%=pre_km%>" ID="Hidden1">
			</form>
		</div>
	</body>
</html>
