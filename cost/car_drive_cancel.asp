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
Dim u_type, run_date, mg_ce_id, run_seq, title_line, rs, end_view, mg_ce
Dim car_owner, car_no, car_name, oil_kind, start_company, start_point, start_time, start_km
Dim end_company, end_point, end_time, end_km, far, repair_pay, repair_cost, run_memo
Dim oil_amt, oil_pay, oil_price, parking_pay, parking, toll_pay, toll, cancel_yn, end_yn
Dim reg_id, reg_date, reg_user, mod_id, mod_date, mod_user

u_type = Request.QueryString("u_type")
run_date = Request.QueryString("run_date")
mg_ce_id = Request.QueryString("mg_ce_id")
run_seq = Request.QueryString("run_seq")

title_line = "차량 운행일지 지급 등록"

'sql = "select * from transit_cost where run_date ='"&run_date&"' and mg_ce_id ='"&mg_ce_id&"' and run_seq ='"&run_seq&"'"
objBuilder.Append "SELECT car_owner, car_no, car_name, oil_kind, start_company, start_point, start_time, start_km, "
objBuilder.Append "	end_company, end_point, end_time, end_km, far, repair_pay, repair_cost, run_memo, "
objBuilder.Append "	oil_amt, oil_pay, oil_price, parking_pay, parking, toll_pay, toll, cancel_yn, end_yn,"
objBuilder.Append "	trct.reg_id, trct.reg_date, trct.reg_user, trct.mod_id, trct.mod_date, trct.mod_user, "
objBuilder.Append "	memt.user_name "
objBuilder.Append "FROM transit_cost AS trct "
objBuilder.Append "LEFT OUTER JOIN memb AS memt ON trct.mg_ce_id = memt.user_id "
objBuilder.Append "	AND memt.grade < '5' "
objBuilder.Append "WHERE run_date ='"&run_date&"' AND mg_ce_id ='"&mg_ce_id&"' AND run_seq ='"&run_seq&"';"

Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If f_toString(rs("user_name"), "") = "" Then
	mg_ce = "ERROR"
Else
	mg_ce = rs("user_name")
End If

car_owner = rs("car_owner")
car_no = rs("car_no")
car_name = rs("car_name")
oil_kind = rs("oil_kind")
start_company = rs("start_company")
start_point = rs("start_point")
start_time = rs("start_time")
start_km = Int(rs("start_km"))
end_company = rs("end_company")
end_point = rs("end_point")
end_time = rs("end_time")
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

rs.Close() : Set rs = Nothing
DBConn.Close() : Set DBConn = Nothing

If end_yn = "Y" then
	end_view = "마감"
Else
  	end_view = "진행"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
			function goAction(){
			   window.close();
			}

			function goBefore(){
			   history.back();
			}

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				a = confirm('저장 하시겠습니까?');

				if(a == true){
					return true;
				}
				return false;
			}
        </script>
	</head>
	<body>
		<div id="container">
			<h3 class="tit"><%=title_line%></h3>
			<form action="/cost/car_drive_cancel_save.asp" method="post" name="frm">
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableWrite">
					<colgroup>
						<col width="15%" >
						<col width="35%" >
						<col width="15%" >
						<col width="*" >
					</colgroup>
					<tbody>
						<tr>
							<th class="first">운행일</th>
							<td class="left"><%=run_date%></td>
							<th>운행자</th>
							<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)</td>
						</tr>
						<tr>
							<th class="first">차량정보</th>
							<td colspan="3" class="left">
								<strong>소유 :</strong><%=car_owner%>&nbsp;
								<strong>차량번호 :</strong><%=car_no%>&nbsp;
								<strong>차종 :</strong><%=car_name%>&nbsp;
								<strong>유종 :</strong><%=oil_kind%>
							</td>
						</tr>
						<tr>
							<th class="first">출발회사</th>
							<td class="left"><%=start_company%></td>
							<th>출발주소</th>
							<td class="left"><%=start_point%></td>
						</tr>
						<tr>
							<th class="first">출발KM</th>
							<td class="left"><%=FormatNumber(start_km,0)%></td>
							<th>출발시간</th>
							<td class="left"><%=start_time%></td>
						</tr>
						<tr>
							<th class="first">도착회사</th>
							<td class="left"><%=end_company%></td>
							<th>도착주소</th>
							<td class="left"><%=end_point%></td>
						</tr>
						<tr>
							<th class="first">도착KM</th>
							<td class="left"><%=FormatNumber(end_km,0)%></td>
							<th>도착시간</th>
							<td class="left"><%=end_time%></td>
						</tr>
						<tr>
							<th class="first">주행거리</th>
							<td class="left"><%=FormatNumber(far,0)%></td>
							<th>운행목적</th>
							<td class="left"><%=run_memo%></td>
						</tr>
						<tr>
							<th class="first">주유량(L)</th>
							<td class="left"><%=FormatNumber(oil_amt,0)%></td>
							<th>주유금액</th>
							<td class="left"><%=oil_pay%>&nbsp;<%=FormatNumber(oil_price,0)%></td>
						</tr>
						<tr>
							<th class="first">주차비</th>
							<td class="left"><%=parking_pay%>&nbsp;<%=FormatNumber(parking,0)%></td>
							<th>통행료</th>
							<td class="left"><%=toll_pay%>&nbsp;<%=FormatNumber(toll,0)%></td>
						</tr>
    				  <tr>
						<th class="first">취소여부</th>
						<td class="left">
							<input type="radio" name="cancel_yn" value="Y" <%If cancel_yn = "Y" Then %>checked<%End If %> style="width:30px"/>취소
							<input type="radio" name="cancel_yn" value="N" <%If cancel_yn = "N" Then %>checked<%End If %> style="width:30px"/>지급
                        </td>
                        <th>마감여부</th>
						<td class="left"><%=end_view%></td>
					</tr>
					<tr>
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
                </div>
				<input type="hidden" name="mg_ce_id" value="<%=mg_ce_id%>"/>
				<input type="hidden" name="run_date" value="<%=run_date%>"/>
				<input type="hidden" name="run_seq" value="<%=run_seq%>"/>
			</form>
		</div>
	</body>
</html>