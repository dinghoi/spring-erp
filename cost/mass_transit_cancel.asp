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
Dim run_date, mg_ce_id, run_seq, start_point, start_time, company
Dim end_point, end_time, transit, payment, fare, run_memo, cancel_yn
Dim end_yn, reg_id, reg_user, reg_date, mod_id, mod_user, mod_date
Dim mem_name, mg_ce, rsTran, title_line, end_view

run_date = Request.QueryString("run_date")
mg_ce_id = Request.QueryString("mg_ce_id")
run_seq = Int(Request.QueryString("run_seq"))

objBuilder.Append "SELECT start_point, start_time, company, end_point, end_time, "
objBuilder.Append "	transit, payment, fare, run_memo, cancel_yn, end_yn, trct.reg_id, "
objBuilder.Append "	trct.reg_user, trct.reg_date, trct.mod_id, trct.mod_user, trct.mod_date, "
objBuilder.Append "	memt.user_name "
objBuilder.Append "FROM transit_cost AS trct "
objBuilder.Append "LEFT OUTER JOIN memb AS memt ON trct.mg_ce_id = memt.user_id "
objBuilder.Append "	AND memt.grade < '5' "
objBuilder.Append "WHERE run_date ='"&run_date&"' and mg_ce_id ='"&mg_ce_id&"' and run_seq ="&run_seq

Set rsTran = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

start_point = rsTran("start_point")
start_time = rsTran("start_time")
company = rsTran("company")
end_point = rsTran("end_point")
end_time = rsTran("end_time")
transit = rsTran("transit")
payment = rsTran("payment")
fare = int(rsTran("fare"))
run_memo = rsTran("run_memo")
cancel_yn = rsTran("cancel_yn")
end_yn = rsTran("end_yn")
reg_id = rsTran("reg_id")
reg_user = rsTran("reg_user")
reg_date = rsTran("reg_date")
mod_id = rsTran("mod_id")
mod_user = rsTran("mod_user")
mod_date = rsTran("mod_date")
mem_name = rsTran("user_name")

If f_toString(mem_name, "") = "" Then
	mg_ce = "ERROR"
Else
	mg_ce = mem_name
End If

rsTran.Close() : Set rsTran = Nothing
DBConn.Close() : Set DBConn = Nothing

title_line = "대중 교통비 지급 변경"

If end_yn = "Y" Then
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
				a=confirm('저장 하시겠습니까?');

				if(a == true){
					return true;
				}
				return false;
			}

 			function week_check(){
				a = document.frm.run_date.value.substring(0,4);
				b = document.frm.run_date.value.substring(5,7);
				c = document.frm.run_date.value.substring(8,10);

				var newDate = new Date(a,b-1,c);
				var s = newDate.getDay();

				switch(s){
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
       </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/cost/mass_transit_cancel_save.asp" method="post" name="frm">
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
								<th class="first">이용일</th>
								<td class="left"><%=run_date%></td>
								<th>이용자</th>
								<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)</td>
							</tr>
							<tr>
								<th class="first">업체</th>
								<td class="left"><%=company%></td>
								<th>출발지</th>
								<td class="left"><%=start_point%></td>
							</tr>
							<tr>
								<th class="first">출발시간</th>
								<td class="left"><%=start_time%></td>
								<th>도착지</th>
								<td class="left"><%=end_point%></td>
							</tr>
							<tr>
								<th class="first">도착시간</th>
								<td class="left"><%=end_time%></td>
								<th>교통편</th>
								<td class="left"><%=transit%></td>
							</tr>
							<tr>
								<th class="first">교통비</th>
								<td class="left"><%=payment%>&nbsp;<%=FormatNumber(fare,0)%></td>
								<th>작업내용</th>
								<td class="left"><%=run_memo%></td>
							</tr>
    				  <tr>
						<th class="first">취소여부</th>
						<td class="left">
						<input type="radio" name="cancel_yn" value="Y" <% if cancel_yn = "Y" then %>checked<% end if %> style="width:40px"/>취소
                        <input type="radio" name="cancel_yn" value="N" <% if cancel_yn = "N" then %>checked<% end if %> style="width:40px"/>지급
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