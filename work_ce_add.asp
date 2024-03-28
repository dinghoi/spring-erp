<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

acpt_no = request("acpt_no")
as_type = request("as_type")
work_man_cnt = request.form("work_man_cnt")
dev_inst_cnt = request.form("dev_inst_cnt")
ran_cnt = request.form("ran_cnt")
alba_cnt = request.form("alba_cnt")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql = "select * from as_acpt where acpt_no = "&int(acpt_no)
Set rs = DbConn.Execute(SQL)
Sql = "select * from memb where user_id = '"&rs("mg_ce_id")&"'"
Set rs_memb = DbConn.Execute(SQL)

if	work_man_cnt = "" then
	work_man_cnt = 1
	if rs("dev_inst_cnt") = "" or isnull(rs("dev_inst_cnt")) then
		dev_inst_cnt = 0
	  else		
		dev_inst_cnt = rs("dev_inst_cnt")
	end if
	ran_cnt = rs("ran_cnt")
	alba_cnt = rs("alba_cnt")
end if

title_line = "작업 인원 추가"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=work_date%>" );
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
			function frmcheck1 () {
				if (chkfrm1()) {
				readonly_send();
				document.frm1.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.work_man_cnt.value > 30 ||document.frm.work_man_cnt.value < 2) {
					alert('작업인원수는 1명 보다 커야 하고 30명 보단 적어야 합니다');
					frm.work_man_cnt.focus();
					return false;}
				if(document.frm.dev_inst_cnt.value == 0 && document.frm.ran_cnt.value == 0) {
					alert('작업수량이 둘다 0 입니다');
					frm.dev_inst_cnt.focus();
					return false;}
				if(document.frm.dev_inst_cnt.value == "-" || document.frm.ran_cnt.value == "-" || document.frm.alba_cnt.value == "-") {
					alert('작업수량 또는 알바인원이 잘못 되어 있습니다.');
					frm.dev_inst_cnt.focus();
					return false;}
				if(document.frm.dev_inst_cnt.value == 0 && document.frm.ran_cnt.value == 0) {
					alert('작업수량이 둘다 0 입니다');
					frm.dev_inst_cnt.focus();
					return false;}
						
				{
				a=confirm('작업인원을 추가하겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function readonly_send() {
				opener.document.frm.dev_inst_cnt.readOnly = true; 
				opener.document.frm.ran_cnt.readOnly = true; 
				opener.document.frm.work_man_cnt.readOnly = true; 
				opener.document.frm.alba_cnt.readOnly = true; 
			}
			function chkfrm1() {

				if(document.frm1.work_man_cnt.value < 2 || document.frm1.work_man_cnt.value > 30) {
					alert('소요 인력이 잘못 되었습니다');
//					frm.work_man_cnt.focus();
					return false;}

				if(document.frm1.work_man_cnt.value > 0) {
					if(document.frm1.mg_ce1.value == "") {
						alert('1번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view1.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 1) {
					if(document.frm1.mg_ce2.value == "") {
						alert('2번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view2.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 2) {
					if(document.frm1.mg_ce3.value == "") {
						alert('3번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view3.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 3) {
					if(document.frm1.mg_ce4.value == "") {
						alert('4번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view4.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 4) {
					if(document.frm1.mg_ce5.value == "") {
						alert('5번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view5.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 5) {
					if(document.frm1.mg_ce6.value == "") {
						alert('6번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view6.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 6) {
					if(document.frm1.mg_ce7.value == "") {
						alert('7번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view7.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 7) {
					if(document.frm1.mg_ce8.value == "") {
						alert('8번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view8.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 8) {
					if(document.frm1.mg_ce9.value == "") {
						alert('9번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view9.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 9) {
					if(document.frm1.mg_ce10.value == "") {
						alert('10번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view10.focus();
						return false;}}
				if(document.frm1.work_man_cnt.value > 10) {
					if(document.frm1.mg_ce11.value == "") {
						alert('11번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view11.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 11) {
					if(document.frm1.mg_ce12.value == "") {
						alert('12번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view12.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value >12) {
					if(document.frm1.mg_ce13.value == "") {
						alert('13번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view13.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 13) {
					if(document.frm1.mg_ce14.value == "") {
						alert('14번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view14.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 14) {
					if(document.frm1.mg_ce15.value == "") {
						alert('15번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view15.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 15) {
					if(document.frm1.mg_ce16.value == "") {
						alert('16번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view16.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 16) {
					if(document.frm1.mg_ce17.value == "") {
						alert('17번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view17.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 17) {
					if(document.frm1.mg_ce18.value == "") {
						alert('18번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view18.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 18) {
					if(document.frm1.mg_ce19.value == "") {
						alert('19번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view19.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 19) {
					if(document.frm1.mg_ce20.value == "") {
						alert('20번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view20.focus();
						return false;}}
				if(document.frm1.work_man_cnt.value > 20) {
					if(document.frm1.mg_ce21.value == "") {
						alert('21번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view21.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 21) {
					if(document.frm1.mg_ce22.value == "") {
						alert('22번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view22.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 22) {
					if(document.frm1.mg_ce23.value == "") {
						alert('23번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view23.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 23) {
					if(document.frm1.mg_ce24.value == "") {
						alert('24번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view24.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 24) {
					if(document.frm1.mg_ce25.value == "") {
						alert('25번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view25.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 25) {
					if(document.frm1.mg_ce26.value == "") {
						alert('26번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view26.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 26) {
					if(document.frm1.mg_ce27.value == "") {
						alert('27번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view27.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 27) {
					if(document.frm1.mg_ce28.value == "") {
						alert('28번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view28.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 28) {
					if(document.frm1.mg_ce29.value == "") {
						alert('29번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view29.focus();
						return false;}}
			
				if(document.frm1.work_man_cnt.value > 29) {
					if(document.frm1.mg_ce30.value == "") {
						alert('30번째 작업자가 지정이 되지 않았습니다');
						frm1.ce_view30.focus();
						return false;}}
			
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body onLoad="menu1()">
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="work_ce_add.asp?acpt_no=<%=acpt_no%>">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>고객명</th>
							  <td class="left">
							  <%=rs("acpt_user")%>&nbsp;<%=rs("user_grade")%>&nbsp;<%=rs("company")%>
                         	  <input name="acpt_no" type="hidden" id="acpt_no" value="<%=rs("acpt_no")%>">
                              </td>
							  <th>회사/부서</th>
							  <td class="left"><%=rs("company")%>&nbsp;<%=rs("dept")%></td>
						    </tr>
							<tr>
							  <th>작업내용</th>
							  <td colspan="3" class="left">
                              설치대수 :&nbsp;
                              <input name="dev_inst_cnt" type="text" id="dev_inst_cnt" onKeyUp="checkNum(this);" value="<%=dev_inst_cnt%>" maxlength="3" style="width:30px;text-align:right">
                              공사대수 :&nbsp;
                              <input name="ran_cnt" type="text" id="ran_cnt" onKeyUp="checkNum(this);" value="<%=ran_cnt%>" maxlength="3" style="width:30px;text-align:right">
                              작업인력 :&nbsp;
                              <input name="work_man_cnt" type="text" id="work_man_cnt" onKeyUp="checkNum(this);" value="<%=work_man_cnt%>" maxlength="2" style="width:30px;text-align:right">
                              알바인원 :&nbsp;
                              <input name="alba_cnt" type="text" id="alba_cnt" onKeyUp="checkNum(this);" value="<%=alba_cnt%>" maxlength="2" style="width:30px;text-align:right">
						      <a href="#" class="btnType03"  onclick="javascript:frmcheck();">작업자추가</a>
                              <br>
                              <br>
                              <strong>설치 및 공사 수량과 작업 인원수를 입력 후 작업자 추가를 하시면 작업자가 증가됩니다.</strong>
                              </td>
						    </tr>
						</tbody>
					</table>
				</form>
			<h3 class="stit">* 작업자 선택</h3>
				<form method="post" name="frm1" action="work_ce_add_save.asp">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="6%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">NO</th>
								<th scope="col">인력검색</th>
								<th scope="col">이름</th>
								<th scope="col">직급</th>
								<th scope="col">사번</th>
								<th scope="col">부서명</th>
							</tr>
						</thead>
						<tbody>
			  				<tr>
								<td class="first">1</td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=1%>','ce_search','scrollbars=yes,width=650,height=400')">조회</a></td>
								<td><input name="mg_ce1" type="text" id="mg_ce1" style="width:60px" readonly="true" value="<%=rs("mg_ce")%>"></td>
								<td><input name="grade1" type="text" id="grade1" style="width:50px" readonly="true" value="<%=rs_memb("user_grade")%>"></td>
								<td><input name="mg_ce_id1" type="text" id="mg_ce_id1" style="width:50px" readonly="true" value="<%=rs("mg_ce_id")%>"></td>
								<td>
									<input name="emp_company1" type="hidden" id="emp_company1" value="<%=rs_memb("emp_company")%>">
                  <input name="bonbu1" type="hidden" id="bonbu1" value="<%=rs_memb("bonbu")%>">
                  <input name="saupbu1" type="hidden" id="saupbu1" value="<%=rs_memb("saupbu")%>">
                  <input name="team1" type="hidden" id="team1" value="<%=rs_memb("team")%>">
                  <input name="reside1" type="hidden" id="reside1" value="<%=rs_memb("reside")%>">
                  <input name="reside_place1" type="hidden" id="reside_place1" value="<%=rs_memb("reside_place")%>">
                  <input name="reside_company1" type="hidden" id="reside1" value="<%=rs_memb("reside_company")%>">
                  <input name="org_name1" type="text" id="org_name1" value="<%=rs_memb("org_name")%>" style="width:200px" readonly="true">
                </td>
							</tr>
						<%
							for i = 2 to work_man_cnt
						%>
			  				<tr>
								<td class="first"><%=i%></td>
								<td><a href="#" class="btnType03" onClick="pop_Window('ce_search.asp?seq=<%=i%>','ce_search','scrollbars=yes,width=650,height=400')">조회</a></td>
								<td><input name="mg_ce<%=i%>" type="text" id="mg_ce<%=i%>" style="width:60px" readonly="true"></td>
								<td><input name="grade<%=i%>" type="text" id="grade<%=i%>" style="width:50px" readonly="true"></td>
								<td><input name="mg_ce_id<%=i%>" type="text" id="mg_ce_id<%=i%>" style="width:50px" readonly="true"></td>
								<td>
									<input name="emp_company<%=i%>" type="hidden" id="emp_company<%=i%>">
                  <input name="bonbu<%=i%>" type="hidden" id="bonbu<%=i%>">
                  <input name="saupbu<%=i%>" type="hidden" id="saupbu<%=i%>">
                  <input name="team<%=i%>" type="hidden" id="team<%=i%>">
                  <input name="org_name<%=i%>" type="text" id="org_name<%=i%>" style="width:200px" readonly="true">
                  <input name="reside<%=i%>" type="hidden" id="reside<%=i%>">
                  <input name="reside_place<%=i%>" type="hidden" id="reside_place<%=i%>">
                  <input name="reside_company<%=i%>" type="hidden" id="reside_company<%=i%>">
                </td>
							</tr>
						<%
							next
						%>
						</tbody>
					</table>                    
                  <input name="acpt_no" type="hidden" id="acpt_no" value="<%=rs("acpt_no")%>">
                    <input name="as_type" type="hidden" id="as_type" value="<%=rs("as_type")%>">
                    <input name="work_man_cnt" type="hidden" id="work_man_cnt" value="<%=work_man_cnt%>">
                    <input name="dev_inst_cnt" type="hidden" id="dev_inst_cnt" value="<%=dev_inst_cnt%>">
                    <input name="ran_cnt" type="hidden" id="ran_cnt" value="<%=ran_cnt%>">
                    <input name="alba_cnt" type="hidden" id="alba_cnt" value="<%=alba_cnt%>">
                    <input name="company" type="hidden" id="company" value="<%=rs("company")%>">
					<br>
                   		<div align=center>
						<% if work_man_cnt > 1 then	%>
                            <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck1();"></span>
						<% end if	%>
                            <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                    	</div>
					<br>
				</form>
                </div>
			</div>
	</body>
</html>

