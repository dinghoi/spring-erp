<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
curr_date = mid(cstr(now()),1,10)

acpt_no = request.form("acpt_no")
company = request.form("company")
dept = request.form("dept")
work_date = request.form("work_date")
week = request.form("week")
work_item = request.form("work_item")
dev_inst_cnt = request.form("dev_inst_cnt")
ran_cnt = request.form("ran_cnt")
work_man_cnt = request.form("work_man_cnt")
from_hh = request.form("from_hh")
from_mm = request.form("from_mm")
to_hh = request.form("to_hh")
to_mm = request.form("to_mm")
work_gubun = request.form("work_gubun")
sign_no = request.form("sign_no")
reg_sw = request.form("reg_sw")
you_yn = request.form("you_yn")

title_line = "서비스 연동 야특근 등록 (2015년 부터)"

work_man = 1
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
					regview();
					document.frm.submit ();
				}
			}			
			function frmcheck1 () {
				document.frm1.submit ();
			}			
			function chkfrm() {
//				if(document.frm.work_date.value < "2015-01-01") {
//					alert('2015년부터 입력이 가능합니다.');
//					frm.work_date.focus();
//					return false;}
				if(document.frm.work_man_cnt.value == "0") {
					alert('작업자 등록이 되어 있지 않습니다 !!!');
					frm.work_man_cnt.focus();
					return false;}			
				if(document.frm.acpt_no.value =="" || document.frm.acpt_no.value =="0") {
					alert('해당 서비스를 선택하셔야 합니다 !!!');
					frm.as_view.focus();
					return false;}			
				if(document.frm.from_hh.value >"23"||document.frm.from_hh.value <"00") {
					alert('시작 시간이 잘못되었습니다');
					frm.from_hh.focus();
					return false;}
				if(document.frm.from_mm.value >"59"||document.frm.from_mm.value <"00") {
					alert('시작 분이 잘못되었습니다');
					frm.from_mm.focus();
					return false;}
				if(document.frm.to_hh.value >"23"||document.frm.to_hh.value <"00") {
					alert('종료 시간이 잘못되었습니다');
					frm.to_hh.focus();
					return false;}
				if(document.frm.to_mm.value >"59"||document.frm.to_mm.value <"00") {
					alert('종료 분이 잘못되었습니다');
					frm.to_mm.focus();
					return false;}			
//				if(document.frm.to_hh.value < document.frm.from_hh.value) {
//					alert('종료시간이 시작시간 보다 빠름니다');
//					frm.to_hh.focus();
//					return false;}
			
//				if(document.frm.from_hh.value == document.frm.to_hh.value) {
//					if(document.frm.to_mm.value <= document.frm.from_mm.value) {
//						alert('종료시간이 시작시간 보다 빠름니다');
//						frm.to_mm.focus();
//						return false;}}
				
				if(document.frm.work_gubun.value =="") {
					alert('야근항목을 선택하세요');
					frm.work_gubun.focus();
					return false;}

				if(document.frm.sign_yn.value == "Y") {
					if(document.frm.sign_no.value =="" ) {
						alert('전자결재NO를 입력하세요');
						frm.sign_no.focus();
						return false;}}							

				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.you_yn[" + j + "].checked")) {
						k = k + 1
					}
				}
				if (k==0) {
					alert ("유무상 구분을 체크하세요");
					return false;
				}	
				if(document.frm.holi_id.value == "휴일") {
					if(document.frm.holi_sw.value == "N" ) {
						alert('휴일수당인데 근무일자는 평일입니다');
						frm.work_gubun.focus();
						return false;}}							
				if(document.frm.holi_id.value == "평일") {
					if(document.frm.holi_sw.value == "Y" ) {
						alert('평일수당과 근무일자는 휴일입니다');
						frm.work_gubun.focus();
						return false;}}							
						
				{
//				a=confirm('작업자를 조회하시겠습니까?')
//				if (a==true) {
					document.frm.reg_sw.value = "Y";
					return true;
//				}
//				return false;
				}
			}
			function regview() {
				document.getElementById('reg_view').style.display = '';
			}
        </script>
	</head>
	<body>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="overtime_as_add_15.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="12%" >
							<col width="38%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>서비스NO</th>
							  <td class="left">
							  <input name="acpt_no" type="text" id="acpt_no" style="width:80px" readonly="true" value="<%=acpt_no%>">
							  <a href="#" class="btnType03" onClick="pop_Window('as_search.asp?work_item=<%=work_item%>','as_search','scrollbars=yes,width=800,height=400')">서비스조회</a>
                              </td>
							  <th>회사/조직명</th>
							  <td class="left">
                              <input name="company" type="text" id="company" style="width:100px" readonly="true" value="<%=company%>">
                              <input name="dept" type="text" id="dept" style="width:150px" readonly="true" value="<%=dept%>">
                              </td>
						    </tr>
							<tr>
							  <th>작업일</th>
							  <td class="left">
                              <input name="work_date" type="text" style="width:80px" readonly="true" value="<%=work_date%>">
                              <input name="week" type="text" style="width:50px" readonly="true" value="<%=week%>">
                              </td>
							  <th>작업내용</th>
							  <td class="left">
                              <input name="work_item" type="text" style="width:100px" readonly="true" value="<%=work_item%>">
                              &nbsp;<strong>* 유무상구분</strong>
                              <input type="radio" name="you_yn" value="Y" <% if you_yn = "Y" then %>checked<% end if %> style="width:20px" id="Radio1">
                              유상
                              <input type="radio" name="you_yn" value="N" <% if you_yn = "N" then %>checked<% end if %> style="width:20px" id="Radio2">
                              무상 </td>
						    </tr>
							<tr>
							  <th>작업수량</th>
							  <td class="left">
                              설치 :
                              <input name="dev_inst_cnt" type="text" id="dev_inst_cnt" size="3" readonly="true" style="text-align:right" value="<%=dev_inst_cnt%>">
							  &nbsp;랜공사 :
							  <input name="ran_cnt" type="text" id="ran_cnt" size="3" readonly="true" style="text-align:right" value="<%=ran_cnt%>">
							  &nbsp;작업인원 :
							  <input name="work_man_cnt" type="text" id="work_man_cnt" size="2" readonly="true" style="text-align:right" value="<%=work_man_cnt%>">
                              </td>
							  <th>작업시간</th>
								<td class="left"><input name="from_hh" type="text" id="from_hh" size="2" maxlength="2" value="<%=from_hh%>">
								  시
                                    <input name="from_mm" type="text" id="from_mm" size="2" maxlength="2" value="<%=from_mm%>">
                                    분 ~
                                    <input name="to_hh" type="text" id="to_hh" size="2" maxlength="2" value="<%=to_hh%>">
                                    시
                                    <input name="to_mm" type="text" id="to_mm" size="2" maxlength="2" value="<%=to_mm%>">
                                분 </td>
						    </tr>
							<tr>
							  <th>야근항목</th>
								<td class="left">
                                <input name="work_gubun" type="text" value="<%=work_gubun%>" readonly="true" style="width:150px">
				          		<a href="#" onClick="pop_Window('overtime_code_search.asp?gubun=<%="AS"%>','overtime_code_pop','scrollbars=yes,width=700,height=400')" class="btnType03">조회</a>
								</td>
							  <th>전자결재NO</th>
								<td class="left"><input name="sign_no" type="text" id="sign_no" style="width:40px" onKeyUp="checkNum(this);" value="<%=sign_no%>" maxlength="4"> *숫자4자리만 입력 가능  <input type="hidden" name="reg_sw" value="<%=reg_sw%>" ID="reg_sw"></td>
						    </tr>
						</tbody>
					</table>
				<h3 class="stit">* 작업자 내역&nbsp;&nbsp;<a href="#" class="btnType03" onClick="javascript:frmcheck();">작업자조회</a></h3>
                    <input type="hidden" name="holi_id" value="<%=holi_id%>" ID="Hidden1">
                    <input type="hidden" name="sign_yn" value="<%=sign_yn%>" ID="Hidden1">
                    <input type="hidden" name="holi_sw" value="<%=holi_sw%>" ID="Hidden1">
				</form>
				<form method="post" name="frm1" action="overtime_as_add_15_save.asp">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="49%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                            <colgroup>
                              <col width="10%" >
                              <col width="20%" >
                              <col width="*" >
                              <col width="30%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">순번</th>
                                <th scope="col">작업자</th>
                                <th scope="col">팀명</th>
                                <th scope="col">소속</th>
                              </tr>
                            </thead>
                            <tbody>
							<%
							if isnull(work_man_cnt) or work_man_cnt = "" then
								record_cnt = 0
								work_man_cnt = 0
							  else
								record_cnt = int((int(work_man_cnt) + 1)/2)
							end if
							sql = "select * from ce_work where work_id = '2' and acpt_no ="&int(acpt_no)&" limit 0," &record_cnt
							Rs.Open Sql, Dbconn, 1
							i = 0
							do until rs.eof
								i = i + 1
								sql = "select * from memb where user_id = '"&rs("mg_ce_id")&"'"
								set rs_memb=dbconn.execute(sql)
								if	rs_memb.eof or rs_memb.bof then
									mg_ce = "ERROR"
								  else
									mg_ce = rs_memb("user_name")
								end if
								rs_memb.close()														
							%>
                                <tr>
                                  <td class="first"><%=i%></td>
                                  <td><%=mg_ce%></td>
                                  <td><%=rs("org_name")%></td>
                                  <td><%=rs("reside_place")%>&nbsp;</td>
                                </tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
                            </tbody>                        
                        </table>
                        </td>
                        <td width="2%"></td>
                        <td width="49%" valign="top">
                        <table cellpadding="0" cellspacing="0" class="tableList">
                            <colgroup>
                              <col width="10%" >
                              <col width="20%" >
                              <col width="*" >
                              <col width="30%" >
                            </colgroup>
                            <thead>
                              <tr>
                                <th class="first" scope="col">순번</th>
                                <th scope="col">작업자</th>
                                <th scope="col">팀명</th>
                                <th scope="col">상주처</th>
                              </tr>
                            </thead>
                            <tbody>
							<%
							sql = "select * from ce_work where work_id = '2' and acpt_no ="&int(acpt_no)&" limit "&record_cnt&"," &work_man_cnt
							Rs.Open Sql, Dbconn, 1
							i = record_cnt
							do until rs.eof
								i = i + 1
								sql = "select * from memb where user_id = '"&rs("mg_ce_id")&"'"
								set rs_memb=dbconn.execute(sql)
								if	rs_memb.eof or rs_memb.bof then
									mg_ce = "ERROR"
								  else
									mg_ce = rs_memb("user_name")
								end if
								rs_memb.close()														
							%>
                                <tr>
                                  <td class="first"><%=i%></td>
                                  <td><%=mg_ce%></td>
                                  <td><%=rs("team")%></td>
                                  <td><%=rs("reside_place")%>&nbsp;</td>
                                </tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
                            </tbody>
                        </table>
                        </td>
                      </tr>
                    </table>
					<br>
                   		<div align=center id="reg_view">
						<% if reg_sw = "Y" and work_man_cnt > 0 then	%>
                            <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck1();"></span>
                      <% end if %>
                            <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                    	</div>
					<input type="hidden" name="work_man_cnt" value="<%=work_man_cnt%>" ID="work_man_cnt">
					<input type="hidden" name="reg_sw" value="<%=reg_sw%>" ID="reg_sw">
					<input type="hidden" name="acpt_no" value="<%=acpt_no%>" ID="acpt_no">
					<input type="hidden" name="company" value="<%=company%>" ID="company">
					<input type="hidden" name="dept" value="<%=dept%>" ID="dept">
					<input type="hidden" name="work_item" value="<%=work_item%>" ID="work_item">
					<input type="hidden" name="from_hh" value="<%=from_hh%>" ID="from_hh">
					<input type="hidden" name="from_mm" value="<%=from_mm%>" ID="from_mm">
					<input type="hidden" name="to_hh" value="<%=to_hh%>" ID="to_hh">
					<input type="hidden" name="to_mm" value="<%=to_mm%>" ID="to_mm">
					<input type="hidden" name="sign_no" value="<%=sign_no%>" ID="sign_no">
					<input type="hidden" name="work_gubun" value="<%=work_gubun%>" ID="work_gubun">
					<input type="hidden" name="you_yn" value="<%=you_yn%>" ID="you_yn">
                </form>
				</div>
			</div>
	</body>
</html>

