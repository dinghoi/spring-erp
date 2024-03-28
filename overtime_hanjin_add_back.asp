<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
dim holi_tab(10)
u_type = request("u_type")

mg_ce_id = user_id
mg_ce = user_name
apct_no = 0
company = ""
dept = ""
work_item = ""
from_hh = ""
from_mm = ""
to_hh = ""
to_mm = ""
work_gubun = ""
overtime_amt = 0
work_memo = ""
sign_no = ""
cancel_yn = "N"
end_yn = "N"
reg_id = user_id
reg_date = now()

curr_date = mid(cstr(now()),1,10)
work_date = mid(cstr(now()),1,10)
company = reside_company

be_date = cstr(dateadd("m", -2,curr_date))

sql = "select * from holiday where holiday >= '" + be_date  + "' and holiday <= '" + curr_date  + "'"
Rs.Open Sql, Dbconn, 1
i = 0
do until rs.eof
	i = i + 1
	holi_tab(i) = rs("holiday")
	rs.movenext()
loop
holi_seq = i
rs.close()

title_line = "한진 당직 및 스케쥴 등록"
if u_type = "U" then

	work_date = request("work_date")
	mg_ce_id = request("mg_ce_id")

	sql = "select * from overtime where work_date = '" + work_date + "' and mg_ce_id = '" + mg_ce_id + "'"
	set rs = dbconn.execute(sql)

	sql="select * from memb where user_id = '" + rs("mg_ce_id") + "'"
	set rs_memb=dbconn.execute(sql)

	if	rs_memb.eof or rs_memb.bof then
		mg_ce = "ERROR"
	  else
		mg_ce = rs_memb("user_name")
	end if
	rs_memb.close()						

	if isnull(rs("acpt_no")) then
		acpt_no = 0
	  else
		acpt_no = rs("acpt_no")
	end if
	mg_ce_id = rs("mg_ce_id")
	company = rs("company")
	dept = rs("dept")
	work_item = rs("work_item")
	from_hh = mid(rs("from_time"),1,2)
	from_mm = mid(rs("from_time"),3,2)
	to_hh = mid(rs("to_time"),1,2)
	to_mm = mid(rs("to_time"),3,2)
	work_gubun = rs("work_gubun")
	overtime_amt = int(rs("overtime_amt"))
	work_memo = rs("work_memo")
	sign_no = rs("sign_no")
	cancel_yn = rs("cancel_yn")
	you_yn = rs("you_yn")
	reg_id = rs("reg_id")
	reg_user = rs("reg_user")
	reg_date = rs("reg_date")
	mod_id = rs("mod_id")
	mod_user = rs("mod_user")
	mod_date = rs("mod_date")
	rs.close()

	title_line = "한진 당직 및 스케쥴 수정"
end if
if end_yn = "Y" then
	end_view = "마감"
  else
  	end_view = "진행"
end if

strNowWeek = WeekDay(work_date)
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
if week = "토요일" or week = "일요일" then
	holi_sw = "Y"
  else	
  	holi_sw = "N"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
			function chkfrm() {
				var holi_sw;
				var holi_ary = new Array();
				holi_ary[1] = document.frm.holi_tab1.value;
				holi_ary[2] = document.frm.holi_tab2.value;
				holi_ary[3] = document.frm.holi_tab3.value;
				holi_ary[4] = document.frm.holi_tab4.value;
				holi_ary[5] = document.frm.holi_tab5.value;
				holi_ary[6] = document.frm.holi_tab6.value;
				holi_ary[7] = document.frm.holi_tab7.value;
				holi_ary[8] = document.frm.holi_tab8.value;
				holi_ary[9] = document.frm.holi_tab9.value;
				holi_ary[10] = document.frm.holi_tab10.value;
				holi_sw = "N";
				for (i=1;i<11;i++) {
					if (document.frm.work_date.value == holi_ary[i]) {
						holi_sw = "Y";
					}				
				}
				if(document.frm.week.value == "토요일" || document.frm.week.value == "일요일" ) {
					holi_sw = "Y";
				}
				document.frm.holi_sw.value = holi_sw;
				if(document.frm.work_date.value < "2014-12-01") {
					alert('2014년12월부터 입력이 가능합니다.');
					frm.work_date.focus();
					return false;}
				if(document.frm.company.value =="") {
					alert('회사명을 선택하세요');
					frm.company.focus();
					return false;}
				if(document.frm.dept.value =="") {
					alert('부서명을 입력하세요');
					frm.dept.focus();
					return false;}
				if(document.frm.work_gubun.value =="") {
					alert('수당구분을 선택하세요');
					frm.work_gubun.focus();
					return false;}
				if(document.frm.sign_yn.value == "Y") {
					if(document.frm.from_hh.value >"23"||document.frm.from_hh.value <"00") {
						alert('시작 시간이 잘못되었습니다');
						frm.from_hh.focus();
						return false;}}
				if(document.frm.sign_yn.value == "Y") {
					if(document.frm.from_mm.value >"59"||document.frm.from_mm.value <"00") {
						alert('시작 분이 잘못되었습니다');
						frm.from_mm.focus();
						return false;}}
				if(document.frm.sign_yn.value == "Y") {
					if(document.frm.to_hh.value >"23"||document.frm.to_hh.value <"00") {
						alert('종료 시간이 잘못되었습니다');
						frm.to_hh.focus();
						return false;}}
				if(document.frm.sign_yn.value == "Y") {
					if(document.frm.to_mm.value >"59"||document.frm.to_mm.value <"00") {
						alert('종료 분이 잘못되었습니다');
						frm.to_mm.focus();
						return false;}}
			
//				if(document.frm.to_hh.value < document.frm.from_hh.value) {
//					alert('종료시간이 시작시간 보다 빠름니다');
//					frm.to_hh.focus();
//					return false;}
			
//				if(document.frm.from_hh.value == document.frm.to_hh.value) {
//					if(document.frm.to_mm.value <= document.frm.from_mm.value) {
//						alert('종료시간이 시작시간 보다 빠름니다');
//						frm.to_mm.focus();
//						return false;}}							
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
					alert ("유무상 구분을 선택하세요");
					return false;
				}	
				if(document.frm.work_memo.value =="" ) {
					alert('작업내용을 입력하세요');
					frm.work_memo.focus();
					return false;}
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
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function week_check() {
			
			a = document.frm.work_date.value.substring(0,4);
			b = document.frm.work_date.value.substring(5,7);
			c = document.frm.work_date.value.substring(8,10);
			
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
				holi_sw = "N"
				if(str == "토요일" || str == "일요일" ) {
					holi_sw = "Y";
				}
				document.frm.holi_sw.value = holi_sw;
			}
			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U') 
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
			function delcheck() 
				{
				a=confirm('정말 삭제하시겠습니까?')
				if (a==true) {
					document.frm.action = "overtime_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="overtime_hanjin_add_save.asp" method="post" name="frm">
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
								<th class="first">작업일</th>
								<td class="left">
                                <input name="work_date" type="text" id="datepicker" style="width:70px;text-align:center" value="<%=work_date%>" readonly="true" onChange="week_check();">
                                <input name="week" type="text" id="week" style="width:40px;text-align:center" value="<%=week%>" readonly="true">
                                &nbsp;마감일자 : <%=end_date%>
                                </td>
								<th>작업자</th>
								<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)
                                <input name="curr_date" type="hidden" id="curr_date" value="<%=now()%>">
                                <input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=mg_ce_id%>">
                                </td>
							</tr>
							<tr>
								<th class="first">회사명</th>
								<td class="left"><input name="company" type="text" value="<%=company%>" readonly="true" style="width:150px">
				          <a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="4"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조회</a></td>
								<th>부서명</th>
								<td class="left"><input name="dept" type="text" id="dept" onKeyUp="checklength(this,50)"  style="width:200px" value="<%=dept%>"></td>
							</tr>
							<tr>
								<th class="first">수당구분</th>
								<td class="left">
                                <input name="work_gubun" type="text" value="<%=work_gubun%>" readonly="true" style="width:150px">
				          		<a href="#" onClick="pop_Window('overtime_code_search.asp?gubun=<%="한진"%>','overtime_code_pop','scrollbars=yes,width=700,height=400')" class="btnType03">조회</a>
                                </td>
								<th>작업시간</th>
								<td class="left">
                                <input name="from_hh" type="text" id="from_hh" size="2" maxlength="2" value="<%=from_hh%>">시
								<input name="from_mm" type="text" id="from_mm" size="2" maxlength="2" value="<%=from_mm%>">분 ~
								<input name="to_hh" type="text" id="to_hh" size="2" maxlength="2" value="<%=to_hh%>">시
								<input name="to_mm" type="text" id="to_mm" size="2" maxlength="2" value="<%=to_mm%>">분
                                </td>
							</tr>
							<tr>
								<th class="first">전자결재NO</th>
								<td class="left">
                                <input name="sign_no" type="text" id="sign_no" style="width:40px" onKeyUp="checkNum(this);" value="<%=sign_no%>" maxlength="4">&nbsp;*숫자4자리만 입력 가능
  								<input type="hidden" name="reg_sw" value="<%=reg_sw%>" ID="reg_sw">
  								</td>
								<th><span class="first">유무상구분</span></th>
								<td class="left"><input type="radio" name="you_yn" value="N" <% if you_yn = "N" then %>checked<% end if %> style="width:40px" id="Radio4">
								  무상
                                    <input type="radio" name="you_yn" value="Y" <% if you_yn = "Y" then %>checked<% end if %> style="width:40px" id="Radio3">
                                유상 </td>
							</tr>
							<tr>
							  <th class="first">작업내용</th>
							  <td colspan="3" class="left"><input name="work_memo" type="text" id="work_memo" onKeyUp="checklength(this,50)"  style="width:300px" value="<%=work_memo%>"></td>
						  </tr>
							<tr id="cancel_col" style="display:none">
								<th class="first">취소여부</th>
								<td class="left">
								<input type="radio" name="cancel_yn" value="Y" <% if cancel_yn = "Y" then %>checked<% end if %> style="width:40px" ID="Radio1">취소
				                <input type="radio" name="cancel_yn" value="N" <% if cancel_yn = "N" then %>checked<% end if %> style="width:40px" ID="Radio2">지급
								</td>
                                <th>마감여부</th>
								<td class="left"><%=end_view%><input name="end_yn" type="hidden" id="end_yn" value="<%=end_yn%>"></td>
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
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="acpt_no" value="<%=acpt_no%>" ID="Hidden1">
				<input type="hidden" name="holi_id" value="<%=holi_id%>" ID="Hidden1">
				<input type="hidden" name="sign_yn" value="<%=sign_yn%>" ID="Hidden1">
				<% for i = 1 to 10	%>
				<input type="hidden" name="holi_tab<%=i%>" value="<%=holi_tab(i)%>" ID="Hidden1">
                <% next	%>
				<input type="hidden" name="holi_sw" value="<%=holi_sw%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

