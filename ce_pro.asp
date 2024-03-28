<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim com_cnt(18)
dim sum_cnt(18)

from_date=Request.form("from_date")
to_date=Request.form("to_date")

If to_date = "" or from_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
End If

for i = 0 to 18
	com_cnt(i) = 0
	sum_cnt(i) = 0
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = ""
sql = sql & "select mg_ce, mg_ce_id, team from as_acpt "
'//2016-09-21 쿼리 속도 개선
'sql = sql & "Where (Cast(acpt_date as date) >= '" + from_date + "' and Cast(acpt_date as date) <= '"+to_date+"') "
sql = sql & "Where (acpt_date between str_to_date('" & from_date & " 000000','%Y-%m-%d %H%i%s') and str_to_date('" & to_date & " 235959','%Y-%m-%d %H%i%s') ) "
sql = sql & "GROUP BY mg_ce, mg_ce_id, team "
sql = sql & "Order By team, mg_ce, mg_ce_id Asc"
Rs.Open Sql, Dbconn, 1

title_line = "CE별 유형별 처리현황"
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
			function getPageCode(){
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.from_date.value > document.frm.to_date.value) {
					alert ("시작일이 종료일보다 클수가 없습니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=ce_pro.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="*" >
							<col width="5%" >
							<col width="5%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">CE명</th>
								<th scope="col" rowspan="2">소속</th>
								<th scope="col" rowspan="2">총건수</th>
								<th scope="col" rowspan="2">처리율</th>
								<th scope="col" colspan="9" style=" border-bottom:1px solid #e3e3e3;">처 리 완 료</th>
								<th scope="col" colspan="9" style=" border-bottom:1px solid #e3e3e3;">미 처 리</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">소계</th>
								<th scope="col">원격</th>
								<th scope="col">방문</th>
								<th scope="col">신규<br>설치</th>
								<th scope="col">이설<br>설치</th>
								<th scope="col">랜<br>공사</th>
								<th scope="col">회수</th>
								<th scope="col">예방</th>
								<th scope="col">기타</th>
								<th scope="col">소계</th>
								<th scope="col">원격</th>
								<th scope="col">방문</th>
								<th scope="col">신규<br>설치</th>
								<th scope="col">이설<br>설치</th>
								<th scope="col">랜<br>공사</th>
								<th scope="col">회수</th>
								<th scope="col">예방</th>
								<th scope="col">기타</th>
							</tr>
						</thead>
						<tbody>
						<% 
                        do until rs.eof 
                ' 월간 유형별 미처리
							sql = ""
                            sql = sql & "select as_type, as_process, count(*) as end_cnt "
							sql = sql & "from as_acpt "
							sql = sql & "where (mg_ce_id='"+rs("mg_ce_id")+"') "
							'//2016-09-21 쿼리 속도 개선
							'sql = sql & "and (Cast(acpt_date as date) >= '" + from_date + "' and Cast(acpt_date as date) <= '"+to_date+"') "
							sql = sql & "and (acpt_date between str_to_date('" & from_date & " 000000','%Y-%m-%d %H%i%s') and str_to_date('" & to_date & " 235959','%Y-%m-%d %H%i%s') ) "
							sql = sql & "GROUP BY as_type, as_process"

                            rs_as.Open Sql, Dbconn, 1
                            do until rs_as.eof
                                select case rs_as("as_type")
                                    case "원격처리"
										if rs_as("as_process") = "완료" or rs_as("as_process") = "취소" then
											com_cnt(2) = com_cnt(2) + cint(rs_as("end_cnt"))
											com_cnt(1) = com_cnt(1) + cint(rs_as("end_cnt"))
										  else
											com_cnt(11) = com_cnt(11) + cint(rs_as("end_cnt"))
											com_cnt(10) = com_cnt(10) + cint(rs_as("end_cnt"))
										end if
                                    case "방문처리"
										if rs_as("as_process") = "완료" or rs_as("as_process") = "취소" then
											com_cnt(3) = com_cnt(3) + cint(rs_as("end_cnt"))
											com_cnt(1) = com_cnt(1) + cint(rs_as("end_cnt"))
										  else
											com_cnt(12) = com_cnt(12) + cint(rs_as("end_cnt"))
											com_cnt(10) = com_cnt(10) + cint(rs_as("end_cnt"))
										end if
                                    case "신규설치" , "신규설치공사"
										if rs_as("as_process") = "완료" or rs_as("as_process") = "취소" then
											com_cnt(4) = com_cnt(4) + cint(rs_as("end_cnt"))
											com_cnt(1) = com_cnt(1) + cint(rs_as("end_cnt"))
										  else
											com_cnt(13) = com_cnt(13) + cint(rs_as("end_cnt"))
											com_cnt(10) = com_cnt(10) + cint(rs_as("end_cnt"))
										end if
                                    case "이전설치" , "이전설치공사"
										if rs_as("as_process") = "완료" or rs_as("as_process") = "취소" then
											com_cnt(5) = com_cnt(5) + cint(rs_as("end_cnt"))
											com_cnt(1) = com_cnt(1) + cint(rs_as("end_cnt"))
										  else
											com_cnt(14) = com_cnt(14) + cint(rs_as("end_cnt"))
											com_cnt(10) = com_cnt(10) + cint(rs_as("end_cnt"))
										end if
                                    case "랜공사" , "이전랜공사"
										if rs_as("as_process") = "완료" or rs_as("as_process") = "취소" then
											com_cnt(6) = com_cnt(6) + cint(rs_as("end_cnt"))
											com_cnt(1) = com_cnt(1) + cint(rs_as("end_cnt"))
										  else
											com_cnt(15) = com_cnt(15) + cint(rs_as("end_cnt"))
											com_cnt(10) = com_cnt(10) + cint(rs_as("end_cnt"))
										end if
                                    case "장비회수"
										if rs_as("as_process") = "완료" or rs_as("as_process") = "취소" then
											com_cnt(7) = com_cnt(7) + cint(rs_as("end_cnt"))
											com_cnt(1) = com_cnt(1) + cint(rs_as("end_cnt"))
										  else
											com_cnt(16) = com_cnt(16) + cint(rs_as("end_cnt"))
											com_cnt(10) = com_cnt(10) + cint(rs_as("end_cnt"))
										end if
                                    case "예방점검"
										if rs_as("as_process") = "완료" or rs_as("as_process") = "취소" then
											com_cnt(8) = com_cnt(8) + cint(rs_as("end_cnt"))
											com_cnt(1) = com_cnt(1) + cint(rs_as("end_cnt"))
										  else
											com_cnt(17) = com_cnt(17) + cint(rs_as("end_cnt"))
											com_cnt(10) = com_cnt(10) + cint(rs_as("end_cnt"))
										end if
                                    case "기타"
										if rs_as("as_process") = "완료" or rs_as("as_process") = "취소" then
											com_cnt(9) = com_cnt(9) + cint(rs_as("end_cnt"))
											com_cnt(1) = com_cnt(1) + cint(rs_as("end_cnt"))
										  else
											com_cnt(18) = com_cnt(18) + cint(rs_as("end_cnt"))
											com_cnt(10) = com_cnt(10) + cint(rs_as("end_cnt"))
										end if
                                	end select												
                                rs_as.movenext()
                            loop
                            rs_as.close()
                
'							sql = "select * from memb where user_id = '" + rs("mg_ce_id") + "'"
'							Set rs_memb=DbConn.Execute(SQL)
'							if rs_memb.eof or rs_memb.bof then
'								team = "퇴직자"
'							  else
'							  	team = rs_memb("team")
'							end if
							com_cnt(0) = com_cnt(1) + com_cnt(10)
                            if com_cnt(0) <> 0 then
                        %>
							<tr>
                              <td><%=rs("mg_ce")%></td>
                              <td><%=rs("team")%>&nbsp;</td>
                              <td bgcolor="#EEFFFF" class="right"><%=formatnumber(com_cnt(0),0)%></td>
                              <td bgcolor="#EEFFFF" class="right"><%=formatnumber(com_cnt(1)/com_cnt(0)*100,2)%>%</td>
                              <td bgcolor="#FFFFCA" class="right"><%=formatnumber(com_cnt(1),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(2),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(3),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(4),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(5),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(6),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(7),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(8),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(9),0)%></td>
                              <td bgcolor="#FFE8E8" class="right"><%=formatnumber(com_cnt(10),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(11),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(12),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(13),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(14),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(15),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(16),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(17),0)%></td>
                              <td class="right"><%=formatnumber(com_cnt(18),0)%></td>
							</tr>
                		<%
							end if
			
							for i = 0 to 18
								sum_cnt(i) = sum_cnt(i) + com_cnt(i)
								com_cnt(i) = 0
							next

							rs.movenext()
						loop
						rs.close()
						if sum_cnt(0) = 0 then
							sum_pro_per = 0
						  else
							sum_pro_per = sum_cnt(1)/sum_cnt(0)*100
						end if
						%>
							<tr>
                              <th colspan="2">총계</th>
                              <th><%=formatnumber(sum_cnt(0),0)%></th>
                              <th><%=formatnumber(sum_pro_per,2)%>%</th>
                              <th><%=formatnumber(sum_cnt(1),0)%></th>
                              <th><%=formatnumber(sum_cnt(2),0)%></th>
                              <th><%=formatnumber(sum_cnt(3),0)%></th>
                              <th><%=formatnumber(sum_cnt(4),0)%></th>
                              <th><%=formatnumber(sum_cnt(5),0)%></th>
                              <th><%=formatnumber(sum_cnt(6),0)%></th>
                              <th><%=formatnumber(sum_cnt(7),0)%></th>
                              <th><%=formatnumber(sum_cnt(8),0)%></th>
                              <th><%=formatnumber(sum_cnt(9),0)%></th>
                              <th><%=formatnumber(sum_cnt(10),0)%></th>
                              <th><%=formatnumber(sum_cnt(11),0)%></th>
                              <th><%=formatnumber(sum_cnt(12),0)%></th>
                              <th><%=formatnumber(sum_cnt(13),0)%></th>
                              <th><%=formatnumber(sum_cnt(14),0)%></th>
                              <th><%=formatnumber(sum_cnt(15),0)%></th>
                              <th><%=formatnumber(sum_cnt(16),0)%></th>
                              <th><%=formatnumber(sum_cnt(17),0)%></th>
                              <th><%=formatnumber(sum_cnt(18),0)%></th>
							</tr>
 						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

