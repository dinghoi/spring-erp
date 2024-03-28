<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
acpt_no = request("acpt_no")
win_sw = request("win_sw")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_work = Server.CreateObject("ADODB.Recordset")
Set rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_mod = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

Sql = "select a.* "
sql = sql & ", (select concat(emp_hp_ddd,'-',emp_hp_no1, '-', emp_hp_no2) from emp_master where emp_no=a.mg_ce_id) AS ce_tel "
sql = sql & ", (SELECT b.au_name FROM as_unitprice_month b WHERE b.use_yn = 'Y' AND b.au_code = a.as_unit) AS au_name "
sql = sql & ", CASE a.night WHEN 'Y' THEN 'YES' ELSE 'NO' END AS night_name "
sql = sql & ", CASE a.weekend_work WHEN 'Y' THEN 'YES' ELSE 'NO' END AS weekend_work_name "
sql = sql & "from as_acpt a where a.acpt_no = "&int(acpt_no)
'Response.write sql
Set rs=DbConn.Execute(SQL)


as_memo = replace(rs("as_memo"),chr(10),"<br>")





if rs("overtime") = "Y" then
	overtime_view = "야특근신청"
  else
  	overtime_view = ""
end if

request_date = cstr(rs("request_date")) + " " + mid(cstr(rs("request_time")),1,2) + ":" + mid(cstr(rs("request_time")),3)
if rs("visit_date") = "" or isnull(rs("visit_date")) then
	visit_date = "."
  else
	visit_date = cstr(rs("visit_date")) + " " + mid(cstr(rs("visit_time")),1,2) + ":" + mid(cstr(rs("visit_time")),3)
end if




sql_etc = "select * from memb where user_id = '" + rs("mg_ce_id") +"'"
'response.write(sql_etc)
set rs_etc=dbconn.execute(sql_etc)
if rs_etc.eof then
	hp_no = "퇴사자"
  else
  	hp_no = rs_etc("hp")
end if
rs_etc.close()

if rs("visit_request_yn") = "Y" then
	visit_request_view = "고객방문요청"
  else
  	visit_request_view = ""
end if


sql = "select a.* "
sql = sql & ", (select concat(emp_hp_ddd,'-',emp_hp_no1, '-', emp_hp_no2) from emp_master where emp_no=a.mg_ce_id) AS ce_tel "
sql = sql & " from as_acpt a"
sql = sql & where_sql & base_sql & order_sql & " limit "& stpage & "," &pgsize


title_line = "A/S 세부내역"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 세부 내역</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>

		<script type="text/javascript">

			function goAction () {
		  		 window.close () ;
			}

			function printWindow(){
        //		viewOff("button");
                factory.printing.header = ""; //머리말 정의
                factory.printing.footer = ""; //꼬리말 정의
                factory.printing.portrait = false; //출력방향 설정: true - 가로, false - 세로
                factory.printing.leftMargin = 13; //외쪽 여백 설정
                factory.printing.topMargin = 10; //윗쪽 여백 설정
                factory.printing.rightMargin = 13; //오른쯕 여백 설정
                factory.printing.bottomMargin = 15; //바닦 여백 설정
        //		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
        //		factory.printing.printer = ""; //프린터 할 프린터 이름
        //		factory.printing.paperSize = "A4"; //용지선택
        //		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
        //		factory.printing.collate = true; //순서대로 출력하기
        //		factory.printing.copies = "1"; //인쇄할 매수
        //		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
        //		factory.printing.Printer(true); //출력하기
                factory.printing.Preview(); //윈도우를 통해서 출력
                factory.printing.Print(false); //윈도우를 통해서 출력
            }
        </script>

	</head>

	<style media="print">
    .noprint     { display: none }
    </style>

	<body>
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="container">
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="12%" >
							<col width="13%" >
							<col width="13%" >
							<col width="12%" >
							<col width="12%" >
							<col width="14%" >
							<col width="10%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>접수번호</th>
							  <td class="left"><%=acpt_no%></td>
							  <th>접수일자</th>
							  <td class="left" colspan="3"><%=rs("acpt_date")%></td>
					      	</tr>
							<tr>
							  <th>접수자</th>
							  <td class="left"><%=rs("acpt_man")%>&nbsp;<%=rs("acpt_grade")%></td>
							  <th>사용자</th>
							  <td class="left"><%=rs("acpt_user")%>&nbsp;<%=rs("user_grade")%></td>
							  <th>담당CE</th>
							  <td class="left"><%=rs("mg_ce")%>
							  <th>CE TEL</th>
							  <td class="left"><%=rs("ce_tel")%></td>
   		      	</tr>
							<tr>
							  <th>전화번호</th>
							  <td class="left"><%=rs("tel_ddd")%>-<%=rs("tel_no1")%>-<%=rs("tel_no2")%></td>
							  <th>회사</th>
							  <td class="left"><%=rs("company")%></td>
							  <th>조직명</th>
							  <td class="left" colspan="3"><%=rs("dept")%></td>
					      	</tr>
							<tr>
							  <th>주소</th>
							  <td class="left" colspan="7"><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%>&nbsp;<%=rs("addr")%></td>
					      	</tr>
							<tr>
							  <th>장애내용</th>
							  <td class="left" colspan="7"><%=as_memo%></td>
					      	</tr>
							<tr>
							  <th>AS표준단가<br>유형</th>
							  <td class="left" colspan="2"><%=rs("au_name")%></td>
							   <th>야간작업</th>
  							  <td class="left"><%=rs("night_name")%></td>
							  <th>주말작업</th>
							  <td class="left" colspan="2"><%=rs("weekend_work_name")%></td>
					      	</tr>
							<tr>
							  <th>요청일</th>
							  <td class="left"><%=request_date%></td>
							  <th>처리일</th>
							  <td class="left"><%=visit_date%></td>
							  <th>처리유형</th>
							  <td class="left"><%=rs("as_type")%>&nbsp;<%=visit_request_view%></td>
							  <th>협업여부</th>
							  <td class="left">
							  	<% if rs("cowork_yn") = "Y" then	%>
                      	<%="협업"%>
                  <% else	%>
                      	<%="일반"%>
						      <% end if	%>
								</td>
					    </tr>
							<tr>
							  <th>처리현황</th>
							  <td class="left"><%=rs("as_process")%></td>
							  <th>지연/입고사유</th>
							  <td class="left" colspan="3">&nbsp;<%=rs("into_reason")%></td>
							   <th> </th>
							  <td class="left"> </td>
					      	</tr>
							<tr>
							  <th>제조사</th>
							  <td class="left"><%=rs("maker")%></td>
							  <th>장애장비</th>
							  <td class="left"><%=rs("as_device")%></td>
							  <th>모델번호</th>
							  <td class="left">&nbsp;<%=rs("model_no")%></td>
							   <th> </th>
							  <td class="left"> </td>
					      	</tr>
							<tr>
							  <th>시리얼NO</th>
							  <td class="left">&nbsp;<%=rs("serial_no")%></td>
							  <th>자산번호</th>
							  <td class="left">&nbsp;<%=rs("asets_no")%></td>
							  <th>사용부품</th>
							  <td class="left">&nbsp;<%=rs("as_parts")%></td>
							   <th> </th>
							  <td class="left"> </td>
					      	</tr>
					      	<tr>
							  <th>작업내역</th>
							  <td class="left" colspan="5">&nbsp;<%=rs("as_history")%></td>
							  <th> </th>
							  <td class="left"> </td>
					      	</tr>
							<%
                                dim error_pro
                                dim err_name

                                error_pro = rs("err_pc_sw")+rs("err_pc_hw")+rs("err_monitor")+rs("err_printer")+rs("err_network")+rs("err_server")+rs("err_adapter")+rs("err_etc")

                                if error_pro <> "" then
                                    error_pro = replace(error_pro,",","")
                                    error_pro = replace(error_pro," ","")

                                    j = len(error_pro)

                                    for i = 4 to j step 4

                                        err_code = mid(error_pro,i-3,4)

                                        sql_etc = "select * from etc_code where etc_code = '"&err_code&"'"
                                        set rs_etc=dbconn.execute(sql_etc)

										if rs_etc.eof or rs_etc.bof then
											etc_name = ""
										  else
											etc_name = rs_etc("etc_name")
											if err_memo = "" then
												err_memo = etc_name
											  else
												err_memo = err_memo + "," +etc_name
											end if
										end if
                                        rs_etc.close()
                                    next
                                end if

							path = "/att_file/" + rs("company")

							sql_att = "select * from att_file where acpt_no = "&int(acpt_no)
							set rs_att=dbconn.execute(sql_att)
							if rs_att.eof or rs_att.bof then
								not_att = "Y"
							  else
								not_att = "N"
							end if
							if rs("dev_inst_cnt") = "" or isnull(rs("dev_inst_cnt")) then
								dev_inst_cnt = "0"
							  else
							  	dev_inst_cnt = rs("dev_inst_cnt")
							end if
                            if rs("as_process") = "완료" and ( rs("as_type") = "신규설치" or rs("as_type") = "신규설치공사" or rs("as_type") = "이전설치" or rs("as_type") = "이전설치공사" or rs("as_type") = "랜공사" or rs("as_type") = "이전랜" ) then
								err_name = " 설치 : " + cstr(dev_inst_cnt) + "대, 공사: " + cstr(rs("ran_cnt")) + "대, 작업인원: " + cstr(rs("work_man_cnt")) + "명, 알바: " + cstr(rs("alba_cnt"))
							end if
                            if rs("as_process") = "완료" and ( rs("as_type") = "장비회수" or rs("as_type") = "예방점검" ) then
								err_name = "작업: " + cstr(dev_inst_cnt) + "대"
							end if
                            %>
					      	<tr>
							  <th>조치내역</th>
							  <td class="left" colspan="5">&nbsp;<%=err_name%></td>
							  <th> </th>
							  <td class="left"> </td>
					      	</tr>
					      	<tr>
							  <th>첨부파일</th>
							  <td colspan="3" class="left">&nbsp;
								<%
                                if not_att = "N" then
                                    if rs_att("att_file1") <> "" then
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs_att("att_file1")%>">첨부1</a>&nbsp;
                                <%
                                    end if
                                    if rs_att("att_file2") <> "" then
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs_att("att_file2")%>">첨부2</a>&nbsp;
                                <%
                                    end if
                                    if rs_att("att_file3") <> "" then
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs_att("att_file3")%>">첨부3</a>&nbsp;
                                <%
                                    end if
                                    if rs_att("att_file4") <> "" then
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs_att("att_file4")%>">첨부4</a>&nbsp;
                                <%
                                    end if
                                    if rs_att("att_file5") <> "" then
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs_att("att_file5")%>">첨부5</a>&nbsp;
                                <%
                                    end if
                                end if
                                %>
                              </td>
							  <th>야특근등록</th>
							  <td class="left"><%=overtime_view%>&nbsp;</td>
							  <th> </th>
							  <td class="left"> </td>
					      	</tr>
					      	<tr>
					      	  <th>작업인력</th>
					      	  <td colspan="7" class="left">
						<%
							j = 0
							sql_work = "select * from ce_work where acpt_no = "&int(acpt_no)
							'Response.write sql_work & "<br>"
							rs_work.Open sql_work, Dbconn, 1
							do until rs_work.eof
								j = j + 1
								sql_etc = "select * from memb where user_id = '" + rs_work("mg_ce_id") +"'"
								'Response.write sql_etc & "<br>"
								set rs_etc=dbconn.execute(sql_etc)
								if rs_etc.eof then
									work_man = "ERROR"
								  else
									work_man = rs_etc("user_name") + " " + rs_etc("user_grade")
								end if
								rs_etc.close()
						%>
 								<%=j%>.&nbsp;<%=work_man%>(<%=rs_work("mg_ce_id")%>)&nbsp;<%=rs_work("org_name")%>&nbsp;&nbsp;
                        <%
								rs_work.movenext()
							loop
						%>
                              &nbsp;
                              </td>
			      	      </tr>
						</tbody>
					</table>
        <h3 class="stit">* 입고 History</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="13%" >
							<col width="13%" >
							<col width="13%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">진행일자</th>
								<th scope="col">입고처</th>
								<th scope="col">입고진행</th>
								<th scope="col" class="left">입고세부내역</th>
							</tr>
						</thead>
						<tbody>
						<%
                            Sql_in="select * from as_into where acpt_no = " & int(acpt_no) & " order by in_seq asc"
                            Rs_in.Open Sql_in, Dbconn, 1
                            i = 0
                            do until rs_in.eof
                        %>
							<tr>
								<td class="first"><%=rs_in("into_date")%></td>
								<td><%=rs_in("in_place")%></td>
								<td><%=rs_in("in_process")%></td>
								<td style="text-align:left" class="left"><%=rs_in("in_remark")%></td>
							</tr>
						<%
							rs_in.movenext()
						loop
						rs_in.close()
						%>
						</tbody>
					</table>
					<br>
      <h3 class="stit">* 변경 History</h3>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">변경내용</th>
								<th scope="col">변경자</th>
								<th scope="col">변경일자</th>
							</tr>
						</thead>
						<tbody>
						<%
                            Sql_mod="select * from as_mod where acpt_no = " & int(acpt_no) & " order by mod_date asc"
                            ' response.write(Sql_mod)
                            Rs_mod.Open Sql_mod, Dbconn, 1
                            i = 0
                            do until Rs_mod.eof
                        %>
							<tr>
								<td class="first"><%=Rs_mod("mod_pg")%></td>
								<td><%=Rs_mod("mod_name")%>(<%=Rs_mod("mod_id")%>)</td>
								<td><%=Rs_mod("mod_date")%></td>
							</tr>
						<%
							Rs_mod.movenext()
						loop
						Rs_mod.close()
						%>
						</tbody>
					</table>
					<br>
				</form>
     				<div class="noprint">
                   		<div align=center>
                    		<span class="btnType01"><input type="button" value="출력" onclick="javascript:printWindow();"></span>
                    		<span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                    	</div>
    				</div>
				<br>
		  </div>
			</div>
	</body>
</html>

