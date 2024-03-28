<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim acpt_no, win_sw, title_line
Dim rsAs, as_memo, overtime_view, mg_ce_id, acpt_man, acpt_grade, acpt_user
Dim mg_ce, company, dept, tel_no2

acpt_no = Int(f_Request("acpt_no"))
win_sw = f_Request("win_sw")

title_line = "A/S 세부내역"

'Sql = "select a.* "
'sql = sql & ", (select concat(emp_hp_ddd,'-',emp_hp_no1, '-', emp_hp_no2) from emp_master where emp_no=a.mg_ce_id) AS ce_tel "
'sql = sql & ", (SELECT b.au_name FROM as_unitprice_month b WHERE b.use_yn = 'Y' AND b.au_code = a.as_unit) AS au_name "
'sql = sql & ", CASE a.night WHEN 'Y' THEN 'YES' ELSE 'NO' END AS night_name "
'sql = sql & ", CASE a.weekend_work WHEN 'Y' THEN 'YES' ELSE 'NO' END AS weekend_work_name "
'sql = sql & "from as_acpt a where a.acpt_no = "&int(acpt_no)

objBuilder.Append "SELECT asat.mg_ce_id, asat.acpt_man, asat.acpt_grade, asat.acpt_user, "
objBuilder.Append "	asat.user_grade, asat.mg_ce, asat.tel_ddd, asat.tel_no1, asat.tel_no2, "
objBuilder.Append "	asat.company, asat.dept, asat.sido, asat.gugun, asat.dong, asat.addr, "
objBuilder.Append "	asat.as_type, asat.cowork_yn, asat.as_process, asat.into_reason, "
objBuilder.Append "	asat.as_memo, asat.overtime, asat.request_date, asat.request_time, "
objBuilder.Append "	asat.visit_date, asat.visit_time, asat.visit_request_yn, asat.acpt_date, "
objBuilder.Append "	asat.maker, asat.as_device, asat.model_no, asat.serial_no, "
objBuilder.Append "	asat.asets_no, asat.as_parts, asat.as_history, "
objBuilder.Append "	asat.err_pc_sw, asat.err_pc_hw, asat.err_monitor, asat.err_printer, "
objBuilder.Append "	asat.err_network, asat.err_server, asat.err_adapter, asat.err_etc, "
objBuilder.Append "	asat.dev_inst_cnt, asat.work_man_cnt, asat.ran_cnt, asat.alba_cnt, "
objBuilder.Append "	(SELECT CONCAT(emp_hp_ddd, '-', emp_hp_no1, '-', emp_hp_no2) "
objBuilder.Append "	FROM emp_master "
objBuilder.Append "	WHERE emp_no = asat.mg_ce_id) AS 'ce_tel', "

objBuilder.Append "	(SELECT au_name "
objBuilder.Append "	FROM as_unitprice_month "
objBuilder.Append "	WHERE use_yn = 'Y' "
objBuilder.Append "		AND au_code = asat.as_unit) AS 'au_name', "

objBuilder.Append "	CASE asat.night WHEN 'Y' THEN 'YES' "
objBuilder.Append "	ELSE 'NO' "
objBuilder.Append "	END AS 'night_name', "

objBuilder.Append "	CASE asat.weekend_work WHEN 'Y' THEN 'YES' "
objBuilder.Append "	ELSE 'NO' END AS 'weekend_work_name' "
objBuilder.Append "FROM as_acpt AS asat "
objBuilder.Append "WHERE asat.acpt_no = "&acpt_no

Set rsAs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

mg_ce_id = rsAs("mg_ce_id")
acpt_man = rsAs("acpt_man")
acpt_grade = rsAs("acpt_grade")
acpt_user = rsAs("acpt_user")
user_grade = rsAs("user_grade")
mg_ce = rsAs("mg_ce")
company = rsAs("company")
dept = rsAs("dept")
tel_no2 = rsAs("tel_no2")
sido = rsAs("sido")
gugun = rsAs("gugun")
dong = rsAs("dong")
addr = rsAs("addr")
as_type = rsAs("as_type")
cowork_yn = rsAs("cowork_yn")
as_process = rsAs("as_process")
into_reason = rsAs("into_reason")
as_memo = rsAs("as_memo")
overtime = rsAs("overtime")
request_date = rsAs("request_date")
request_time = rsAs("request_time")
visit_date = rsAs("visit_date")
visit_time = rsAs("visit_time")
visit_request_yn = rsAs("visit_request_yn")
acpt_date = rsAs("acpt_date")
maker = rsAs("maker")
as_device = rsAs("as_device")
serial_no = rsAs("serial_no")
asets_no = rsAs("asets_no")
as_parts = rsAs("as_parts")
as_history = rsAs("as_history")
err_pc_sw = rsAs("err_pc_sw")
err_monitor = rsAs("err_monitor")
err_printer = rsAs("err_printer")
err_network = rsAs("err_network")
err_server = rsAs("err_server")
err_adapter = rsAs("err_adapter")
err_etc = rsAs("err_etc")
dev_inst_cnt = rsAs("dev_inst_cnt")
work_man_cnt = rsAs("work_man_cnt")
ran_cnt = rsAs("ran_cnt")
alba_cnt = rsAs("alba_cnt")
ce_tel = rsAs("ce_tel")
au_name = rsAs("au_name")
weekend_work_name = rsAs("weekend_work_name")

rsAs.Close() : Set rsAs = Nothing

as_memo = Replace(as_memo, Chr(10), "<br>")

If rsAs("overtime") = "Y" Then
	overtime_view = "야특근신청"
Else
  	overtime_view = ""
End If

request_date = CStr(rsAs("request_date")) & " " & Mid(CStr(rsAs("request_time")), 1, 2) & ":" & Mid(CStr(rsAs("request_time")), 3)

If rsAs("visit_date") = "" Or IsNull(rsAs("visit_date")) Then
	visit_date = "."
Else
	visit_date = CStr(rsAs("visit_date")) & " " & Mid(CStr(rsAs("visit_time")), 1, 2) & ":" & Mid(CStr(rsAs("visit_time")), 3)
End If

Dim rsMem, hp_no, visit_request_view

'sql_etc = "select * from memb where user_id = '" + rs("mg_ce_id") +"'"
objBuilder.Append "SELECT hp FROM memb WHERE user_id = '"&rsAs("mg_ce_id")&"' "

Set rsMem = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rsMem.EOF Then
	hp_no = "퇴사자"
Else
  	hp_no = rsMem("hp")
End If
rsMem.Close() : Set rsMem = Nothing

If rsAs("visit_request_yn") = "Y" Then
	visit_request_view = "고객방문요청"
Else
  	visit_request_view = ""
End If

'sql = "select a.* "
'sql = sql & ", (select concat(emp_hp_ddd,'-',emp_hp_no1, '-', emp_hp_no2) from emp_master where emp_no=a.mg_ce_id) AS ce_tel "
'sql = sql & " from as_acpt a"
'sql = sql & where_sql & base_sql & order_sql & " limit "& stpage & "," &pgsize
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
			function goAction(){
		  		 window.close();
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

		<style media="print">
			.noprint { display: none }
		 </style>
	</head>
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
								<td class="left" colspan="3"><%=rsAs("acpt_date")%></td>
					      	</tr>
							<tr>
								<th>접수자</th>
								<td class="left"><%=rsAs("acpt_man")%>&nbsp;<%=rsAs("acpt_grade")%></td>
								<th>사용자</th>
								<td class="left"><%=rsAs("acpt_user")%>&nbsp;<%=rsAs("user_grade")%></td>
								<th>담당CE</th>
								<td class="left"><%=rsAs("mg_ce")%>
								<th>CE TEL</th>
								<td class="left"><%=rsAs("ce_tel")%></td>
   		      				</tr>
							<tr>
								<th>전화번호</th>
								<td class="left"><%=rsAs("tel_ddd")%>-<%=rsAs("tel_no1")%>-<%=rsAs("tel_no2")%></td>
								<th>회사</th>
								<td class="left"><%=rsAs("company")%></td>
								<th>조직명</th>
								<td class="left" colspan="3"><%=rsAs("dept")%></td>
					      	</tr>
							<tr>
								<th>주소</th>
								<td class="left" colspan="7"><%=rsAs("sido")%>&nbsp;<%=rsAs("gugun")%>&nbsp;<%=rsAs("dong")%>&nbsp;<%=rsAs("addr")%></td>
					      	</tr>
							<tr>
								<th>장애내용</th>
								<td class="left" colspan="7"><%=as_memo%></td>
					      	</tr>
							<tr>
								<th>AS표준단가<br>유형</th>
								<td class="left" colspan="2"><%=rsAs("au_name")%></td>
								<th>야간작업</th>
								<td class="left"><%=rsAs("night_name")%></td>
								<th>주말작업</th>
								<td class="left" colspan="2"><%=rsAs("weekend_work_name")%></td>
					      	</tr>
							<tr>
								<th>요청일</th>
								<td class="left"><%=request_date%></td>
								<th>처리일</th>
								<td class="left"><%=visit_date%></td>
								<th>처리유형</th>
								<td class="left"><%=rsAs("as_type")%>&nbsp;<%=visit_request_view%></td>
								<th>협업여부</th>
								<td class="left">
							  	<%If rsAs("cowork_yn") = "Y" Then
                      				Response.Write "협업"
								Else
                      				Response.Write "일반"
								End If %>
								</td>
							</tr>
							<tr>
								<th>처리현황</th>
								<td class="left"><%=rsAs("as_process")%></td>
								<th>지연/입고사유</th>
								<td class="left" colspan="3">&nbsp;<%=rsAs("into_reason")%></td>
								<th> </th>
								<td class="left"> </td>
					      	</tr>
							<tr>
								<th>제조사</th>
								<td class="left"><%=rsAs("maker")%></td>
								<th>장애장비</th>
								<td class="left"><%=rsAs("as_device")%></td>
								<th>모델번호</th>
								<td class="left">&nbsp;<%=rsAs("model_no")%></td>
								<th> </th>
								<td class="left"> </td>
					      	</tr>
							<tr>
								<th>시리얼NO</th>
								<td class="left">&nbsp;<%=rsAs("serial_no")%></td>
								<th>자산번호</th>
								<td class="left">&nbsp;<%=rsAs("asets_no")%></td>
								<th>사용부품</th>
								<td class="left">&nbsp;<%=rsAs("as_parts")%></td>
								<th> </th>
								<td class="left"> </td>
					      	</tr>
					      	<tr>
								<th>작업내역</th>
								<td class="left" colspan="5">&nbsp;<%=rsAs("as_history")%></td>
								<th> </th>
								<td class="left"> </td>
					      	</tr>
							<%
                            Dim error_pro, err_name, err_code, j, i
							Dim rsEtc, etc_name, err_memo

                            error_pro = rsAs("err_pc_sw") & rsAs("err_pc_hw") & rsAs("err_monitor") & rsAs("err_printer") & rsAs("err_network")
							error_pro = error_pro & rsAs("err_server") & rsAs("err_adapter") & rsAs("err_etc")

							If error_pro <> "" Then
								error_pro = Replace(error_pro, ",", "")
								error_pro = Replace(error_pro, " ", "")

								j = Len(error_pro)

								For i = 4 To j Step 4
									err_code = Mid(error_pro, i-3, 4)

									objBuilder.Append "SELECT etc_name FROM etc_code WHERE etc_code = '"&err_code&"' "
									Set rsEtc = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rsEtc.EOF Or rsEtc.BOF Then
										etc_name = ""
									Else
										etc_name = rsEtc("etc_name")

										If err_memo = "" Then
											err_memo = etc_name
										Else
											err_memo = err_memo & "," & etc_name
										End If
									End If
									rsEtc.Close()
								Next
								Set rsEtc = Nothing
							End If

							Dim path, rsFile, not_att, dev_inst_cnt

							path = "/att_file/" & rsAs("company")

							objBuilder.Append "SELECT acpt_no FROM att_file WHERE acpt_no = "&acpt_no
							Set rsFile = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							If rsFile.EOF Or rsFile.BOF Then
								not_att = "Y"
							Else
								not_att = "N"
							End If

							If rsAs("dev_inst_cnt") = "" Or IsNull(rsAs("dev_inst_cnt")) Then
								dev_inst_cnt = "0"
							Else
							  	dev_inst_cnt = rsAs("dev_inst_cnt")
							End If

                            'If rsAs("as_process") = "완료" and ( rsAs("as_type") = "신규설치" rsAs rs("as_type") = "신규설치공사" or rsAs("as_type") = "이전설치" or rsAs("as_type") = "이전설치공사" or rsAs("as_type") = "랜공사" or rsAs("as_type") = "이전랜" ) Then
							If rsAs("as_process") = "완료" Then
								Select Case rsAs("as_type")
									Case "신규설치", "신규설치공사", "이전설치", "이전설치공사", "랜공사", "이전랜"
										err_name = " 설치 : " & CStr(dev_inst_cnt) & "대, 공사: " & CStr(rsAs("ran_cnt")) & "대, "
										err_name = err_name & "작업인원: " & CStr(rsAs("work_man_cnt")) & "명, 알바: " & CStr(rsAs("alba_cnt"))
								End Select
							End If

                            If rsAs("as_process") = "완료" And (rsAs("as_type") = "장비회수" Or rsAs("as_type") = "예방점검") Then
								err_name = "작업: " & CStr(dev_inst_cnt) & "대"
							End If
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
                                If not_att = "N" Then
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

