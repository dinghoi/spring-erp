<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<%
acpt_no = request("acpt_no")
win_sw = request("win_sw")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_mod = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

Sql="select * from large_acpt where acpt_no = "&int(acpt_no)
Set rs=DbConn.Execute(SQL)

as_memo = replace(rs("as_memo"),chr(10),"<br>")

request_date = cstr(rs("request_date")) + " " + mid(cstr(rs("request_time")),1,2) + ":" + mid(cstr(rs("request_time")),3)
if rs("visit_date") = "" or isnull(rs("visit_date")) then
	visit_date = "."
  else
	visit_date = cstr(rs("visit_date")) + " " + mid(cstr(rs("visit_time")),1,2) + ":" + mid(cstr(rs("visit_time")),3)
end if

sql_etc = "select * from memb where user_id = '" + rs("mg_ce_id") +"'"
set rs_etc=dbconn.execute(sql_etc)				
if rs_etc.eof then
	hp_no = "퇴사자"
  else
  	hp_no = rs_etc("hp")
end if
rs_etc.close()

title_line = "대량건 세부내역"

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

	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
    <object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
	</object>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
							<col width="20%" >
							<col width="13%" >
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
							  <td class="left"><%=rs("mg_ce")%>(<%=hp_no%>)</td>
					      	</tr>
							<tr>
							  <th>전화번호</th>
							  <td class="left"><%=rs("tel_ddd")%>-<%=rs("tel_no1")%>-<%=rs("tel_no2")%></td>
							  <th>회사</th>
							  <td class="left"><%=rs("company")%></td>
							  <th>조직명</th>
							  <td class="left"><%=rs("dept")%></td>
					      	</tr>
							<tr>
							  <th>주소</th>
							  <td class="left" colspan="5"><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%>&nbsp;<%=rs("addr")%></td>
					      	</tr>
							<tr>
							  <th>장애내용</th>
							  <td class="left" colspan="5"><%=as_memo%></td>
					      	</tr>
							<tr>
							  <th>요청일</th>
							  <td class="left"><%=request_date%></td>
							  <th>방문일</th>
							  <td class="left"><%=visit_date%></td>
							  <th>처리유형</th>
							  <td class="left"><%=rs("as_type")%></td>
					      	</tr>
							<tr>
							  <th>처리현황</th>
							  <td class="left"><%=rs("as_process")%></td>
							  <th>지연사유</th>
							  <td class="left" colspan="3">&nbsp;<%=rs("into_reason")%></td>
					      	</tr>
							<tr>
							  <th>제조사</th>
							  <td class="left"><%=rs("maker")%></td>
							  <th>장애장비</th>
							  <td class="left"><%=rs("as_device")%></td>
							  <th>모델번호</th>
							  <td class="left">&nbsp;<%=rs("model_no")%></td>
					      	</tr>
							<tr>
							  <th>시리얼NO</th>
							  <td class="left">&nbsp;<%=rs("serial_no")%></td>
							  <th>자산번호</th>
							  <td class="left">&nbsp;<%=rs("asets_no")%></td>
							  <th>사용부품</th>
							  <td class="left">&nbsp;<%=rs("as_parts")%></td>
					      	</tr>
					      	<tr>
							  <th>A/S내역</th>
							  <td class="left" colspan="5">&nbsp;<%=rs("as_history")%></td>
					      	</tr>
							<% 
                            if rs("as_process") = "완료" and ( rs("as_type") = "신규설치" or rs("as_type") = "이전설치" or rs("as_type") = "장비회수" or rs("as_type") = "예방점검" ) then
                                err_name = rs("dev_inst_cnt") + " 대 완료"
                            end if
                            if rs("as_process") = "완료" and rs("as_type") = "랜공사" then
                                err_name = rs("dev_inst_cnt") + " 회선 완료"
                            end if

							path = "/up_image/" + rs("company")
		
							sql_att = "select * from att_file where acpt_no = "&int(acpt_no)
							set rs_att=dbconn.execute(sql_att)
							if rs_att.eof or rs_att.bof then
								not_att = "Y"
							  else
								not_att = "N"
							end if
                            %>
					      	<tr>
							  <th>조치내역</th>
							  <td class="left" colspan="5">&nbsp;<%=err_name%></td>
					      	</tr>
					      	<tr>
							  <th>첨부파일</th>
							  <td class="left" colspan="5">&nbsp;
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
					      	</tr>
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
				</div>
			</div>
	</body>
</html>

