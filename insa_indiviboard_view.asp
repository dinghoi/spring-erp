<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
board_seq = Request("board_seq")
board_gubun = Request("board_gubun")
board_back = Request("board_back")
be_pg = Request("be_pg")
page = request("page")
condi = request("condi")
condi_value = request("condi_value")
ck_sw = Request("ck_sw")

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_no = request.cookies("nkpmg_user")("coo_user_id")

if condi = "" then
	condi = "all"
end if

ins_gubun = "조회"

if board_gubun = "1" then
	gubun_name = "인사공지"
  elseif board_gubun = "2" then
  	gubun_name = "인사게시판"
  elseif board_gubun = "3" then
  	gubun_name = "급여공지"
  elseif board_gubun = "4" then
  	gubun_name = "자료실"
  else
  	gubun_name = "전체게시판"  
end if
title_line = gubun_name + " " + ins_gubun

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_gubun = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql="select * from emp_board where board_seq = '" + board_seq + "'"
'Sql="select * from board where board_seq = '" + board_seq + "'"
Set rs=DbConn.Execute(Sql)

if condi = "all" then
	condi_sql = " "
  else
  	condi_sql = " and " + condi + " like '%" + condi_value  + "%'"
end if

if board_back = "0" then
	sql_p = "SELECT board_seq, board_gubun, board_title FROM emp_board where board_seq = ( select max(board_seq) from emp_board where board_seq < "&board_seq&condi_sql&")"
  else
	sql_p = "SELECT board_seq, board_gubun, board_title FROM emp_board where board_seq = ( select max(board_seq) from emp_board where board_seq < "&board_seq& " and board_gubun = '"& board_gubun& "' "&condi_sql&")"
end if
set rs_p = Dbconn.execute(sql_p)

if board_back = "0" then
	sql_a = "SELECT board_seq, board_gubun, board_title FROM emp_board where board_seq = ( select min(board_seq) from emp_board where board_seq > "&board_seq&condi_sql&")"
  else
	sql_a = "SELECT board_seq, board_gubun, board_title FROM emp_board where board_seq = ( select min(board_seq) from emp_board where board_seq > "&board_seq& " and board_gubun = '"& board_gubun& "' "&condi_sql&")"
end if
set rs_a = Dbconn.execute(sql_a)

sql = "update emp_board  set read_cnt = read_cnt + 1  where board_seq ="&board_seq
dbconn.execute(sql)
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.pass.value == "") {
					alert ("비밀번호를 입력하세요.");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<tbody>
							<tr>
							  <th scope="row" width="15%">위치</th>
								<td class="left"><%=gubun_name%>
								  <label>
								    <input name="board_gubun" type="hidden" id="board_gubun" value="<%=board_gubun%>">
						        </label></td>
							</tr>
							<tr>
							  	<th scope="row" width="15%">작성자</th>
							  	<td class="left"><%=rs("reg_name")%>&nbsp;(&nbsp;<%=rs("reg_id")%>&nbsp;)</td>
							</tr>
							<tr>
							  	<th scope="row">제목</th>
						    	<td class="left"><%=rs("board_title")%></td>
							</tr>
							<tr>
								<th scope="row">내용</th>
								<td class="left"><%=rs("board_body")%></td>
							</tr>
							<tr>
							  	<th scope="row">작성일</th>
							  	<td class="left"><%=rs("reg_date")%></td>
					      </tr>
							<tr>
							  	<th scope="row">조회수</th>
							  	<td class="left"><%=rs("read_cnt")%></td>
					      </tr>
							<tr>
							  	<th scope="row">첨부파일</th>
							  	<td class="left">
								<% 
                                If rs("att_file") <> "" Then 
                                    path = "/insanotice_upload" 
                                %>
                                  <a href="att_file_download.asp?path=<%=path%>&att_file=<%=rs("att_file")%>"><%=rs("att_file")%></a>
                                  <% Else %>
				                    &nbsp;
                                <% End If %>
							  	</td>
					      </tr>
							<tr>
								<th scope="row">이전글</th>
								<td class="left">
								<img src="image/up_16x16.png" width="16" height="16">
								<% if not rs_a.eof Then %> 
									<a href="insa_indiviboard_view.asp?board_seq=<%=rs_a("board_seq")%>&board_gubun=<%=rs_a("board_gubun")%>&page=<%=page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>"><%=rs_a("board_title")%></a>
								<% Else %>
									처음 자료입니다.
								<% End If %>
                                </td>
							</tr>
							<tr>
								<th scope="row">다음글</th>
								<td class="left">
								<img src="image/down_16x16.png" width="16" height="16">
								<% if not rs_p.eof Then %> 
									<a href="insa_indiviboard_view.asp?board_seq=<%=rs_p("board_seq")%>&board_gubun=<%=rs_p("board_gubun")%>&page=<%=page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>"><%=rs_p("board_title")%></a>
								<% Else %>
									처음 자료입니다.
								<% End If %>
                                </td>
							</tr>
						</tbody>
					</table>
				</div>
			</div>				
	</div>        				
	<form method="post" name="frm" action="insa_board_del.asp?page=<%=page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>">
	<div align=center>
		<p>
		<% if be_pg = "C" then	%>
		<a href="communication.asp">
		<%	else	%>        
		<a href="insa_individual_board.asp?board_gubun=<%=board_back%>&page=<%=page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>" class="btnType04">
		<%	end if	%>
		목록</a> &nbsp;
		</p>
		<br>
        <br>
        <br>
	</div>
	<p>
	  <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
	  <input type="hidden" name="board_seq" value="<%=board_seq%>" ID="Hidden1">
    </p>
	</form>
	</body>
</html>

