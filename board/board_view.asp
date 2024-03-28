<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
'==================================================
Dim board_seq, board_gubun, board_back, be_pg, page
Dim condi, condi_value, ck_sw, condi_sql, title_line
Dim rsStr, sql_p, sql_a, rs_p, rs_a
Dim url, ins_gubun, gubun_name
Dim path, u_type

board_seq = Request("board_seq")
board_gubun = Request("board_gubun")
board_back = Request("board_back")
be_pg = Request("be_pg")
page = Request("page")
condi = Request("condi")
condi_value = Request("condi_value")
ck_sw = Request("ck_sw")

If board_back = "" Then
	board_back = board_gubun
End If

If condi = "" Then
	condi = "all"
End If

ins_gubun = "조회"

If board_gubun = "1" Then
	gubun_name = "사내공지"
ElseIf board_gubun = "2" Then
	gubun_name = "사내게시판"
ElseIf board_gubun = "3" Then
	gubun_name = "A/S공지"
ElseIf board_gubun = "4" Then
	gubun_name = "자료실"
Else
	gubun_name = "전체게시판"
End If

title_line = gubun_name & " " & ins_gubun

'Set Rs = Server.CreateObject("ADODB.Recordset")
'Set rs_gubun = Server.CreateObject("ADODB.Recordset")

'Sql="select * from board where board_seq = '" & board_seq & "'"
objBuilder.Append "select board_seq, reg_name, reg_id, board_title, board_body, reg_date, read_cnt, att_file "
objBuilder.Append "FROM board WHERE board_seq = '" & board_seq & "' "
Set rsStr = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If condi = "all" Then
	condi_sql = ""
Else
  	condi_sql = "AND " & condi & " LIKE '%" & condi_value  & "%' "
End If

'다음글
objBuilder.Append "SELECT board_seq, board_gubun, board_title FROM board "
objBuilder.Append "WHERE board_seq = (SELECT MAX(board_seq) FROM board WHERE board_seq < "
If board_back = "0" Then
	'sql_p = "SELECT board_seq, board_gubun, board_title FROM board where board_seq = ( select max(board_seq) from board where board_seq < "&board_seq&condi_sql&")"
	sql_p = board_seq & condi_sql & ") "
Else
	'sql_p = "SELECT board_seq, board_gubun, board_title FROM board where board_seq = ( select max(board_seq) from board where board_seq < "&board_seq& " and board_gubun = '"& board_gubun& "' "&condi_sql&")"
	sql_p = board_seq & " AND board_gubun = '" & board_gubun & "' " & condi_sql & ") "
End If
objBuilder.Append sql_p

Set rs_p = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'이전글
objBuilder.Append "SELECT board_seq, board_gubun, board_title FROM board "
objBuilder.Append "WHERE board_seq = (SELECT MIN(board_seq) FROM board WHERE board_seq > "
If board_back = "0" Then
	'sql_a = "SELECT board_seq, board_gubun, board_title FROM board where board_seq = ( select min(board_seq) from board where board_seq > "&board_seq&condi_sql&")"
	sql_a = board_seq & condi_sql & ") "
Else
	'sql_a = "SELECT board_seq, board_gubun, board_title FROM board where board_seq = ( select min(board_seq) from board where board_seq > "&board_seq& " and board_gubun = '"& board_gubun& "' "&condi_sql&")"
	sql_a = board_seq & " AND board_gubun = '" & board_gubun & "' " & condi_sql & ") "
End If
objBuilder.Append sql_a

Set rs_a = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "update board  set read_cnt = read_cnt + 1  where board_seq ="&board_seq
objBuilder.Append "UPDATE board SET read_cnt = read_cnt + 1  WHERE board_seq =" & board_seq
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

url = "/main/nkp_main.asp?board_gubun="&board_back&"&page="&page&"&condi="&condi&"&condi_value="&condi_value&"&ck_sw=y"
%>
<!DOCTYPE HTML>
<html lang="ko">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>NKP 시스템</title>
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

			function goBefore(){
					//history.back() ;
				location.replace("<%=url%>");
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.pass.value == ""){
					alert("비밀번호를 입력하세요.");
					return false;
				}
				return true;
			}
		</script>
	</head>
	<body>
		<div id="wrap">
        	<!--#include virtual = "/include/main_header.asp" -->
			<!--#include virtual = "/include/main_menu.asp" -->
			<div id="container">

				<h3 class="tit"><%=title_line%></h3>
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
							  	<td class="left"><%=rsStr("reg_name")%>&nbsp;(&nbsp;<%=rsStr("reg_id")%>&nbsp;)</td>
							</tr>
							<tr>
							  	<th scope="row">제목</th>
						    	<td class="left"><%=rsStr("board_title")%></td>
							</tr>
							<tr>
								<th scope="row">내용</th>
								<td class="left"><%=rsStr("board_body")%></td>
							</tr>
							<tr>
							  	<th scope="row">작성일</th>
							  	<td class="left"><%=rsStr("reg_date")%></td>
					      </tr>
							<tr>
							  	<th scope="row">조회수</th>
							  	<td class="left"><%=rsStr("read_cnt")%></td>
					      </tr>
							<tr>
							  	<th scope="row">첨부파일</th>
							  	<td class="left">
								<%
                                If rsStr("att_file") <> "" Then
                                    path = "/nkp_upload"
                                %>
                                  <a href="/att_file_download.asp?path=<%=path%>&att_file=<%=rsStr("att_file")%>"><%=rsStr("att_file")%></a>
                                  <% Else %>
				                    &nbsp;
                                <% End If %>
							  	</td>
					      </tr>
							<tr>
								<th scope="row">이전글</th>
								<td class="left">
								<img src="/image/up_16x16.png" width="16" height="16">
								<% if not rs_a.eof Then %>
									<a href="/board/board_view.asp?board_seq=<%=rs_a("board_seq")%>&board_gubun=<%=rs_a("board_gubun")%>&page=<%=page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>"><%=rs_a("board_title")%></a>
								<% Else %>
									처음 자료입니다.
								<% End If %>
                                </td>
							</tr>
							<tr>
								<th scope="row">다음글</th>
								<td class="left">
								<img src="/image/down_16x16.png" width="16" height="16">
								<% if not rs_p.eof Then %>
									<a href="/board/board_view.asp?board_seq=<%=rs_p("board_seq")%>&board_gubun=<%=rs_p("board_gubun")%>&page=<%=page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>"><%=rs_p("board_title")%></a>
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
	<form method="post" name="frm" action="/board/board_del.asp?page=<%=page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>">
	<div align="center">
		<p>
		<span class="btnType04"><input type="button" value="목록" onclick="javascript:goBefore();"></span>
		<span class="btnType04"><input type="button" value="수정" onclick="pop_Window('/board/board_write.asp?board_seq=<%=rsStr("board_seq")%>&page=<%=page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>&u_type=<%="U"%>','board_write_popup','scrollbars=yes,width=1250,height=600')"></span>
        &nbsp;삭제시 비밀번호를 입력하세요 &nbsp;<input name="pass" type="text" id="pass" title="패스워드" maxlength="4" notnull errname="패스워드" style="width:80px;"/>
        <span class="btnType04"><input type="button" value="삭제" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
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

