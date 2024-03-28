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
'===================================================
Dim board_gubun, first_sw, title_line, page
Dim pgsize, start_page, stpage
Dim total_record, total_page, new_date
Dim strSql, rsBoard, arrBoard, condi, condi_value
Dim be_pg, url_param

board_gubun = f_Request("board_gubun")
first_sw = f_Request("first_sw")

condi = f_Request("condi")
condi_value = f_Request("condi_value")

page = f_Request("page")

'===================================================
'### paging setting
'===================================================
pgsize = 10 ' 화면 한 페이지
be_pg = "/main/nkp_main.asp"

If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
url_param = "&board_gubun="&board_gubun&"&condi="&condi&"&condi_value="&condi_value

If board_gubun = "" Then
	board_gubun = "0"
End If

If board_gubun = "1" Then
	title_line = "사내공지"
ElseIf board_gubun = "3" Then
  	title_line = "A/S공지"
ElseIf board_gubun = "4" Then
  	title_line = "자료실"
Else
  	title_line = "전체게시판"
End If

new_date = Now() - 14

strSql = "CALL USP_MAIN_BOARD_LIST('"&board_gubun&"', '"&condi&"', '"&condi_value&"', "&stpage&", "&pgsize&")"
Set rsBoard = DBConn.Execute(strSql)

If Not rsBoard.EOF Then
	arrBoard = rsBoard.getRows()
	total_record = CInt(arrBoard(0, 0))
Else
	total_record = 0
End If

Call Rs_Close(rsBoard)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
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
		<!--<script type="text/javascript" src="/java/js_window.js"></script>-->

		<script type="text/javascript">
			$(document).ready(function(){
				pop_center('640','700',0,0,0,0,0,0,0);
			});

			function frmcheck(){
				//chkfrm 함수 호출 시 condi 값이 공백일 수 없음(주석 처리)[허정호_20210707]
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
				return;
			}

			function chkfrm(){
				if(document.frm.condi.value === ""){
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}
				return true;
			}

			function pop_center(w,h,tb,st,di,mb,sb,re){
				var mobilecheck = function(){
					var check = false;
					(function(a,b){
						if(/(android|bb\d+|meego).+mobile|avantgo|bada\/|blackberry|blazer|compal|elaine|fennec|hiptop|iemobile|ip(hone|od)|iris|kindle|lge |maemo|midp|mmp|mobile.+firefox|netfront|opera m(ob|in)i|palm( os)?|phone|p(ixi|re)\/|plucker|pocket|psp|series(4|6)0|symbian|treo|up\.(browser|link)|vodafone|wap|windows ce|xda|xiino/i.test(a)||/1207|6310|6590|3gso|4thp|50[1-6]i|770s|802s|a wa|abac|ac(er|oo|s\-)|ai(ko|rn)|al(av|ca|co)|amoi|an(ex|ny|yw)|aptu|ar(ch|go)|as(te|us)|attw|au(di|\-m|r |s )|avan|be(ck|ll|nq)|bi(lb|rd)|bl(ac|az)|br(e|v)w|bumb|bw\-(n|u)|c55\/|capi|ccwa|cdm\-|cell|chtm|cldc|cmd\-|co(mp|nd)|craw|da(it|ll|ng)|dbte|dc\-s|devi|dica|dmob|do(c|p)o|ds(12|\-d)|el(49|ai)|em(l2|ul)|er(ic|k0)|esl8|ez([4-7]0|os|wa|ze)|fetc|fly(\-|_)|g1 u|g560|gene|gf\-5|g\-mo|go(\.w|od)|gr(ad|un)|haie|hcit|hd\-(m|p|t)|hei\-|hi(pt|ta)|hp( i|ip)|hs\-c|ht(c(\-| |_|a|g|p|s|t)|tp)|hu(aw|tc)|i\-(20|go|ma)|i230|iac( |\-|\/)|ibro|idea|ig01|ikom|im1k|inno|ipaq|iris|ja(t|v)a|jbro|jemu|jigs|kddi|keji|kgt( |\/)|klon|kpt |kwc\-|kyo(c|k)|le(no|xi)|lg( g|\/(k|l|u)|50|54|\-[a-w])|libw|lynx|m1\-w|m3ga|m50\/|ma(te|ui|xo)|mc(01|21|ca)|m\-cr|me(rc|ri)|mi(o8|oa|ts)|mmef|mo(01|02|bi|de|do|t(\-| |o|v)|zz)|mt(50|p1|v )|mwbp|mywa|n10[0-2]|n20[2-3]|n30(0|2)|n50(0|2|5)|n7(0(0|1)|10)|ne((c|m)\-|on|tf|wf|wg|wt)|nok(6|i)|nzph|o2im|op(ti|wv)|oran|owg1|p800|pan(a|d|t)|pdxg|pg(13|\-([1-8]|c))|phil|pire|pl(ay|uc)|pn\-2|po(ck|rt|se)|prox|psio|pt\-g|qa\-a|qc(07|12|21|32|60|\-[2-7]|i\-)|qtek|r380|r600|raks|rim9|ro(ve|zo)|s55\/|sa(ge|ma|mm|ms|ny|va)|sc(01|h\-|oo|p\-)|sdk\/|se(c(\-|0|1)|47|mc|nd|ri)|sgh\-|shar|sie(\-|m)|sk\-0|sl(45|id)|sm(al|ar|b3|it|t5)|so(ft|ny)|sp(01|h\-|v\-|v )|sy(01|mb)|t2(18|50)|t6(00|10|18)|ta(gt|lk)|tcl\-|tdg\-|tel(i|m)|tim\-|t\-mo|to(pl|sh)|ts(70|m\-|m3|m5)|tx\-9|up(\.b|g1|si)|utst|v400|v750|veri|vi(rg|te)|vk(40|5[0-3]|\-v)|vm40|voda|vulc|vx(52|53|60|61|70|80|81|83|85|98)|w3c(\-| )|webc|whit|wi(g |nc|nw)|wmlb|wonu|x700|yas\-|your|zeto|zte\-/i.test(a.substr(0,4)))check = true}
					)(navigator.userAgent||navigator.vendor||window.opera);

					return check;
				}

				if(mobilecheck()){
					document.frm.first_sw.value = "n"; // 모바일로 접속시 이동 경로
				}else{
					document.frm.first_sw.value = "y"; // PC로 접속시 이동 경로
				}

				//알림 팝업
				if(document.frm.first_sw.value == "y"){
					//var position ="width="+w+",height="+h+",left=" + 0 + ",top=" + ((screen.height-h)/2) + ",toolbar="+tb+",directories="+di+",status="+st+",menubar="+mb+",scrollbars="+sb+",resizable="+re+"";
					//window.open('nkp_popup.asp', 'aaaa', position);

					if(getCookie('nkp_popup') != 'nkp_popup'){
						var url = "/main/nkp_popup.asp";
						pop_Window(url, 'nkp알림','scrollbars=yes,width=640,height=700');
					}
				}

				//공지 팝업
				/*if(document.frm.first_sw.value == "y"){
					if(getCookie('nkp_notice') != 'nkp_notice'){
						var url = "/main/nkp_popup_noti.asp";
						pop_Window(url, 'nkp공지','scrollbars=yes,width=645,height=700, left=650, top='+((screen.height-h)/2));
					}
				}*/

				//AS 현황 팝업
				if(document.frm.first_sw.value == "y"){
					if(getCookie('first_as_view') != 'first_as_view'){
						var url = "/main/first_as_view.asp";
						//pop_Window(url,'cccc','scrollbars=yes,width=1400,height=450');
						pop_Window(url,'방문처리 지사별 미처리 현황','scrollbars=yes,width=1020,height=480');	// 2018-11-30 팝업크기 변경
					}
				}
			}

			// 쿠키 가져오기
			function getCookie(cname){
				var name = cname + "=";
				var decodedCookie = document.cookie;
				//var decodedCookie = decodeURIComponent(document.cookie); // 여기서 오류가 남..
				var ca = decodedCookie.split(';');

				for(var i = 0; i <ca.length; i++){
					var c = ca[i];

					while(c.charAt(0) == ' '){
						c = c.substring(1);
					}

					if(c.indexOf(name) == 0){
						return c.substring(name.length, c.length);
					}
				}
				return "";
			}
	</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/main_header.asp" -->
			<!--#include virtual = "/include/main_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/main/nkp_main.asp" method="post" name="frm">
					<fieldset class="srch">
						<legend>조회영역</legend>
						<dl>
							<dt>조건 검색</dt>
							<dd>
								<p>
                                <input type="radio" name="board_gubun" value="0" <% If board_gubun = "0" Then %>checked<% End If %> style="width:30px">총괄
                                <input type="radio" name="board_gubun" value="1" <% If board_gubun = "1" Then %>checked<% End If %> style="width:30px">사내공지
                                <input type="radio" name="board_gubun" value="3" <% If board_gubun = "3" Then %>checked<% End If %> style="width:30px">A/S공지
                                <input type="radio" name="board_gubun" value="4" <% If board_gubun = "4" Then %>checked<% End If %> style="width:30px">자료실&nbsp;&nbsp;
                                <strong>조건 : </strong>
                                <select name="condi" style="width:100px">
                                    <option value="all" <%If condi = "all" Then %>selected<% End If %>>전체</option>
                                    <option value="board_title" <%If condi = "board_title" Then %>selected<% End If %>>제목</option>
                                    <option value="board_body" <%If condi = "board_body" Then %>selected<% End If %>>내용</option>
                                    <option value="reg_name" <%If condi = "reg_name" Then %>selected<% End If %>>작성자</option>
                                </select>
                                <input name="condi_value" type="text" value="<%=condi_value%>">
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
                                </p>
                            </dd>
						</dl>
					</fieldset>
					<div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
							<colgroup>
								<col width="*" >
								<col width="50%" >
								<col width="10%" >
								<col width="15%" >
								<col width="10%" >
								<col width="10%" >
							</colgroup>
							<thead>
								<tr>
									<th class="first" scope="col">순번</th>
									<th scope="col">제목</th>
									<th scope="col">작성자</th>
									<th scope="col">작성일</th>
									<th scope="col">조회수</th>
									<th scope="col">첨부</th>
								</tr>
							</thead>
							<tbody>
							<%
							Dim seq, board_title, Path, board_seq, reg_name
							Dim reg_date, read_cnt, att_file, main_title
							Dim intstart, intend, first_page, i

    						seq = total_record - ( page - 1 ) * pgsize + 1

    						If IsArray(arrBoard) Then
								For i = LBound(arrBoard) To UBound(arrBoard, 2)
									board_seq = arrBoard(1, i)
									board_gubun = arrBoard(2, i)
									reg_name = arrBoard(3, i)
									board_title = arrBoard(4, i)
									reg_date = arrBoard(5, i)
									read_cnt = arrBoard(6, i)
									att_file = arrBoard(7, i)

									main_title = ""

									If board_gubun = "0" Then
										If board_gubun = "1" Then
											main_title = "[사내공지]"
										ElseIf board_gubun = "2" Then
											main_title = "[사내게시판]"
										ElseIf board_gubun = "3" Then
											main_title = "[A/S공지]"
										Else
											main_title = "[자료실]"
										End If
									End If
                                %>
                                <tr>
                                    <td class="first"><%=seq%></td>
                                    <td class="left">
                                        <strong><%=main_title%></strong>&nbsp;
										<a href="/board/board_view.asp?board_back=<%=board_gubun%>&board_gubun=<%=board_gubun%>&board_seq=<%=board_seq%>&page=<%=page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>"><%=board_title%></a>

                                        <input name="board_seq" type="hidden" id="board_seq" value="<%=board_seq%>">

                                        <%If reg_date > new_date Then%>
										<img src="/image/new.gif" width="24" height="11" border="0">
                                        <%End If%>
                                    </td>
                                    <td><%=reg_name%></td>
                                    <td><%=reg_date%></td>
                                    <td><%=read_cnt%></td>
                                    <td>
                                    <% If att_file <> "" Then
                                        path = "/nkp_upload"
                                    %>
                                        <a href="/att_file_download.asp?path=<%=path%>&att_file=<%=att_file%>"><img src="/image/att_file.gif" border="0"></a>
                                    <% Else%>
                                        &nbsp;
                                    <% End If %>
                                    </td>
                                </tr>
                                <%
	                            seq = seq - 1

								Next

							Else
								Response.Write "<tr><td colspan='14' style='text-weight:bold;'>해당 내용이 없습니다.</td></tr>"
							End If
							%>
							</tbody>
						</table>
					</div>
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td width="25%">
                                <div class="btnCenter">
                                    <%	if c_grade = "0" then %>
                                    <a href="#" onClick="pop_Window('/board/popup_file_up.asp','popup_file_up_popup','scrollbars=yes,width=500,height=200')" class="btnType04">팝업이미지올리기</a>
                                    <%	end if	%>
                                </div>
                            </td>
                            <td>
							<%
							'Paging 호출
							Call Page_Navi_Ver2(page, be_pg, url_param, total_record, pgsize)
							%>
                            </td>
                            <td width="25%">
                                <div class="btnCenter">
                                    <a href="#" onClick="pop_Window('/board/board_write.asp?board_gubun=<%=board_gubun%>','board_write_popup','scrollbars=yes,width=1250,height=600')" class="btnType04">글올리기</a>
                                </div>
                            </td>
                        </tr>
                    </table>
                <!--<input type="hidden" name="board_back" value="<%=board_gubun%>">-->
                <input type="hidden" name="first_sw" value="<%=first_sw%>">
				</form>
			</div>
		</div>
	</body>
</html>
<!--#include virtual="/common/inc_bottom.asp"-->