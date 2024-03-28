<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")

curr_date = mid(cstr(now()),1,10)
work_date = mid(cstr(now()),1,10)

company = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

if u_type = "U" then

    work_date1 = request("work_date")    
    mg_ce_id = request("mg_ce_id")    

	sql = "select * from overtime where work_date = '" + work_date1 + "' and mg_ce_id = '" + mg_ce_id + "'"
	set rs = dbconn.execute(sql)

	sql="select * from memb where user_id = '" + rs("mg_ce_id") + "'"
	set rs_memb=dbconn.execute(sql)

	if	rs_memb.eof or rs_memb.bof then
		mg_ce = "ERROR"
	else
		mg_ce = rs_memb("user_name")
	end if
	rs_memb.close()						

    work_date2    = rs("end_date")
	if isnull(rs("acpt_no")) then
		acpt_no = 0
	else
		acpt_no = rs("acpt_no")
    end if    
	mg_ce_id     = rs("mg_ce_id")
	company      = rs("company")
	dept         = rs("dept")
	work_item    = rs("work_item")
	from_time    = rs("from_time")
	to_time      = rs("to_time")
	work_gubun   = rs("work_gubun")
	overtime_amt = int(rs("overtime_amt"))
	work_memo    = rs("work_memo")
	cancel_yn    = rs("cancel_yn")
	reg_id       = rs("reg_id")
	reg_user     = rs("reg_user")
	reg_date     = rs("reg_date")
	mod_id       = rs("mod_id")
	mod_user     = rs("mod_user")
	mod_date     = rs("mod_date")
	rs.close()

    Select Case (WeekDay(work_date1))
        Case 1 week1 = "일"
        Case 2 week1 = "월"
        Case 3 week1 = "화"
        Case 4 week1 = "수"
        Case 5 week1 = "목"
        Case 6 week1 = "금"
        Case 7 week1 = "토"
    End Select

    Select Case (WeekDay(work_date2))
        Case 1 week2 = "일"
        Case 2 week2 = "월"
        Case 3 week2 = "화"
        Case 4 week2 = "수"
        Case 5 week2 = "목"
        Case 6 week2 = "금"
        Case 7 week2 = "토"
    End Select

    holiday1 = ""
    holiday2 = ""
    
    sql = " SELECT holiday_memo  FROM holiday WHERE holiday = '" & work_date1 & "'  "
    'Response.write sql&chr(13)
    rs.Open sql, Dbconn, 1
    
    if not (rs.eof or rs.bof) then
        holiday1 = rs("holiday_memo")  	
    end if
    rs.close
    
    sql = " SELECT holiday_memo  FROM holiday WHERE holiday = '" & work_date2 & "'  "
    'Response.write sql&chr(13)
    rs.Open sql, Dbconn, 1
    
    if not (rs.eof or rs.bof) then
        holiday2 = rs("holiday_memo")  	
    end if
    rs.close
    

	title_line = "야특근 지급 변경"
end if

if end_yn = "Y" then
	end_view = "마감"
  else
  	end_view = "진행"
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">

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
			function chkfrm() 
            {
				if (confirm('변경하시겠습니까?')==true) 
 					 return true;
				else return false;
			}

            $(function() 
            {
                switch ('<%=work_gubun%>') 
                {
                    case '스케쥴근무':
                        document.getElementById('idOnday').style.display = '';                     
                        document.getElementById('idFromTo').style.display = 'none';                     
                        break;
                    case '연장근무 휴일':
                        document.getElementById('idOnday').style.display = 'none';
                        document.getElementById('idFromTo').style.display = '';                     
                        break;
                    case '스케쥴근무':
                        document.getElementById('idOnday').style.display = '';
                        document.getElementById('idFromTo').style.display = 'none';
                        break;
                    default:
                        document.getElementById('idOnday').style.display = 'none';
                        document.getElementById('idFromTo').style.display = '';                     
                }
            });
        </script>
	</head>
	<body>
        <div id="container">
            <h3 class="tit"><%=title_line%></h3>
            <form action="overtime_cancel_save.asp" method="post" name="frm">
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
                                    <%=work_date1%>&nbsp;<%=week1%>&nbsp;<%=holiday1%>
                                    &nbsp;~&nbsp;
                                    <%=work_date2%>&nbsp;<%=week2%>&nbsp;<%=holiday2%>
                                    <input name="work_date1" type="hidden" value="<%=work_date1%>">
                                </td>
								<th>작업자</th>
								<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)</td>
							</tr>
							<tr>
								<th class="first">회사명</th>
								<td class="left"><%=company%></td>
								<th>부서명</th>
								<td class="left"><%=dept%></td>
							</tr>
							<tr>
								<th class="first">작업항목</th>
								<td class="left"><%=work_gubun%></td>
								<th>작업시간</th>
								<td class="left">
                                    <span id="idOnday">
                                        <b>1일</b>&nbsp;
                                    </span>
                                    <span id="idFromTo">
                                        <%=from_time%> ~ <%=to_time%>
                                    </span>
                                </td>
							</tr>
							<tr>
								<th class="first">작업내용</th>
                                <td colspan="3" class="left"><%=work_memo%></td>
							</tr>
							<tr>
								<th class="first">취소여부</th>
								<td class="left">
								<input type="radio" name="cancel_yn" value="Y" <% if cancel_yn = "Y" then %>checked<% end if %> style="width:40px" ID="Radio1">취소
				                <input type="radio" name="cancel_yn" value="N" <% if cancel_yn = "N" then %>checked<% end if %> style="width:40px" ID="Radio2">지급
								</td>
                                <th>마감여부</th>
								<td class="left"><%=end_view%></td>
							</tr>
							<tr>
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
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="mg_ce_id" value="<%=mg_ce_id%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

