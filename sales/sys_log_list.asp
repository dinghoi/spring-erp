<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--include virtual="/include/db_create.asp" -->
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
Dim title_line, savefilename, i, rsLog
Dim work_month, from_date, to_date
Dim page, page_cnt, pg_cnt, be_pg, pgsize, start_page, stpage
Dim pg_url, total_record

page = f_Request("page")
page_cnt = f_Request("page_cnt")
pg_cnt = CInt(f_Request("pg_cnt"))
from_date = f_Request("from_date")
to_date = f_Request("to_date")

title_line = "시스템 로그 정보"
be_pg = "/sales/sys_log_list.asp"

'리스트 페이징 설정
pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

pg_url = "&from_date="&from_date&"&to_date="&to_date

'조회 날짜 설정
work_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)

If from_date = "" Then
    from_date = Mid(work_month, 1, 4) & "-" & Mid(work_month, 5, 2) & "-01"
End If

If to_date = "" Then
    to_date = CStr(DateAdd("d", -1, DateAdd("m", 1, DateValue(from_date))))
End If

objBuilder.Append "SELECT (SELECT COUNT(*) FROM emp_sys_log "
objBuilder.Append "		WHERE reg_date BETWEEN '"&from_date&"' AND '"&to_date&"') AS 'tot_cnt', "
objBuilder.Append "	logt.emp_seq, memt.emp_no, memt.user_name, memt.user_grade, memt.position, "
objBuilder.Append "	memt.emp_company, memt.bonbu, memt.saupbu, "
objBuilder.Append "	logt.remote_ip, logt.menu_name, logt.menu_title, logt.excel_yn, SUBSTRING(logt.reg_date, 1, 11) AS 'reg_date' "
objBuilder.Append "FROM emp_sys_log AS logt "
objBuilder.Append "INNER JOIN memb AS memt ON logt.emp_no = memt.emp_no "
objBuilder.Append "WHERE logt.reg_date BETWEEN '"&from_date&"' AND '"&to_date&"' "
objBuilder.Append  "ORDER BY logt.reg_date DESC "
objBuilder.Append "LIMIT "&stpage&", "&pgsize

Set rsLog = DBConn.Execute(objBuilder.ToString)
objBuilder.Clear()

If rsLog.EOF Or rsLog.BOF Then
	total_record = 0
Else
	total_record = CInt(rsLog("tot_cnt"))
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>

		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			function getPageCode(){
				return "3 1";
			}

			$(function() {

				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=from_date%>" );

				$( "#datepicker2" ).datepicker();
				$( "#datepicker2" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker2" ).datepicker("setDate", "<%=to_date%>" );
			});


			function frmcheck(){
				//var st_date = $("#datepicker1").datepicker({dateFormat: 'dd-mm-yy'});
				//var end_date = $("#datepicker2").datepicker({dateFormat : 'dd-mm-yy'});

				var fDate = $("#datepicker1").datepicker('getDate');
				var lDate = $("#datepicker2").datepicker('getDate');

				//console.log(fDate);
				//console.log(lDate);

				var diff = new Date(lDate - fDate);
				var days = diff/1000/60/60/24;


				if(fDate = ""){
					alert("검색 시작년월일이 없습니다.");
					return false;
				}

				if(lDate = ""){
					alert("검색 종료년월일이 없습니다.");
					return false;
				}

				//console.log(days);
				//return false;

				if(days < 0){
					alert("검색 시작년월일이 종료 년월일 보다 작을 수 없습니다.");
					return false;
				}

				document.logFrm.submit();
				return;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/sales_code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/sales/sys_log_list.asp" method="post" name="logFrm">
					<fieldset class="srch">
						<legend>조회영역</legend>
						<dl>
							<dt>처리조건</dt>
							<dd>
								<p>
									<label>
										&nbsp;&nbsp;<strong>시작일자&nbsp;</strong> :
										<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker1">
										&nbsp;~&nbsp;
										&nbsp;&nbsp;<strong>종료일자&nbsp;</strong> :
										<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker2">
									</label>
									<a href="#" onclick="frmcheck();"><img src="/image/but_ser.jpg" alt="검색" /></a>
								</p>
							</dd>
						</dl>
					</fieldset>
				</form>

				<div class="gView">
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="60%" height="356" valign="top">
								<table cellpadding="0" cellspacing="0" class="tableList">
								<colgroup>
									<col width="5%" >
									<col width="5%" >
									<col width="5%" >
									<col width="5%" >
									<col width="5%" >
									<col width="5%" >
									<col width="5%" >
									<col width="5%" >
									<col width="10%" >
									<col width="*" >
									<col width="5%" >
									<col width="7%" >
								</colgroup>
								<thead>
									<tr>
										<th class="first" scope="col">사번</th>
										<th scope="col">성명</th>
										<th scope="col">직급</th>
										<th scope="col">직책</th>
										<th scope="col">회사</th>
										<th scope="col">본부</th>
										<th scope="col">사업부</th>
										<th scope="col">접근IP</th>
										<th scope="col">접근메뉴</th>
										<th scope="col">메뉴명</th>
										<th scope="col">엑셀다운로드</th>
										<th scope="col">접근시간</th>
									</tr>
								</thead>
								<tbody>
								<%
								Do Until rsLog.EOF
								%>
									<tr>
										<td class="first"><%=rsLog("emp_no")%></td>
										<td><%=rsLog("user_name")%></td>
										<td><%=rsLog("user_grade")%></td>
										<td><%=rsLog("position")%></td>
										<td><%=rsLog("emp_company")%></td>
										<td><%=rsLog("bonbu")%></td>
										<td><%=rsLog("saupbu")%></td>
										<td><%=rsLog("remote_ip")%></td>
										<td><%=rsLog("menu_name")%></td>
										<td><%=rsLog("menu_title")%></td>
										<td><%=rsLog("excel_yn")%></td>
										<td><%=rsLog("reg_date")%></td>
									</tr>
								<%
									rsLog.MoveNext()
								Loop
								rsLog.Close() : Set rsLog = Nothing
								%>
								</tbody>
							</td>
						</tr>
					</table>
				</div>

				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="/sales/excel/sys_log_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
					%>
                    </td>
			      </tr>
				  </table>

			</div>
		</div>
	</body>
</html>
<%
DBConn.Close() : Set DBConn = Nothing
%>