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
Dim page, page_cnt, pg_cnt, be_pg, curr_date,cfm_use, cfm_use_dept, cfm_comment
Dim pgsize, start_page, stpage, order_sql, where_sql, str_param
Dim title_line, rsCount, total_record, total_page
Dim view_sort

page = f_Request("page")
page_cnt = f_Request("page_cnt")
pg_cnt = CInt(f_Request("pg_cnt"))
view_sort = f_Request("view_sort")

curr_date = DateValue(Mid(CStr(Now()), 1, 10))
be_pg = "/person/insa_confirm_report.asp"

cfm_use =""
cfm_use_dept =""
cfm_comment =""

If view_sort = "" Then
	view_sort = "DESC"
End If

title_line = "제증명서 발급현황"
pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

str_param = "&view_sort="&view_sort

where_sql = " WHERE cfm_empno = '"&user_id&"' "
order_sql = " ORDER BY cfm_date DESC, cfm_number DESC, cfm_seq "&view_sort&" "

objBuilder.Append "SELECT COUNT(*) FROM emp_confirm WHERE cfm_empno='"&user_id&"';"

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record/pgsize) 'Result.PageCount
Else
	total_page = Int((total_record/pgsize) + 1)
End If

objBuilder.Append "SELECT * "
objBuilder.Append "FROM emp_confirm "
objBuilder.Append where_sql&order_sql
objBuilder.Append "LIMIT "&stpage&","&pgsize

Set rsConf = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무관리</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
		</script>
		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psawo_menu.asp" -->
            <div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
							<strong>사번 : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=user_id%>" class="no-input" style="width:80px;" readonly/>
								</label>
                            <strong>성명 : </strong>
                            <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=user_name%>" class="no-input" style="width:80px;" readonly/>
							</label>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%">
							<col width="10%">
							<col width="6%">
                            <col width="10%">
                            <col width="14%">
							<col width="14%">
							<col width="10%">
							<col width="10%">
							<col width="10%">
							<col width="6%">
							<col width="*">
						</colgroup>
						<thead>
						  <tr>
							<th class="first" scope="col">발급일</th>
							<th scope="col">발급번호</th>
                            <th scope="col">제증명</th>
							<th scope="col">용도</th>
                            <th scope="col">사용처</th>
                            <th scope="col">기타사항</th>
							<th scope="col">주민번호</th>
  						    <th scope="col">회사</th>
                            <th scope="col">소속</th>
                            <th scope="col">직위</th>
                            <th scope="col">직책</th>
						  </tr>
						</thead>
						<tbody>
						<%
						Dim rsConf, cfm_empno, rs_emp, emp_job, emp_position

						If rsConf.EOF Or rsConf.BOF Then
							Response.Write "<tr><td colspan='11' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						Else
							Do Until rsConf.EOF
								cfm_empno = rsConf("cfm_empno")

								If f_toString(cfm_empno, "") <> "" Then
									objBuilder.Append "SELECT emp_job, emp_position FROM emp_master "
									objbuilder.Append "WHERE emp_no = '"&cfm_empno&"';"

									Set rs_emp = DBConn.Execute(objBuilder.Tostring())
									objBuilder.Clear()

									If Not rs_emp.eof Then
										emp_job = rs_emp("emp_job")
										emp_position = rs_emp("emp_position")
									End If
									rs_emp.Close() : Set rs_emp = Nothing
								End If
	           			%>
							<tr>
								<td class="first"><%=rsConf("cfm_date")%></td>
                                <td>제&nbsp;<%=rsConf("cfm_number")%>-<%=rsConf("cfm_seq")%>&nbsp;호</td>
                                <td><%=rsConf("cfm_type")%>&nbsp;</td>
                                <td><%=rsConf("cfm_use")%>&nbsp;</td>
								<td><%=rsConf("cfm_use_dept")%>&nbsp;</td>
                                <td><%=rsConf("cfm_comment")%>&nbsp;</td>
                                <td><%=rsConf("cfm_person1")%>-<%=rsConf("cfm_person2")%>&nbsp;</td>
                                <td><%=rsConf("cfm_company")%>&nbsp;</td>
                                <td><%=rsConf("cfm_org_name")%>&nbsp;</td>
                                <td><%=emp_job%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
							</tr>
						<%
								rsConf.MoveNext()
							Loop
						End If
						rsConf.Close() : Set rsConf = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                    <td>
					<%
					'page navigator[허정호_20210720]
					Call Page_Navi(page, be_pg, str_param, total_page)

					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
			      </tr>
				</table>
		</div>
	</div>
	</body>
</html>