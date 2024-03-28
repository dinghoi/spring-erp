<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim curr_date, curr_year, curr_month, curr_day
Dim ck_sw, page, view_condi, insu_code, insu_yyyy
Dim pgsize, start_page, stpage, be_pg, pg_url

Dim rsCount, total_record, total_page, title_line
Dim rsInsure, rsEtc

view_condi = f_Request("view_condi")
page = f_Request("page")

be_pg = "/pay/insa_pay_rule_mg.asp"

curr_date = Mid(CStr(Now()), 1, 10)
curr_year = Mid(CStr(Now()), 1, 4)
curr_month = Mid(CStr(Now()), 6, 2)
curr_day = Mid(CStr(Now()), 9, 2)

If view_condi = "" Then
	view_condi = "5501국민연금"
End If

insu_code = Mid(CStr(view_condi), 1, 4)
insu_yyyy = Mid(CStr(Now()), 1, 4) '귀속년월

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

pg_url = "&view_condi="&view_condi

'Record Count
objBuilder.Append "SELECT COUNT(*) FROM pay_insurance WHERE insu_id = '" & insu_code & "' "

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(RsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

objBuilder.Append "SELECT * "
objBuilder.Append "FROM pay_insurance WHERE insu_id = '" & insu_code & "' "
objBuilder.Append "ORDER BY insu_id, insu_yyyy, insu_class DESC "
objBuilder.Append "LIMIT " & stpage &  "," & pgsize

Set rsInsure = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "4대보험 요율관리"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>급여관리 시스템</title>
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	<link href="/include/style.css" type="text/css" rel="stylesheet">
	<script src="/java/jquery-1.9.1.js"></script>
	<script src="/java/jquery-ui.js"></script>
	<script src="/java/common.js" type="text/javascript"></script>
	<script src="/java/ui.js" type="text/javascript"></script>
	<script type="text/javascript" src="/java/js_form.js"></script>

	<script type="text/javascript">
		function getPageCode(){
			return "1 0";
		}

		function frmcheck(){
			if(formcheck(document.frm) && chkfrm()){
				document.frm.submit ();
			}
		}

		function chkfrm(){
			if(document.frm.view_condi.value == ""){
				alert ("필드조건을 선택하시기 바랍니다");
				return false;
			}
			return true;
		}
	</script>
</head>
<body>
	<div id="wrap">
		<!--#include virtual = "/include/insa_pay_header.asp" -->
		<!--#include virtual = "/include/insa_pay_rule_menu.asp" -->
		<div id="container">
			<h3 class="insa"><%=title_line%></h3>
			<form action="insa_pay_insurance_mg.asp" method="post" name="frm">
			<fieldset class="srch">
				<legend>조회영역</legend>
				<dl>
					<dt>보험 검색</dt>
					<dd>
						<p>
						   <strong>보험종류 : </strong>
						  <%
							objBuilder.Append "SELECT emp_etc_code, emp_etc_name FROM emp_etc_code "
							objBuilder.Append "WHERE emp_etc_type = '55' "
							objBuilder.Append "ORDER BY emp_etc_code ASC "

							Set rsEtc = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()
						  %>
							<label>
							<select name="view_condi" id="view_condi" type="text" style="width:150px">
						  <%
							Do Until rsEtc.EOF
						  %>
								<option value='<%=rsEtc("emp_etc_code")%><%=rsEtc("emp_etc_name")%>' <%If view_condi = rsEtc("emp_etc_name") Then%>selected<%End If %>><%=rsEtc("emp_etc_code")%><%=rsEtc("emp_etc_name")%></option>
						  <%
								rsEtc.MoveNext()
							Loop
							rsEtc.Close() : Set rsEtc = Nothing
						  %>
							</select>
							</label>
							<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
						</p>
					</dd>
				</dl>
			</fieldset>
			<div class="gView">
				<table cellpadding="0" cellspacing="0" class="tableList">
					<colgroup>
						<col width="8%" >
						<col width="8%" >
						<col width="8%" >
						<col width="8%" >
						<col width="8%" >
						<col width="8%" >
						<col width="8%" >
						<col width="8%" >
						<col width="*" >
						<col width="4%" >
					</colgroup>
					<thead>
						<tr>
						   <th rowspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">기준<br>적용년월</th>
						   <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">등급</th>
						   <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">보수월액</th>
						   <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">평균<br>보수월액</th>
						   <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">사업자 가입자</th>
						   <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">비고</th>
						   <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">변경</th>
					   </tr>
					   <tr>
						  <th scope="col" style=" border-left:1px solid #e3e3e3;">이상</th>
						  <th scope="col" style=" border-bottom:1px solid #e3e3e3;">미만</th>
						  <th scope="col">합계</th>
						  <th scope="col">근로자</th>
						  <th scope="col">사용자</th>
					   </tr>
					</thead>
					<tbody>
					<%
					Do Until rsInsure.EOF
					%>
						<tr>
							<td class="first"><%=rsInsure("insu_yyyy")%>&nbsp;</td>
							<td class="left"><%=rsInsure("insu_class")%>&nbsp;</td>
							<td class="right"><%=FormatNumber(rsInsure("from_amt"),0)%>&nbsp;</td>
							<td class="right"><%=FormatNumber(rsInsure("to_amt"),0)%>&nbsp;</td>
							<td class="right"><%=FormatNumber(rsInsure("st_amt"),0)%>&nbsp;</td>
							<td class="right"><%=FormatNumber(rsInsure("hap_rate"),3)%>&nbsp;</td>
							<td class="right"><%=FormatNumber(rsInsure("emp_rate"),3)%>&nbsp;</td>
							<td class="right"><%=FormatNumber(rsInsure("com_rate"),3)%>&nbsp;</td>
							<td><%=rsInsure("insu_comment")%>&nbsp;</td>
							<td>
							 <a href="#" onClick="pop_Window('/pay/insa_pay_insurance_add.asp?insu_id=<%=insu_code%>&view_condi=<%=view_condi%>&insu_class=<%=rsInsure("insu_class")%>&insu_yyyy=<%=rsInsure("insu_yyyy")%>&u_type=U','4대보험요율 변경','scrollbars=yes,width=750,height=300')">수정</a>
							</td>
						</tr>
					<%
						rsInsure.MoveNext()
					Loop
					rsInsure.Close() : Set rsInsure = Nothing
					%>
					</tbody>
				</table>
			</div>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			  <tr>
				<td>
				<%
				'Page Navi
				Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
				%>
				</td>
				<td width="20%">
				<div class="btnCenter">
				<a href="#" onClick="pop_Window('/pay/insa_pay_insurance_add.asp?insu_id=<%=insu_code%>&view_condi=<%=view_condi%>','4대보험요율 등록','scrollbars=yes,width=750,height=300')" class="btnType04">4대보험요율 등록</a>
				</div>
				</td>
			  </tr>
			  </table>
		</form>
	</div>
</div>
</body>
</html>
<!--#include virtual = "/common/inc_footer.asp" -->