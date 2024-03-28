<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

be_pg = "insa_individual_confirm.asp"

cfm_use =""
cfm_use_dept =""
cfm_comment =""

win_sw = "close"
Page=Request("page")

ck_sw=Request("ck_sw")


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'sql = "select * from emp_master where emp_no = '" + in_empno + "'"
sql = "select emtt.emp_no, emtt.emp_name, emtt.emp_job, emtt.emp_position, "
sql = sql & "	emtt.emp_in_date, emtt.emp_birthday, emtt.emp_company, "
sql = sql & "	emtt.emp_org_name, eomt.org_name, eomt.org_company, eomt.org_bonbu, "
sql = sql & "	eomt.org_saupbu, eomt.org_team "
sql = sql & "FROM emp_master AS emtt "
sql = sql & "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
sql = sql & "WHERE emtt.emp_no < '900000' "
sql = sql & "	AND (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' "
sql = sql & "		OR emtt.emp_end_date  = '0000-00-00') "
sql = sql & "	AND emp_no = '"&in_empno&"' "

Rs.Open Sql, Dbconn, 1

title_line = " 제증명서 신청/발급 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
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
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function chkfrm() {
				if (document.frm.in_empno.value == "") {
					alert ("사번을 입력하시기 바랍니다");
					return false;
				}	
				return true;
			}
            function s_sinchung(val, val2, val3, val4, val5) {

            if (!confirm("재직증명서를 신청하시겠습니까 ?")) return;
            var frm = document.frm;
            document.frm.in_empno.value = val;
            document.frm.in_name.value = val2;
			
            if (document.getElementById(val3).value == "")
            { alert("신청 용도를 선택해주세요!"); return; }
			
            if (document.getElementById(val4).value == "")
            { alert("사용처를 입력하십시요!"); return; }

            document.frm.cfm_use.value = document.getElementById(val3).value;
            document.frm.action = "/insa/insa_certificate_print.asp";
            document.frm.submit();
            }	
			function s_sinchung2(val, val2, val3, val4, val5) {

            if (!confirm("경력증명서를 신청하시겠습니까 ?")) return;
            var frm = document.frm;
            document.frm.in_empno.value = val;
            document.frm.in_name.value = val2;
			

            if (document.getElementById(val3).value == "")
            { alert("신청 용도를 선택해주세요!"); return; }
			
			if (document.getElementById(val4).value == "")
            { alert("사용처를 입력하십시요!"); return; }

            document.frm.cfm_use.value = document.getElementById(val3).value;
            document.frm.action = "/insa/insa_certificate_career.asp";
            document.frm.submit();
            }	

		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psawo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/insa_individual_confirm.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
							<strong>사번 : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=in_empno%>" style="width:100px; text-align:left">
								</label>
                            <strong>성명 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=in_name%>" readonly="true" style="width:150px; text-align:left">
								</label>
                                
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
							<col width="8%" >
							<col width="10%" >
							<col width="10%" >
                            <col width="10%" >
                            <col width="*" >
                            <col width="6%" >
                            <col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성명</th>
								<th scope="col">직위</th>
								<th scope="col">직책</th>
								<th scope="col">입사일</th>
								<th scope="col">생년월일</th>
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
								<th scope="col" style="background:#FFC">용도</th>
								<th scope="col" style="background:#FFC">사용처</th>
                                <th scope="col" style="background:#FFC">비고</th>
                                <th scope="col" style="background:#FFC">재직</th>
                                <th scope="col" style="background:#FFC">경력</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td><%=rs("emp_name")%>&nbsp;</td>
                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td><%=rs("org_company")%>&nbsp;</td>
                                <td><%=rs("org_name")%>&nbsp;</td>
                                <td class="left">
                                <select name="cfm_use" id="cfm_use" value="<%=cfm_use%>" style="width:110px">
			            	        <option value="" <% if cfm_use = "" then %>selected<% end if %>>선택</option>
				                    <option value='대출용' <%If cfm_use = "대출용" then %>selected<% end if %>>대출용</option>
                                    <option value='보증용' <%If cfm_use = "보증용" then %>selected<% end if %>>보증용</option>
                                    <option value='입찰참여제출용' <%If cfm_use = "입찰참여제출용" then %>selected<% end if %>>입찰참여제출용</option>
                                    <option value='교육기관제출용' <%If cfm_use = "교육기관제출용" then %>selected<% end if %>>교육기관제출용</option>
                                    <option value='관공서제출용' <%If cfm_use = "관공서제출용" then %>selected<% end if %>>관공서제출용</option>
                                    <option value='법원제출용' <%If cfm_use = "법원제출용" then %>selected<% end if %>>법원제출용</option>
                                    <option value='회사제출용' <%If cfm_use = "회사제출용" then %>selected<% end if %>>회사제출용</option>
                                    <option value='비자발급용' <%If cfm_use = "비자발급용" then %>selected<% end if %>>비자발급용</option>
                                    <option value='취업용' <%If cfm_use = "취업용" then %>selected<% end if %>>취업용</option>
                                    <option value='노동부(청)제출용' <%If cfm_use = "노동부(청)제출용" then %>selected<% end if %>>노동부(청)제출용</option>
                                    <option value='위임관계확인용' <%If cfm_use = "위임관계확인용" then %>selected<% end if %>>위임관계확인용</option>
                                    <option value='금융기관제출용' <%If cfm_use = "금융기관제출용" then %>selected<% end if %>>금융기관제출용</option>
                                    <option value='협회제출용' <%If cfm_use = "협회제출용" then %>selected<% end if %>>협회제출용</option>
                                    <option value='취업확인용' <%If cfm_use = "취업확인용" then %>selected<% end if %>>취업확인용</option>
                                    <option value='은행제출용' <%If cfm_use = "은행제출용" then %>selected<% end if %>>은행제출용</option>
                                    <option value='학교제출용' <%If cfm_use = "학교제출용" then %>selected<% end if %>>학교제출용</option>
                                    <option value='보험사제출용' <%If cfm_use = "보험사제출용" then %>selected<% end if %>>보험사제출용</option>
                                    <option value='카드사제출용' <%If cfm_use = "카드사제출용" then %>selected<% end if %>>카드사제출용</option>
                                    <option value='증권사제출용' <%If cfm_use = "증권사제출용" then %>selected<% end if %>>증권사제출용</option>
                                </select> 
                                </td>
                                <td class="left">
								<input name="cfm_use_dept" type="text" id="cfm_use_dept" style="width:100px" onKeyUp="checklength(this,30)" value="<%=cfm_use_dept%>">
                                </td>    
                                <td class="left">
								<input name="cfm_comment" type="text" id="cfm_comment" style="width:170px" onKeyUp="checklength(this,50)" value="<%=cfm_comment%>">
                                </td>                                
                                <td>
                                <input type="image" name="rptCert$ctl00$btnRequest" id="rptCert_ctl00_btnRequest" src="/image/b_certifi.jpg" alt="재직증명서 신청" onclick="s_sinchung('<%=rs("emp_no")%>','<%=rs("emp_name")%>', 'cfm_use', 'cfm_use_dept', 'cfm_comment');return false;" style="border-width:0px;" />
                                </td>
                                <% if insa_grade = "0" then %>
                                <td>
                                <input type="image" name="rptCert$ctl01$btnRequest" id="rptCert_ctl01_btnRequest" src="/image/b_certifi.jpg" alt="경력증명서 신청" onclick="s_sinchung2('<%=rs("emp_no")%>','<%=rs("emp_name")%>', 'cfm_use', 'cfm_use_dept', 'cfm_comment');return false;" style="border-width:0px;" />
                                </td>
                                <% end if %>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
                  <input type="hidden" name="emp_empno" value="<%=in_empno%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

