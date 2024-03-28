<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_individual_resign.asp"

cfm_use =""
cfm_use_dept =""
cfm_comment =""

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")

ck_sw=Request("ck_sw")
win_sw = "close"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

If Request.Form("in_empno")  <> "" Then 
   Sql = "SELECT * FROM emp_master where emp_no = '"&in_empno&"'"
   Set rs_emp = DbConn.Execute(SQL)
   if not Rs_emp.eof then
      in_name = rs_emp("emp_name")
	  else
      response.write"<script language=javascript>"
	  response.write"alert('등록된 직원이 아닙니다....');"		
	  response.write"</script>"
	  Response.End	
   end if
   rs_emp.close()
End If

sql = "select * from emp_master where emp_no = '" + in_empno + "'"
Rs.Open Sql, Dbconn, 1


title_line = " 사직원 제출(공사중)....."

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "5 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=end_date%>" );
			});	  
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

            if (!confirm("사직원을 신청하시겠습니까 ?")) return;
            var frm = document.frm;
            document.frm.in_empno.value = val;
            document.frm.in_name.value = val2;
			document.frm.in_name.value = val3;
			document.frm.in_name.value = val5;

            if (document.getElementById(val3).value == "")
            { alert("사직일을 입력해주세요!"); return; }
			if (document.getElementById(val4).value == "")
            { alert("사직유형을 선택해주세요!"); return; }

            document.frm.cfm_use.value = document.getElementById(val4).value;
            document.frm.action = "insa_resign_print.asp";
            document.frm.submit();
            }	
			
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_pappo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_individual_resign.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
							<strong>사번 : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=in_empno%>" readonly="true" style="width:100px; text-align:left">
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
							<col width="6%" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="8%" >
                            <col width="8%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
                            <col width="10%" >
                            <col width="12%" >
                            <col width="3%" >
                            <col width="3%" >
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
								<th scope="col" style="background:#FFC">사직일자</th>
								<th scope="col" style="background:#FFC">사직유형</th>
                                <th scope="col" style="background:#FFC">비고</th>
                                <th colspan="2" scope="col" style="background:#FFC">사직서제출</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td><%=rs("emp_company")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td>
								 <input name="end_date" type="text" size="10" readonly="true" id="datepicker" style="width:60px;">&nbsp;</td>
                                <td>
                                <select name="end_type" id="end_type" value="<%=end_type%>" style="width:90px">
			            	        <option value="" <% if end_type = "" then %>selected<% end if %>>선택</option>
				                    <option value='회사사정' <%If end_type = "회사사정" then %>selected<% end if %>>회사사정</option>
                                    <option value='명예퇴직' <%If end_type = "명예퇴직" then %>selected<% end if %>>명예퇴직</option>
                                    <option value='개인사정' <%If end_type = "개인사정" then %>selected<% end if %>>개인사정</option>
                                    <option value='징계' <%If end_type = "징계" then %>selected<% end if %>>징계</option>
                                    <option value='육아' <%If end_type = "육아" then %>selected<% end if %>>육아</option>
                                    <option value='간병' <%If end_type = "간병" then %>selected<% end if %>>간병</option>
                                    <option value='치료' <%If end_type = "치료" then %>selected<% end if %>>치료</option>
                                </select>                                 
                                </td>
                                <td class="left">
								<input name="end_comment" type="text" id="end_comment" style="width:120px" onKeyUp="checklength(this,30)" value="<%=cfm_comment%>">
                                </td>                                
                                <td colspan="2">
                                 <input type="image" name="rptCert$ctl01$btnRequest" id="rptCert_ctl01_btnRequest" src="/image/btn_career_certificate.gif" alt="사직원 신청" onclick="s_sinchung('<%=rs("emp_no")%>','<%=rs("emp_name")%>', 'end_date', 'end_type', 'end_comment');return false;" style="border-width:0px;" />
                                </td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    <% if end_view = "Y" then %>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_family_add.asp?family_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_family_add_pop','scrollbars=yes,width=750,height=400')" class="btnType04">제증명 발급</a>
					<a href="payment_slip_end.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&over_cash=<%=over_cash%>&use_cash=<%=use_cash%>" class="btnType04">전표마감</a>
					<% end if %>
					<% if user_id = "jinhs" then %>
					<a href="payment_slip_end_cancle.asp?from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">마감취소</a>
					<% end if %>
					</div>                  
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="emp_empno" value="<%=in_empno%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

