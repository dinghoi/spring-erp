<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

curr_date = mid(cstr(now()),1,10)

emp_no = request("emp_no")
be_pg = request("be_pg")
page = request("page")
page_cnt = request("page_cnt")

be_pg1 = "insa_gun_month_list.asp"
'be_pg = "insa_gun_month_list.asp"

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


title_line = " 월별 근태현황(공사중)....."

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
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
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
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
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_gun_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_gun_month_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>연차휴가 일수</dt>
                        <dd>
							<strong>년도: </strong>
								<label>
        						<input name="in_year" type="text" id="in_year" value="<%=year_year%>" readonly="true" style="width:40px; text-align:left">
								</label>
                            <strong>년차기산일: </strong>
                                <label>
                               	<input name="in_yuncha_date" type="text" id="in_yuncha_date" value="<%=year_yuncha_date%>" readonly="true" style="width:60px; text-align:left">
								</label>
                            <strong>근속년수: </strong>
                                <label>
                               	<input name="in_continu_year" type="text" id="in_continu_year" value="<%=year_continu_year%>" readonly="true" style="width:40px; text-align:left">
                                -
                                <input name="in_continu_month" type="text" id="in_continu_month" value="<%=year_continu_month%>" readonly="true" style="width:40px; text-align:left">
								</label>
                            <strong>발생연차: </strong>
                                <label>
                               	<input name="in_basic_count" type="text" id="in_basic_count" value="<%=year_basic_count%>" readonly="true" style="width:40px; text-align:left">
                                -
                                <input name="in_add_count" type="text" id="in_add_count" value="<%=year_add_count%>" readonly="true" style="width:40px; text-align:left">
								</label>
                            <strong>사용연차: </strong>
                                <label>
                               	<input name="in_use_count" type="text" id="in_use_count" value="<%=year_use_count%>" readonly="true" style="width:40px; text-align:left">
								</label>
                            <strong>잔여연차: </strong>
                                <label>
                               	<input name="in_remain_count" type="text" id="in_remain_count" value="<%=year_remain_count%>" readonly="true" style="width:40px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
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
							<col width="8%" >
                            <col width="8%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="*%" >
						</colgroup>
						<thead>
						    <tr>
				                <th rowspan="2" class="first" scope="col" style=" border-left:1px solid #e3e3e3;">년월</th>
                                <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">휴&nbsp;&nbsp;&nbsp;가</th>
				                <th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3;">근&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;태</th>
			                </tr>
                            <tr>
								<th class="first" scope="col" style=" border-right:1px solid #e3e3e3;">연차</th>
								<th scope="col">반차</th>
								<th scope="col">대휴</th>
								<th scope="col">공가</th>
								<th scope="col">정기휴가</th>
                                <th scope="col">시간외근무</th>
                                <th scope="col">휴일근무</th>
                                <th scope="col">외근</th>
                                <th scope="col">출장</th>
                                <th scope="col">조퇴</th>
                                <th scope="col">결근</th>
                                <th scope="col">기타</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

	           			%>
							<tr>
                                <td class="first">
                                <a href="#" onClick="pop_Window('insa_gun_month_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&page_cnt=<%=page_cnt%>','gun_monthview','scrollbars=yes,width=800,height=400')"><%=rs("emp_no")%></a>
								</td>
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
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
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
                       <div class="btnRight">
                         <span class="btnType01"><input type="button" value="이전" onclick="javascript:goBefore();"></span>
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

