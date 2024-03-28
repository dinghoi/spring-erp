<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

emp_no = request("emp_no")
emp_name = request("emp_name")

in_empno =""
in_name = ""
If Request.Form("in_name")  <> "" Then 
  in_empno = Request.Form("in_empno") 
  in_name = Request.Form("in_name") 
End If

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

Sql = "SELECT count(*) FROM emp_appoint where (app_empno = '"+emp_no+"') and (app_id = '포상발령' or app_id = '징계발령')"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

if tottal_record = 0 then
   first_view = "Y"
end if

Sql = "SELECT * FROM emp_appoint where (app_empno = '"+emp_no+"') and (app_id = '포상발령' or app_id = '징계발령') ORDER BY app_empno,app_date,app_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "◈ 상벌 사항 ◈"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>상벌 사항</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function chkfrm() {
				if(document.frm.in_name.value =="") {
					alert('성명을 입력하세요');
					frm.in_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_reward_punish_view.asp?emp_no=<%=emp_no%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>사번 : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=emp_no%>" style="width:60px; text-align:left" readonly="true">
								</label>
                            <strong>성명 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=emp_name%>" style="width:100px; text-align:left" readonly="true">
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="12%" >
							<col width="17%" >
                            <col width="*" >
                            <col width="33%" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">상벌일자</th>
                                <th>상벌유형</th>
                                <th>징계기간</th>
                                <th>상벌내용</th>
                                <th>직급/직책 및 소속</th>
 							</tr>
						</thead>
						<tbody>
						<%
						     v_cnt = 0
							do until rs.eof or rs.bof
							 v_cnt = v_cnt + 1
							 
							 'task_memo = replace(rs("career_task"),chr(34),chr(39))
							 'view_memo = task_memo
							 'if len(task_memo) > 10 then
							 ' 	view_memo = mid(task_memo,1,10) + ".."
							 'end if
							 
						%>
							<tr>
							  <td><%=rs("app_date")%>&nbsp;</td>
                        <% if rs("app_id") = "포상발령" then %>
						      <td class="left">(포상)<%=rs("app_id_type")%>&nbsp;</td>
                              <td class="left">&nbsp;</td> 
                              <td class="left"><%=rs("app_reward")%>&nbsp;</td> 
                        <%    elseif rs("app_id") = "징계발령" then %>
                              <td class="left">(징계)<%=rs("app_id_type")%>&nbsp;</td>
                              <td class="left"><%=rs("app_start_date")%>∼<%=rs("app_finish_date")%>&nbsp;</td>
                              <td class="left"><%=rs("app_comment")%>&nbsp;</td>
                        <% end if %>
                              <td class="left"><%=rs("app_to_grade")%>-<%=rs("app_to_position")%>(<%=rs("app_to_company")%>&nbsp;<%=rs("app_to_org")%>(<%=rs("app_to_orgcode")%>)</td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
						<%
						  if v_cnt = 0 then
						%>
							<tr>
								<td class="first" colspan="5">내역이 없습니다</td>
							</tr>
                        <%
						  end if
						%>
						</tbody>
					</table>
				</div>
			</div>				
	   </div>    
                   	<br>
               		<div align=right>
						<a href="#" class="btnType04" onclick="javascript:goAction()" >닫기</a>&nbsp;&nbsp;
					</div>
                    <br>    				
	</form>
	</body>
</html>

