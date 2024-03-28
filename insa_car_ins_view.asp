<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

car_no = request("car_no")
car_name = request("car_name")
car_year = request("car_year")
car_reg_date = request("car_reg_date")
u_type = request("u_type")

If Request.Form("in_carno")  <> "" Then 
  car_no = Request.Form("in_carno") 
  car_name = Request.Form("in_name") 
  car_year = Request.Form("car_year")
  car_reg_date = Request.Form("car_reg_date")
End If

If ck_sw = "y" Then
	car_no = request("car_no")
    car_name = request("car_name")
    car_year = request("car_year")
    car_reg_date = request("car_reg_date")
'  else
'	car_no = Request.form("in_carno")
'    car_name = Request.form("in_name")
'    car_year = Request.form("in_year")
'    car_reg_date = Request.form("car_reg_date")
End if

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

Sql = "SELECT count(*) FROM car_insurance where ins_car_no = '"&car_no&"'"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from car_insurance where ins_car_no = '" + car_no + "' ORDER BY ins_car_no,ins_date DESC limit "& stpage & "," &pgsize 

Rs.Open Sql, Dbconn, 1

title_line = " 차량 보험가입 현황 "

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
					alert('차량명을 입력하세요');
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
				<form action="insa_car_ins_view.asp?car_no=<%=car_no%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>차량번호 : </strong>
								<label>
        						<input name="in_carno" type="text" id="in_carno" value="<%=car_no%>" style="width:100px; text-align:left" readonly="true">
								</label>
                            <strong>차종/연식/취득일 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=car_name%>" style="width:100px; text-align:left" readonly="true">
                                -
                                <input name="in_year" type="text" id="in_year" value="<%=car_year%>" style="width:100px; text-align:left" readonly="true">
                                 -
                                <input name="car_reg_date" type="text" id="car_reg_date" value="<%=car_reg_date%>" style="width:100px; text-align:left" readonly="true">
								</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="10%" >
                            <col width="6%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="4%" >
                            <col width="*" >
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">가입일</th>
                                <th scope="col">보험사</th>
                                <th scope="col">보험기간</th>
                                <th scope="col">보험료</th>
                                <th scope="col">대인1</th>
                                <th scope="col">대인2</th>
                                <th scope="col">대물</th>
                                <th scope="col">자기보험</th>
                                <th scope="col">무상해</th>
                                <th scope="col">자차</th>
                                <th scope="col">연령</th>
                                <th scope="col">긴급<br>출동</th>
                                <th scope="col">계약내용</th>
 							</tr>
						</thead>
						<tbody>
						<%
							do until rs.eof or rs.bof
						%>
							<tr>
								<td><%=rs("ins_date")%>&nbsp;</td>
								<td><%=rs("ins_company")%>&nbsp;</td>
                                <td><%=rs("ins_last_date")%>&nbsp;</td>
                                <td><%=formatnumber(rs("ins_amount"),0)%>&nbsp;</td>
                                <td><%=rs("ins_man1")%>&nbsp;</td>
                                <td><%=rs("ins_man2")%>&nbsp;</td>
                                <td><%=rs("ins_object")%>&nbsp;</td>
                                <td><%=rs("ins_self")%>&nbsp;</td>
                                <td><%=rs("ins_injury")%>&nbsp;</td>
                                <td><%=rs("ins_self_car")%>&nbsp;</td>
                                <td><%=rs("ins_age")%>&nbsp;</td>
                                <td><%=rs("ins_scramble")%>&nbsp;</td>
                       <% if rs("ins_contract_yn") = "Y" then %>
                                <td>계약내용포함&nbsp;</td>
                         <%    else %>
                                <td>계약내용미포함(<%=rs("ins_comment")%>)&nbsp;</td>
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
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1
                
                if intend > total_page then
                    intend = total_page
                end if
                %>     
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    <div id="paging">
                        <a href="insa_car_ins_view.asp?page=<%=first_page%>&car_no=<%=car_no%>&car_name=<%=car_name%>&car_year=<%=car_year%>&car_reg_date=<%=car_reg_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_car_ins_view.asp?page=<%=intstart -1%>&car_no=<%=car_no%>&car_name=<%=car_name%>&car_year=<%=car_year%>&car_reg_date=<%=car_reg_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_car_ins_view.asp?page=<%=i%>&car_no=<%=car_no%>&car_name=<%=car_name%>&car_year=<%=car_year%>&car_reg_date=<%=car_reg_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_car_ins_view.asp?page=<%=intend+1%>&car_no=<%=car_no%>&car_name=<%=car_name%>&car_year=<%=car_year%>&car_reg_date=<%=car_reg_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_car_ins_view.asp?page=<%=total_page%>&car_no=<%=car_no%>&car_name=<%=car_name%>&car_year=<%=car_year%>&car_reg_date=<%=car_reg_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div align=right>
						<a href="#" class="btnType04" onclick="javascript:goAction()" >닫기</a>&nbsp;&nbsp;
					</div>              
                    </td>
			      </tr>
			  </table>
         </div>	
	</form>
	  </div>				
	</body>
</html>

