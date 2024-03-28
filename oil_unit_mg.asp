<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
oil_unit_month = request("oil_unit_month")

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_unit = Server.CreateObject("ADODB.Recordset")					
Set rs_max = Server.CreateObject("ADODB.Recordset")					
DbConn.Open dbconnect

if u_type = "U" then
	sql = "select * from oil_unit where oil_unit_month = '" + oil_unit_month + "'"
	rs_unit.Open Sql, Dbconn, 1
	do until rs_unit.eof
		if rs_unit("oil_unit_id") = "1" and rs_unit("oil_kind") = "휘발유" then
			oil_unit_middle11 = rs_unit("oil_unit_middle")
			oil_unit_last11 = rs_unit("oil_unit_last")
		end if
		if rs_unit("oil_unit_id") = "1" and rs_unit("oil_kind") = "디젤" then
			oil_unit_middle12 = rs_unit("oil_unit_middle")
			oil_unit_last12 = rs_unit("oil_unit_last")
		end if
		if rs_unit("oil_unit_id") = "1" and rs_unit("oil_kind") = "가스" then
			oil_unit_middle13 = rs_unit("oil_unit_middle")
			oil_unit_last13 = rs_unit("oil_unit_last")
		end if
		if rs_unit("oil_unit_id") = "2" and rs_unit("oil_kind") = "휘발유" then
			oil_unit_middle21 = rs_unit("oil_unit_middle")
			oil_unit_last21 = rs_unit("oil_unit_last")
		end if
		if rs_unit("oil_unit_id") = "2" and rs_unit("oil_kind") = "디젤" then
			oil_unit_middle22 = rs_unit("oil_unit_middle")
			oil_unit_last22 = rs_unit("oil_unit_last")
		end if
		if rs_unit("oil_unit_id") = "2" and rs_unit("oil_kind") = "가스" then
			oil_unit_middle23 = rs_unit("oil_unit_middle")
			oil_unit_last23 = rs_unit("oil_unit_last")
		end if
		rs_unit.movenext()
	loop
	rs_unit.close()
  else
	oil_unit_middle11 = 0
	oil_unit_last11 = 0
	oil_unit_middle12 = 0
	oil_unit_last12 = 0
	oil_unit_middle13 = 0
	oil_unit_last13 = 0
	oil_unit_middle21 = 0
	oil_unit_last21 = 0
	oil_unit_middle22 = 0
	oil_unit_last22 = 0
	oil_unit_middle23 = 0
	oil_unit_last23 = 0

	sql="select max(oil_unit_month) as max_month from oil_unit"
	set rs_max=dbconn.execute(sql)
	if isnull(rs_max("max_month")) then
		oil_unit_month = mid(now(),1,4) + mid(now(),6,2)
	  else
		curr_date = mid(rs_max("max_month"),1,4) + "-" + mid(rs_max("max_month"),5,2) + "-01"
		curr_date = datevalue(curr_date)
		next_date = dateadd("m",1,curr_date)
		oil_unit_month = mid(next_date,1,4) + mid(next_date,6,2)	
	end if
end if	

curr_month = mid(now(),1,4) + mid(now(),6,2)

Sql = "SELECT count(*) FROM oil_unit"
Set RsCount = Dbconn.Execute (sql)
tottal_record = cint(RsCount(0)) 'Result.RecordCount
stpage = tottal_record -12
if stpage < 0 then
	stpage = 0
end if
pgsize = 12

sql = "select * from oil_unit order by oil_unit_month, oil_unit_id asc, oil_kind desc limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "월별 유류비 단가 관리"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
				return "7 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {  $( "#datepicker" ).datepicker();
							$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker" ).datepicker("setDate", "<%=holiday%>" );
			});	  
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {

				if(document.frm.oil_unit_month.value > document.frm.curr_month.value) {
					alert('발생일자가 현재일보다 클수가 없습니다.');
					frm.oil_unit_month.focus();
					return false;}

				a=confirm('등록하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_car_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="59%" height="356" valign="top">
					  <form action="holi_del_ok.asp" method="post" name="frm_del">
                      <table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="*" >
				          <col width="17%" >
				          <col width="17%" >
				          <col width="17%" >
				          <col width="17%" >
				          <col width="17%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">년월</th>
				            <th scope="col">구분</th>
				            <th scope="col">유종</th>
				            <th scope="col">월초단가</th>
				            <th scope="col">월말단가</th>
				            <th scope="col">평균단가</th>
			              </tr>
			            </thead>
			            <tbody>
						<%
                        do until rs.eof
							if rs("oil_unit_id") = "1" then
								unit_id_view = "본사팀"
							  else	
								unit_id_view = "지방"
							end if							  	
                            %>
                            <tr>
                            <td class="first"><a href="oil_unit_mg.asp?oil_unit_month=<%=rs("oil_unit_month")%>&u_type=<%="U"%>"><%=rs("oil_unit_month")%></a></td>
                            <td><%=unit_id_view%></td>
                            <td><%=rs("oil_kind")%></td>
                            <td><%=formatnumber(rs("oil_unit_middle"),0)%></td>
                            <td><%=formatnumber(rs("oil_unit_last"),0)%></td>
                            <td><%=formatnumber(rs("oil_unit_average"),0)%></td>
                            </tr>
                            <%
							rs.movenext()
						loop
						%>
			            </tbody>
			          </table>
					  <br>
                      </form>
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="39%" valign="top"><form method="post" name="frm" action="oil_unit_reg_ok.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				        <colgroup>
				          <col width="*" >
				          <col width="30%" >
				          <col width="20%" >
				          <col width="30%" >
			            </colgroup>
				          <tbody>
				            <tr>
				              <th style="background-color:#E8FFFF">년월</th>
				              <td colspan="3" class="left" style="background-color:#E8FFFF"><input name="oil_unit_month" type="text" style="width:70px; text-align:center" readonly="true" value="<%=oil_unit_month%>"></td>
			                </tr>
				            <tr>
				              <th>구분</th>
				              <td class="left">본사팀</td>
				              <th>유종</th>
				              <td class="left">휘발유</td>
			                </tr>
				            <tr>
				              <th>월초단가</th>
				              <td class="left"><input name="oil_unit_middle11" type="text" id="oil_unit_middle11" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_middle11,0)%>" onKeyUp="plusComma(this);" ></td>
				              <th>월말단가</th>
				              <td class="left"><input name="oil_unit_last11" type="text" id="oil_unit_last11" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_last11,0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">구분</th>
				              <td class="left" style="background-color:#E8FFFF">본사팀</td>
				              <th style="background-color:#E8FFFF">유종</th>
				              <td class="left" style="background-color:#E8FFFF">디젤</td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">월초단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_middle12" type="text" id="oil_unit_middle12" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_middle12,0)%>" onKeyUp="plusComma(this);" ></td>
				              <th style="background-color:#E8FFFF">월말단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_last12" type="text" id="oil_unit_last12" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_last12,0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
				            <tr>
				              <th>구분</th>
				              <td class="left">본사팀</td>
				              <th>유종</th>
				              <td class="left">가스</td>
			                </tr>
				            <tr>
				              <th>월초단가</th>
				              <td class="left"><input name="oil_unit_middle13" type="text" id="oil_unit_middle13" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_middle13,0)%>" onKeyUp="plusComma(this);" ></td>
				              <th>월말단가</th>
				              <td class="left"><input name="oil_unit_last13" type="text" id="oil_unit_last13" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_last13,0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">구분</th>
				              <td class="left" style="background-color:#E8FFFF">지방</td>
				              <th style="background-color:#E8FFFF">유종</th>
				              <td class="left" style="background-color:#E8FFFF">휘발유</td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">월초단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_middle21" type="text" id="oil_unit_middle21" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_middle21,0)%>" onKeyUp="plusComma(this);" ></td>
				              <th style="background-color:#E8FFFF">월말단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_last21" type="text" id="oil_unit_last21" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_last21,0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
				            <tr>
				              <th>구분</th>
				              <td class="left">지방</td>
				              <th>유종</th>
				              <td class="left">디젤</td>
			                </tr>
				            <tr>
				              <th>월초단가</th>
				              <td class="left"><input name="oil_unit_middle22" type="text" id="oil_unit_middle22" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_middle22,0)%>" onKeyUp="plusComma(this);" ></td>
				              <th>월말단가</th>
				              <td class="left"><input name="oil_unit_last22" type="text" id="oil_unit_last22" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_last22,0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">구분</th>
				              <td class="left" style="background-color:#E8FFFF">지방</td>
				              <th style="background-color:#E8FFFF">유종</th>
				              <td class="left" style="background-color:#E8FFFF">가스</td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">월초단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_middle23" type="text" id="oil_unit_middle23" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_middle23,0)%>" onKeyUp="plusComma(this);" ></td>
				              <th style="background-color:#E8FFFF">월말단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_last23" type="text" id="oil_unit_last23" style="width:50px;text-align:right" value="<%=formatnumber(oil_unit_last23,0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
			              </tbody>
			            </table>
						<br>
                        <% 
                        if u_type = "U" then
                            u_btn = "변경"
                        else
                            u_btn = "등록"
                        end if
                        %>
				        <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				        <input type="hidden" name="curr_month" value="<%=curr_month%>" ID="Hidden1">
				        <div align=center>
                        	<span class="btnType01"><input type="button" value="<%=u_btn%>" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                        </div>
			          </form>
                      </td>
			        </tr>
			      </table>
                </div>
			</div>				
	</div>        				
	</body>
</html>

