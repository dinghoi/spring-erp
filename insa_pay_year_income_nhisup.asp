<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

dim abc,filenm
dim year_tab(3,2)

Set abc = Server.CreateObject("ABCUpload4.XForm")
abc.AbsolutePath = True
abc.Overwrite = true
abc.MaxUploadSize = 1024*1024*50

pay_company = abc("pay_company")
pay_yyyy = abc("pay_yyyy")
give_date = abc("give_date")
file_type = abc("file_type")

if pay_company = "" then
	ck_sw = "y"
  else
  	ck_sw = "n"
end if
	
if pay_company = "" then
    pay_company = "전체"
    curr_dd = cstr(datepart("d",now))
    give_date = mid(cstr(now()),1,10)
    from_date = mid(cstr(now()-curr_dd+1),1,10)
    pay_month = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
end if
	
' 최근3개년도 테이블로 생성
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"
	
	Set DbConn = Server.CreateObject("ADODB.Connection")
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")	
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set Rs_org = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set rs_com = Server.CreateObject("ADODB.Recordset")
	Set Rs_year = Server.CreateObject("ADODB.Recordset")
	Set Rs_ins = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect

	If ck_sw = "n" Then
		Set filenm = abc("att_file")(1)
		
		path = Server.MapPath ("/pay_file")
		filename = filenm.safeFileName
		fileType = mid(filename,inStrRev(filename,".")+1)
		file_name = pay_company + "_" + pay_yyyy + "_건강보험 표준월액" + give_date
		
'		save_path = path & "\" & filename
		save_path = path & "\" & file_name&"."&fileType

		if fileType = "xls" or fileType = "xlk" then
			file_type = "Y"
			filenm.save save_path
		
			objFile = save_path
	'		objFile = Request.form("att_file")
	'		objFile = SERVER.MapPath("att_file")
	'		objFile = SERVER.MapPath(".") & "\kwon_upload\excel_data.xls"
	'		response.write(objFile)
			
			cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
			rs.Open "select * from [1:10000]",cn,"0"
				
			rowcount=-1
			xgr = rs.getrows
			rowcount = ubound(xgr,2)
			fldcount = rs.fields.count
			tot_cnt = rowcount + 1
		  else
			objFile = "none"
			rowcount=-1
			file_type = "N"
		end if		  
	  else
		objFile = "none"
		rowcount=-1
	end if
	title_line = "건강보험 표준월액 자료 업로드"
	
incom_year = pay_yyyy

'국민연금 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&pay_yyyy&"' and insu_id = '5501' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	nps_emp = formatnumber(rs_ins("emp_rate"),3)
		nps_com = formatnumber(rs_ins("com_rate"),3)
		nps_from = rs_ins("from_amt")
		nps_to = rs_ins("to_amt")
   else
		nps_emp = 0
		nps_com = 0
		nps_from = 0
		nps_to = 0
end if
rs_ins.close()

'건강보험 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&pay_yyyy&"' and insu_id = '5502' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	nhis_emp = formatnumber(rs_ins("emp_rate"),3)
		nhis_com = formatnumber(rs_ins("com_rate"),3)
		nhis_from = rs_ins("from_amt")
		nhis_to = rs_ins("to_amt")
   else
		nhis_emp = 0  
		nhis_com = 0
		nhis_from = 0
		his_to = 0
end if
rs_ins.close()	
	
	

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "3 1";
			}
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=give_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=end_date%>" );
			});	  
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.pay_company.value == "") {
					alert ("회사를 선택하세요");
					return false;
				}	
				if (document.frm.pay_yyyy.value == "") {
					alert ("귀속년월을 선택하세요");
					return false;
				}	
				if (document.frm.att_file.value == "") {
					alert ("업로드 엑셀 파일을 선택하세요");
					return false;
				}	
				return true;
			}
			function frm1check () {
				if (chkfrm1()) {
					document.frm1.submit ();
				}
			}
			
			function chkfrm1() {
				a=confirm('DB에 업로드 하시겠습니까?');
				if (a==true) {
					return true;
				}
				return false;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_income_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_year_income_nhisup.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>업로드내용</dt>
                        <dd>
                            <p>
								<label>
								<strong>회사: </strong>
                              <%
								Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = '회사' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
								<select name="pay_company" id="pay_company" type="text" style="width:110px">
                                    <option value="전체" <%If pay_company = "전체" then %>selected<% end if %>>전체</option>
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If pay_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>
                                </label>
								<label>
								<strong>귀속년도 : </strong>
                                <select name="pay_yyyy" id="pay_yyyy" type="text" value="<%=pay_yyyy%>" style="width:90px">
                                    <%	for i = 3 to 1 step -1	%>
                                    <option value="<%=year_tab(i,1)%>" <%If pay_yyyy = cstr(year_tab(i,1)) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                                    <%	next	%>
                                </select>
								</label>
                                <br>
                                <label>
								<strong>업로드파일: </strong>
								<input name="att_file" type="file" id="att_file" size="100" value="<%=att_file%>" style="text-align:left"> 
								</label>
            					<input name="file_type" type="hidden" id="file_type" value="<%=file_type%>">
            					<a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="3%" >
                            <col width="5%" >
							<col width="*" >
                            
							<col width="8%" >
                            <col width="8%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
                            <col width="8%" >
                            
							<col width="7%" >
							<col width="6%" >
                            
							<col width="2%" >
                            <col width="2%" >
                            <col width="2%" >
                            <col width="2%" >
                            <col width="2%" >
                            <col width="2%" >
                            <col width="2%" >
                            <col width="2%" >
                            <col width="2%" >
                            <col width="2%" >
                            <col width="2%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">건수</th>
								<th scope="col">등록</th>
                                <th scope="col">사번</th>
								<th scope="col">성명</th>
                                
								<th scope="col">연봉</th>
                                <th scope="col">기본급</th>
								<th scope="col">연장수당</th>
								<th scope="col">식대</th>
								<th scope="col">퇴직금</th>
								<th scope="col">평균<br>소득월액</th>
                                
                                <th scope="col">건강보험<br>표준월액</th>
								<th scope="col">건강보험<br>월납부액</th>
                                
								<th scope="col">고용<br>보험</th>
								<th scope="col">산재<br>보험</th>
                                <th scope="col">장기<br>요양</th>
                                <th scope="col">청년<br>감면</th>
                                <th scope="col">부양<br>가족수</th>
                                <th scope="col">배우자</th>
                                <th scope="col">20세<br>이하</th>
                                <th scope="col">60세<br>이상</th>
                                <th scope="col">경로<br>우대</th>
                                <th scope="col">장애<br>인</th>
                                <th scope="col">부녀<br>자</th>
							</tr>
						</thead>
						<tbody>
				  <%
						tot_emp = 0
						tot_name = 0
						tot_bank = 0
						tot_err = 0
						
						tot_total_pay = 0
	                    tot_base_pay = 0
	                    tot_overtime_pay = 0
	                    tot_meals_pay = 0
	                    tot_severance_pay = 0
	                    tot_month_amount = 0
	                    tot_nps_amount = 0
	                    tot_nps = 0
	                    tot_nhis_amount = 0
	                    tot_nhis = 0
	                    tot_go_yn = 0
	                    tot_san_yn = 0
	                    tot_long_yn = 0
	                    tot_incom_yn = 0
	                    tot_family_cnt = 0
	                    tot_wife_yn = 0
	                    tot_age20 = 0
	                    tot_age60 = 0
	                    tot_old = 0
	                    tot_disab = 0
	                    tot_woman = 0
	                    tot_retirement_bank = 0	
												
						if rowcount > -1 then
							for i=0 to rowcount
							if xgr(1,i) = "" or isnull(xgr(1,i)) then
								exit for
							end if
					' 사번체크 				
							emp_sw = "Y"
							emp_no = xgr(1,i)
							Sql = "select * from emp_master where emp_no = '"&xgr(1,i)&"'"
							Set rs_emp = DbConn.Execute(SQL)
							if rs_emp.eof then
								tot_emp = tot_emp + 1
								tot_err = tot_err + 1
								emp_sw = "N"
								emp_name =""
							  else
								emp_name = rs_emp("emp_name")	  
							end if
							name_sw = "Y"
							if xgr(2,i) <> emp_name then
							    tot_name = tot_name + 1
								tot_err = tot_err + 1
								name_sw = "N"
								emp_name = xgr(2,i)	
							end if
							
							incom_total_pay = xgr(18,i)
					'		incom_base_pay = xgr(4,i)
					'		incom_overtime_pay = xgr(5,i)
					'		incom_meals_pay = xgr(6,i)
					'		incom_severance_pay = xgr(7,i)
					'		incom_month_amount = xgr(8,i)
					 ' 기본급등 계산		
							mon13_pay = int(incom_total_pay / 13)
							meals_pay = 100000
							ot_pay = int((mon13_pay - meals_pay) * 0.09)
							base_pay = int(mon13_pay - meals_pay - ot_pay)
							mon_amt = base_pay + ot_pay
							se_pay = int(mon13_pay)
							
							incom_base_pay = base_pay
							incom_overtime_pay = ot_pay
							incom_meals_pay = meals_pay
							incom_severance_pay = se_pay
							incom_month_amount = mon_amt
					 		
                     '건강보험 계산
					        incom_nhis_amount = xgr(26,i)
					'		incom_nhis = xgr(12,i)
                            nhis_amt = incom_nhis_amount * (nhis_emp / 100)
                            nhis_amt = int(nhis_amt)
                            incom_nhis = (int(nhis_amt / 10)) * 10

					' 항목
	                        incom_go_yn = xgr(28,i)
	                        incom_san_yn = xgr(29,i)
	                        incom_long_yn = xgr(30,i)
	                        incom_incom_yn = xgr(31,i)
	                        incom_wife_yn = xgr(32,i)
	                        incom_age20 = xgr(33,i)
	                        incom_age60 = xgr(34,i)
	                        incom_old = xgr(35,i)
	                        incom_disab = xgr(36,i)
	                        incom_woman = xgr(37,i)
							incom_family_cnt = xgr(38,i)
'	                        incom_retirement_bank = xgr(24,i)

							Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&pay_yyyy&"'"
							set Rs_year=dbconn.execute(sql)				
							if Rs_year.eof or Rs_year.bof then
								reg_sw = "N"
							  else
								reg_sw = "Y"
							end if
						    Rs_year.close()
							
						    tot_total_pay = tot_total_pay + incom_total_pay
							
					%>
							<tr>
								<td class="first"><%=i+1%></td>
							<% if reg_sw = "N" then %>
								<td>No</td>
                            <%   else	%>
								<td bgcolor="#FFCCFF">Yes</td>
							<% end if 	%>                                
							<% if emp_sw = "Y" then %>
								<td><%=emp_no%></td>
                            <%   else	%>
								<td bgcolor="#FFCCFF"><%=emp_no%></td>
							<% end if 	%>                                
							<% if name_sw = "Y" then %>
								<td><%=emp_name%></td>
                            <%   else	%>
								<td bgcolor="#FFCCFF"><%=emp_name%></td>
							<% end if 	%>                                
								<td class="right"><%=formatnumber(incom_total_pay,0)%></td>
								<td class="right"><%=formatnumber(incom_base_pay,0)%></td>
								<td class="right"><%=formatnumber(incom_overtime_pay,0)%></td>
                                <td class="right"><%=formatnumber(incom_meals_pay,0)%></td>
                                <td class="right"><%=formatnumber(incom_severance_pay,0)%></td>
                                <td class="right"><%=formatnumber(incom_month_amount,0)%></td>

                                <td class="right"><%=formatnumber(incom_nhis_amount,0)%></td>
                                <td class="right"><%=formatnumber(incom_nhis,0)%></td>
                                
                                <td class="center"><%=incom_go_yn%></td>
                                <td class="center"><%=incom_san_yn%></td>
                                <td class="center"><%=incom_long_yn%></td>
                                <td class="center"><%=incom_incom_yn%></td>
                                <td class="center"><%=incom_family_cnt%></td>
                                <td class="center"><%=incom_wife_yn%></td>
                                <td class="center"><%=incom_age20%></td>
                                <td class="center"><%=incom_age60%></td>
                                <td class="center"><%=incom_old%></td>
                                <td class="center"><%=incom_disab%></td>
                                <td class="center"><%=incom_woman%></td>
							</tr>
						<%
							next
						end if
						%>
							<tr>
								<th class="first">오류</th>
                                <th><%=formatnumber(tot_bank,0)%></th>
								<th><%=formatnumber(tot_emp,0)%></th>
								<th><%=formatnumber(tot_name,0)%></th>
								<th class="right">&nbsp;</th>
								<th class="right">&nbsp;</th>
								<th class="right">&nbsp;</th>
                                <th class="right">&nbsp;</th>
                                <th class="right">&nbsp;</th>
                                <th class="right">&nbsp;</th>
                                <th class="right">&nbsp;</th>
                                <th class="right">&nbsp;</th>
                                <th class="right">&nbsp;</th>
                                <th class="right">&nbsp;</th>
                                <th class="right">&nbsp;</th>
                                <th class="right">&nbsp;</th>
                                
                                <th class="center">&nbsp;</th>
                                <th class="center">&nbsp;</th>
                                <th class="center">&nbsp;</th>
                                <th class="center">&nbsp;</th>
                                <th class="center">&nbsp;</th>
                                <th class="center">&nbsp;</th>
                                <th class="center">&nbsp;</th>
							</tr>
						</tbody>
					</table>
				</div>
                    <input type="hidden" name="nps_emp" value="<%=formatnumber(nps_emp,3)%>" ID="Hidden1">
                    <input type="hidden" name="nps_com" value="<%=formatnumber(nps_com,3)%>" ID="Hidden1">
                    <input type="hidden" name="nhis_emp" value="<%=formatnumber(nhis_emp,3)%>" ID="Hidden1">
                    <input type="hidden" name="nhis_com" value="<%=formatnumber(nhis_com,3)%>" ID="Hidden1">
                    <input type="hidden" name="nps_from" value="<%=nps_from%>" ID="Hidden1">
                    <input type="hidden" name="nps_to" value="<%=nps_to%>" ID="Hidden1">
                    <input type="hidden" name="nhis_from" value="<%=nhis_from%>" ID="Hidden1">
                    <input type="hidden" name="nhis_to" value="<%=nhis_to%>" ID="Hidden1">
				</form>
			<% 
			      if tot_cnt <> 0 and tot_err = 0 then %>
				<form action="insa_pay_year_income_nhisup_ok.asp" method="post" name="frm1">
					<br>
                    <div align=center>
                        <span class="btnType01"><input type="button" value="DB UPload 저장" onclick="javascript:frm1check();"NAME="Button1"></span>
                    </div>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>">
                    <input name="incom_year" type="hidden" id="incom_year" value="<%=pay_yyyy%>">
                    <input name="incom_company" type="hidden" id="incom_company" value="<%=pay_company%>">
					<br>
				</form>
			<%    
			   end if %>
		</div>				
	</div>        				
	</body>
</html>

