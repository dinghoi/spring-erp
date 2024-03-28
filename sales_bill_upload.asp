<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

	dim abc,filenm
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

	sales_month = abc("sales_month")
	if sales_month = "" then
		sales_month = mid(now(),1,4) + mid(now(),6,2)
	end if
	from_date = mid(sales_month,1,4) + "-" + mid(sales_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))
	file_type = abc("file_type")

	cost_year = mid(sales_month,1,4)

	if sales_month = "" then
		ck_sw = "y"
	else
	  	ck_sw = "n"
	end if


	Set DbConn = Server.CreateObject("ADODB.Connection")
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set rs_com = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect

	If ck_sw = "n" Then
		Set filenm = abc("att_file")(1)

		path = Server.MapPath ("/large_file")
		filename = filenm.safeFileName
		fileType = mid(filename,inStrRev(filename,".")+1)
		file_name = "사업부별매출"

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
	title_line = "매출 세금계산서 업로드"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리 회계 시스템</title>
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
				return "2 1";
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
//				if (document.frm.bill_id.value == "") {
//					alert ("계산서 유형을 선택하세요");
//					return false;
//				}
				if (document.frm.sales_month.value == "") {
					alert ("년월을 선택하세요");
					return false;
				}
				if (document.frm.att_file.value == "") {
					alert ("업로드 엑셀 파일을 선택하세요");
					return false;
				}
				return true;
			}
			function upload_ok()
				{
				a=confirm('DB에 업로드 하시겠습니까?');
				if (a==true) {
					document.frm.action = "/sales_bill_upload_ok.asp";
					document.frm.submit();
				}
				return false;
				}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/account_cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="sales_bill_upload.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>업로드내용</dt>
                        <dd>
                            <p>
								<label>
								<strong>매출년월 : </strong>
                                	<input name="sales_month" type="text" value="<%=sales_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
                                <label>
								<strong>업로드파일 : </strong>
								<input name="att_file" type="file" id="att_file" size="60" value="<%=att_file%>" style="text-align:left">
								</label>
            					<input name="file_type" type="hidden" id="file_type" value="<%=file_type%>">
            					<a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="4%" >
							<col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="11%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
							<col width="*" >
							<col width="4%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">건수</th>
								<th scope="col">등록</th>
								<th scope="col">매출일</th>
								<th scope="col">매출회사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">상호</th>
								<th scope="col">대표자명</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">품목명</th>
								<th scope="col">담당자</th>
								<th scope="col">관리사업부</th>
								<th scope="col">전표번호</th>
								<th scope="col">입금예정일</th>
							</tr>
						</thead>
						<tbody>
						<%
						tot_price = 0
						tot_cost = 0
						tot_cost_vat = 0
						reg_cnt = 0
						trade_no_err_cnt = 0
						error_cnt = 0

						if rowcount > -1 then
							for i=0 to rowcount
								if xgr(1,i) = "" or isnull(xgr(1,i)) then
									exit for
								end if

								if xgr(0,i) => from_date and xgr(0,i) <= to_date then
									sales_date = xgr(0,i)
									approve_no = xgr(1,i)
									sales_company = xgr(3,i)
									trade_no = xgr(4,i)
									trade_owner = xgr(6,i)
									price = toString(xgr(7,i),"0")
									cost = toString(xgr(8,i),"0")
									cost_vat = toString(xgr(9,i),"0")
									tot_price = tot_price + cdbl(price)
									tot_cost = tot_cost + cdbl(cost)
									tot_cost_vat = tot_cost_vat + cdbl(cost_vat)
									tax_bill_memo = xgr(10,i)
									emp_name = xgr(11,i)
									saupbu = xgr(12,i)
									slip_no = xgr(13,i)
									collect_due_date = xgr(14,i)
									if collect_due_date = "" or isnull(collect_due_date) then
										collect_due_date = ""
									else
									  	collect_due_date = "20" + replace(collect_due_date," . "," ")
									end if
									sql = "select * from saupbu_sales where approve_no = '"&approve_no&"'"
									'Response.write sql
									set rs_etc=dbconn.execute(sql)
									if rs_etc.eof or rs_etc.bof then
										reg_sw = "N"
										'response.write "o"
									Else
									   'response.write "x"
										reg_cnt = reg_cnt + 1
										reg_sw = "Y"
									end if
									rs_etc.close()

									sales_com_err = "N"
									sql = "select * from trade where trade_name = '"&sales_company&"'"
									set rs_trade=dbconn.execute(sql)
									if rs_trade.eof or rs_trade.bof Then
									'Response.write "1"
										error_cnt = error_cnt + 1
										sales_com_err = "Y"
									Else
									'Response.write "2"
										if rs_trade("trade_id") <> "계열사" Then
									'Response.write "3"
											sales_com_err = "Y"
											error_cnt = error_cnt + 1
										end if
									end if
									rs_trade.close()

									emp_name_err = "N"
									if saupbu = "기타사업부" or saupbu = "회사간거래" then
										SQL = "SELECT emp_no FROM emp_master "
										SQL = SQL & "WHERE emp_no = '100359' AND emp_name = '"&emp_name&"' "
										Set rs_emp = dbconn.execute(SQL)

										If rs_emp.EOF Or rs_emp.BOF Then
									'Response.write "4"
											emp_name_err = "Y"
											error_cnt = error_cnt + 1
											emp_no = "error"
										Else
											emp_no = rs_emp("emp_no")
										End If
										rs_emp.close()
									Else
                                        'sql = "select * from emp_master where emp_name = '"&emp_name&"' and emp_saupbu = '"&saupbu&"'"
										SQL = "SELECT emp_no FROM emp_master AS emmt "
										SQL = SQL & "INNER JOIN emp_org_mst AS eomt ON eomt.org_code = emmt.emp_org_code "
										SQL = SQL & "WHERE emmt.emp_name = '"&emp_name&"' AND eomt.org_bonbu = '"&saupbu&"' "
                                        Set rs_emp = dbconn.execute(sql)

										If rs_emp.eof Or rs_emp.bof Then
											emp_name_err = "Y"
									'Response.write "5"

											error_cnt = error_cnt + 1
											emp_no = "error"
										Else
											emp_no = rs_emp("emp_no")
										End If
										rs_emp.close()
									End If

									trade_no = Replace(trade_no,"-","")

									sql = "select trade_name from trade where trade_no = '"&trade_no&"'"
									set rs_trade=dbconn.execute(sql)

									if rs_trade.eof or rs_trade.bof then
										trade_name = xgr(3,i)
									else
										trade_name = rs_trade("trade_name")
									end if
									rs_trade.close()

									cost_sum_err = "N"
									sum_cost = cdbl(cost) + cdbl(cost_vat)
									if sum_cost <> cdbl(price) then
									'Response.write "6"

										error_cnt = error_cnt + 1
										cost_sum_err = "Y"
									end if

									saupbu_err = "N"
									if saupbu = "기타사업부" or saupbu = "회사간거래" then
										saupbu_err = "N"
									else
										sql = "select * from sales_org where saupbu = '"&saupbu&"' and sales_year='" & cost_year & "' order by sort_seq"
										set rs_etc=dbconn.execute(sql)
										if rs_etc.eof or rs_etc.bof then
									'Response.write "7"
											error_cnt = error_cnt + 1
											saupbu_err = "Y"
										end if
										rs_etc.close()
									end if
									slip_no_err = "N"
									collect_due_date_err = "N"
									if collect_due_date <> "" then
										if isdate(collect_due_date) then
											collect_due_date_err = "N"
										Else
									'Response.write "8"

											error_cnt = error_cnt + 1
										  	collect_due_date_err = "Y"
										end if
									end if
									%>
									<tr>
										<td class="first"><%=i+1%></td>
									<% if reg_sw = "N" then %>
										<td>미등록</td>
									<% else	%>
										<td bgcolor="#FFCCFF">등록</td>
									<% end if 	%>
										<td><%=sales_date%></td>
									<% if sales_com_err = "N" then %>
										<td><%=sales_company%></td>
									<% else	%>
										<td bgcolor="#FFCCFF"><%=sales_company%></td>
									<% end if 	%>
										<td><%=trade_no%></td>
										<td><%=trade_name%></td>
										<td><%=trade_owner%></td>
									<% if cost_sum_err = "N" then %>
										<td class="right"><%=formatnumber(price,0)%></td>
									<% else	%>
										<td bgcolor="#FFCCFF" class="right"><%=formatnumber(price,0)%></td>
									<% end if %>
										<td class="right"><%=formatnumber(cost,0)%></td>
										<td class="right"><%=formatnumber(cost_vat,0)%></td>
										<td class="left"><%=tax_bill_memo%></td>
									<% if emp_name_err = "N" then %>
										<td><%=emp_name%></td>
									<% else	%>
										<td bgcolor="#FFCCFF"><%=emp_name%></td>
									<% end if 	%>
									<% if saupbu_err = "N" then %>
										<td><%=saupbu%>&nbsp;</td>
									<% else	%>
										<td bgcolor="#FFCCFF"><%=saupbu%>&nbsp;</td>
									<% end if 	%>
									<% if slip_no_err = "N" then %>
										<td><%=slip_no%>&nbsp;</td>
									<% else	%>
										<td bgcolor="#FFCCFF"><%=slip_no%>&nbsp;</td>
									<% end if 	%>
									<% if collect_due_date_err = "N" then %>
										<td><%=collect_due_date%>&nbsp;</td>
									<% else	%>
										<td bgcolor="#FFCCFF"><%=collect_due_date%>&nbsp;</td>
                                    <% end if %>
									</tr>
						            <%
								end if
							next
						end if
						rec_cnt = i
						%>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>계</strong></td>
								<td class="right"><%=formatnumber(reg_cnt,0)%></td>
								<td><strong>Error</strong></td>
								<td class="right"><%=formatnumber(error_cnt,0)%>건</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(tot_price,0)%></td>
								<td class="right"><%=formatnumber(tot_cost,0)%></td>
								<td class="right"><%=formatnumber(tot_cost_vat,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						</tbody>
					</table>
				</div>
				<%
				If reg_cnt <> rec_cnt  And error_cnt = 0 And rowcount > -1 Then %>
					<br>
                    <div align="center">
                        <span class="btnType01"><input type="button" value="DB에 업로드" onclick="javascript:upload_ok();"NAME="Button1"></span>
                    </div>
				<%End If %>
					<br>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>">
				</form>
		</div>
	</div>
	</body>
</html>

