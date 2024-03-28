<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

	dim abc,filenm
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

	bill_id = abc("bill_id")
	bill_month = abc("bill_month")
	if bill_month = "" then
		bill_month = mid(now(),1,4) + mid(now(),6,2)
	end if
	from_date = mid(bill_month,1,4) + "-" + mid(bill_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))
	file_type = abc("file_type")

	if bill_id = "" then
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
		file_name = "e세로"
		
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

	title_line = "이세로 세금계산서 업로드"

' 영업전표 완료전까지 고정	
	bill_id = "1"

	' 2019.02.15 박성민 요청 19년 부터 X,Y 컬럼에 각각 '수탁사업자등록번호','상호' 가 추가되었음
	' 원칙적으로 프로그램을 수정해야 하나 박성민본인이 이 두 컬럼을 삭제하고 업로드하겠다고 함.. 
	' 이유을 물어보니 다른곳(다른엑셀)에선 	잘된다고 함.. (엑셀이 다 따로 노는건지 의심..)
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
				return "1 1";
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.bill_id.value == "") {
					alert ("계산서 유형을 선택하세요");
					return false;
				}	
				if (document.frm.bill_month.value == "") {
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
					document.frm.action = "tax_bill_upload_ok.asp";
					document.frm.submit();
				}
				return false;
				}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/tax_bill_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="tax_bill_upload.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>업로드내용</dt>
            <dd>
            <p>
								<label>
									<strong>계산서 유형 : </strong>
									<input type="radio" name="bill_id" value="1" <% if bill_id = "1" then %>checked<% end if %> style="width:25px">매입
									<input type="radio" name="bill_id" value="2" <% if bill_id = "2" then %>checked<% end if %> style="width:25px">매출
								</label>
								<label>
									<strong>계산서 발행년월 : </strong>
									<input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
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
							<col width="10%" >
							<col width="7%" >
							<col width="11%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
							<col width="3%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">건수</th>
								<th scope="col">등록</th>
								<th scope="col">발행일</th>
								<th scope="col">계산서소유회사</th>
								<th scope="col">사업자번호</th>
								<th scope="col">상호</th>
								<th scope="col">대표자명</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">청구</th>
								<th scope="col">계산서이메일</th>
								<th scope="col">품목명</th>
							</tr>
						</thead>
						<tbody>
						<%
						tot_price = 0
						tot_cost = 0
						tot_cost_vat = 0						   
						reg_cnt = 0
						trade_no_err_cnt = 0
												
						if rowcount > -1 then
							for i=0 to rowcount
								if xgr(1,i) = "" or isnull(xgr(1,i)) then
									exit for
								end if
	
								if xgr(0,i) => from_date and xgr(0,i) <= to_date then
									bill_date = xgr(0,i)
									approve_no = xgr(1,i)
									if bill_id = "1" then
										trade_no = xgr(4,i)
										trade_name = xgr(6,i)
										trade_owner = xgr(7,i)										
										owner_trade_no = xgr(8,i)
									  else
										trade_no = xgr(8,i)
										trade_name = xgr(10,i)
										trade_owner = xgr(11,i)
										owner_trade_no = xgr(4,i)
									end if
									price = xgr(12,i)
									tot_price = tot_price + price
									cost = xgr(13,i)
									tot_cost = tot_cost + cost
									cost_vat = xgr(14,i)
									tot_cost_vat = tot_cost_vat + cost_vat
									bill_collect = xgr(19,i)
									send_email = xgr(20,i)
									receive_email = xgr(21,i)
									tax_bill_memo = xgr(24,i)

									if bill_id = "1" then
										email_view = send_email
									  else
									  	email_view = receive_email
									end if
									
									sql = "select * from tax_bill where approve_no = '"&approve_no&"'"
									set rs_etc=dbconn.execute(sql)				
									if rs_etc.eof or rs_etc.bof then
										reg_sw = "N"
									  else
										reg_cnt = reg_cnt + 1
										reg_sw = "Y"
									end if
									rs_etc.close()					

									owner_trade_no = Replace(owner_trade_no,"-","")
									sql = "select * from trade where trade_no = '"&owner_trade_no&"'"
									set rs_trade=dbconn.execute(sql)				
									if rs_trade.eof or rs_trade.bof then
										owner_sw = "N"
										owner_cnt = owner_cnt + 1
										owner_company = owner_trade_no + "_Error"
									  else
										owner_sw = "Y"
										owner_company = rs_trade("trade_name")
									end if
									rs_trade.close()					

									trade_no_err = "N"
'									if trade_no = "107-81-54150" then
'										trade_no_err_cnt = trade_no_err_cnt + 1
'										trade_no_err = "Y"														
'									end if  
									%>
									<tr>
										<td class="first"><%=i+1%></td>
									<% if reg_sw = "N" then %>
										<td>미등록</td>
									<%   else	%>
										<td bgcolor="#FFCCFF">등록</td>
									<% end if 	%>                                
										<td><%=bill_date%></td>
									<% if owner_sw = "Y" then %>
										<td><%=owner_company%></td>
									<%   else	%>
										<td bgcolor="#FFCCFF"><%=owner_company%></td>
									<% end if 	%>                                
									<% if trade_no_err = "N" then %>
										<td><%=trade_no%></td>
									<%   else	%>
										<td bgcolor="#FFCCFF"><%=trade_no%></td>
									<% end if 	%>                                
										<td><%=trade_name%></td>
										<td><%=trade_owner%></td>
										<td class="right"><%=formatnumber(price,0)%></td>
										<td class="right"><%=formatnumber(cost,0)%></td>
										<td class="right"><%=formatnumber(cost_vat,0)%></td>
										<td><%=bill_collect%></td>
										<td><%=email_view%>&nbsp;</td>
										<td class="left"><%=tax_bill_memo%></td>
									<%
								end if
							next
						end if
						rec_cnt = i
						%>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>계</strong></td>
								<td class="right"><%=formatnumber(reg_cnt,0)%></td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(owner_cnt,0)%></td>
								<td class="right"><%=formatnumber(trade_no_err_cnt,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(tot_price,0)%></td>
								<td class="right"><%=formatnumber(tot_cost,0)%></td>
								<td class="right"><%=formatnumber(tot_cost_vat,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						</tbody>
					</table>
				</div>
				<% if reg_cnt <> rec_cnt  and owner_cnt = 0 and trade_no_err_cnt = 0 and rowcount > -1 then %>
					<br>
                    <div align=center>
                        <span class="btnType01"><input type="button" value="DB에 업로드" onclick="javascript:upload_ok();"NAME="Button1"></span>
                    </div>
				<% end if %>
					<br>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>">
				</form>
		</div>				
	</div>        				
	</body>
</html>
