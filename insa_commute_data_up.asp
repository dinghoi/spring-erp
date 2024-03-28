<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

	dim abc,filenm
	Set abc = Server.CreateObject("ABCUpload4.XForm")
	abc.AbsolutePath = True
	abc.Overwrite = true
	abc.MaxUploadSize = 1024*1024*50

	tot_cnt = 0
	tot_err = 0
	tot_dept = 0
	tot_cust = 0
	tot_ddd = 0
	tot_tel = 0
	tot_sido = 0
	tot_gugun = 0
	tot_dong = 0
	tot_addr = 0
	tot_ce = 0
	tot_cnt = 0

	Set DbConn = Server.CreateObject("ADODB.Connection")
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")	
	Set Rs_etc = Server.CreateObject("ADODB.Recordset")
	Set rs_com = Server.CreateObject("ADODB.Recordset")
	DbConn.Open dbconnect

		Set filenm = abc("att_file")(1)
		
		att_file = abc("att_file")
		
		path = Server.MapPath ("/large_file")

		filename = filenm.safeFileName

		'Response.write filename
		
		fileType = mid(filename,inStrRev(filename,".")+1)
		
'		save_path = path & "\" & file_name&"."&fileType
		save_path = path & "\" & filename

		if fileType = "xls" or fileType = "xlk" then
			file_type = "Y"
			filenm.save save_path
		
			objFile = save_path

		'Response.write objFile
							
			cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
			rs.Open "select * from [1:10000]",cn,"0"

			rowcount=-1
			xgr = rs.getrows
			rowcount = ubound(xgr,2)
			
			fldcount = rs.fields.count
			tot_cnt = rowcount + 1
			
			'Response.write tot_cnt
							
							
		  else
			objFile = "none"
			rowcount=-1
			file_type = "N"
		end if		  
	title_line = "출근 파일  엑셀 업로드"
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
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "3 1";
			}
			function frmcheck () {
				//if (formcheck(document.frm) && chkfrm()) {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				//alert(document.frm.att_file.value);
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
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_gun_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="insa_commute_data_up.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
                                <label>
								<strong>업로드파일 : </strong>
								<input name="att_file" type="file" id="att_file" size="113" value="<%=att_file%>" style="text-align:left"> 
								</label>
								<input name="file_type" type="hidden" id="file_type" value="<%=file_type%>">
                <!--a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="내용보기"></a -->
                  <span class="btnType01"><input type="button" value="업로드" onclick="javascript:frmcheck();"NAME="Button1"></span>
                </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<!--col width="10%" -->
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" scope="col">사번</th>
								<th rowspan="2" scope="col">출근일자</th>
								<th rowspan="2" scope="col">근무형태</th>
								<th rowspan="2" scope="col">출근시간</th>
								<!--th rowspan="2" scope="col">기타</th-->
							</tr>
						</thead>
						<tbody>
						<%
						
						'Response.write rowcount
						  if rowcount > -1 then
							for i=0 to rowcount %>
							<tr>
							<%
							 for j=0 to fldcount-1
							   'Response.write xgr(j,i)
							 %>
							   <td><%=xgr(j,i)%></td>
              <% next %>
                <td><%=mid(xgr(2,i), 1, 5)%></td>
                <!--td> </td-->           
                 <tr>
						<%	next
						end if
						%>
						</tbody>
					</table>
				</div>
				</form>
				<!--% if tot_err = 0 and tot_cnt <> 0 then % -->
				<form action="insa_commute_data_up_ok.asp" method="post" name="frm1">
					<br>
                    <div align=center>
                        <span class="btnType01"><input type="button" value="DB저장" onclick="javascript:frm1check();"NAME="Button1"></span>
                    </div>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>">
					<br>
				</form>
				<!--% end if %-->
		</div>				
	</div>        				
	</body>
</html>

