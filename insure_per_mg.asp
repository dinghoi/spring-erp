<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%

u_type = request("u_type")
insure_year = request("insure_year")

sql = "select * from insure_per order by insure_year desc"
Rs.Open Sql, Dbconn, 1

if u_type = "U" then
	sql = "select * from insure_per where insure_year = '" + insure_year + "'"
	Set rs_etc=DbConn.Execute(Sql)
	nps_per = rs_etc("nps_per")
	nhis_per = rs_etc("nhis_per")
	longcare_per = rs_etc("longcare_per")
	epi_person_per = rs_etc("epi_person_per")
	epi_company_per = rs_etc("epi_company_per")
	comwel_per = rs_etc("comwel_per")
	insure_tot_per = rs_etc("insure_tot_per")
	income_tax_per = rs_etc("income_tax_per")
	annual_pay_per = rs_etc("annual_pay_per")
	retire_pay_per = rs_etc("retire_pay_per")
	person_tot_per = rs_etc("person_tot_per")
	insure_memo = rs_etc("insure_memo")
  else
	nps_per = 0
	nhis_per = 0
	longcare_per = 0
	epi_person_per = 0
	epi_company_per = 0
	comwel_per = 0
	insure_tot_per = 0
	income_tax_per = 0
	annual_pay_per = 0
	retire_pay_per = 0
	person_tot_per = 0
	insure_memo = ""
end if	

title_line = "4대보험 요율 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			function frmsubmit () {
				document.condi_frm.submit ();
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.insure_year.value =="") {
					alert('기준년도를 입력하세요');
					frm.insure_year.focus();
					return false;}
				if(document.frm.nps_per.value ==0) {
					alert('국민연금율을 입력하세요');
					frm.nps_per.focus();
					return false;}
				if(document.frm.nhis_per.value ==0) {
					alert('건강보험율을 입력하세요');
					frm.nhis_per.focus();
					return false;}
				if(document.frm.longcare_per.value ==0) {
					alert('장기요양보험율을 입력하세요');
					frm.longcare_per.focus();
					return false;}
				if(document.frm.epi_person_per.value ==0) {
					alert('실업급여율를 입력하세요');
					frm.epi_person_per.focus();
					return false;}
				if(document.frm.epi_company_per.value ==0) {
					alert('고용안정율을 입력하세요');
					frm.epi_company_per.focus();
					return false;}
				if(document.frm.comwel_per.value ==0) {
					alert('산재보험율을 입력하세요');
					frm.comwel_per.focus();
					return false;}
				if(document.frm.income_tax_per.value ==0) {
					alert('소득세율을 입력하세요');
					frm.income_tax_per.focus();
					return false;}
				if(document.frm.annual_pay_per.value ==0) {
					alert('연차율을 입력하세요');
					frm.annual_pay_per.focus();
					return false;}
				if(document.frm.retire_pay_per.value ==0) {
					alert('퇴직금율을 입력하세요');
					frm.retire_pay_per.focus();
					return false;}
				if(document.frm.insure_memo.value =="") {
					alert('비고를 입력하세요');
					frm.insure_memo.focus();
					return false;}

				a=confirm('등록하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			
			}
			function Num_Check(obj) 
			{
				var word = obj.value;
				var str = "-.1234567890";
				for (i=0;i< word.length;i++){
					if(str.indexOf(word.charAt(i)) < 0){
						alert("숫자 조합만 가능합니다..");
						obj.value="";
						obj.focus();
						return false;
					}
				}
				nps_per = eval("document.frm.nps_per.value").replace(/,/g,"");
				nhis_per = eval("document.frm.nhis_per.value").replace(/,/g,"");
				longcare_per = eval("document.frm.longcare_per.value").replace(/,/g,"");
				epi_person_per = eval("document.frm.epi_person_per.value").replace(/,/g,"");
				epi_company_per = eval("document.frm.epi_company_per.value").replace(/,/g,"");
				comwel_per = eval("document.frm.comwel_per.value").replace(/,/g,"");
				income_tax_per = eval("document.frm.income_tax_per.value").replace(/,/g,"");
				annual_pay_per = eval("document.frm.annual_pay_per.value").replace(/,/g,"");
				retire_pay_per = eval("document.frm.retire_pay_per.value").replace(/,/g,"");
				insure_tot_per = parseFloat(nps_per) + parseFloat(nhis_per) + parseFloat(longcare_per) + parseFloat(epi_person_per) + parseFloat(epi_company_per) + parseFloat(comwel_per);
				person_tot_per = parseFloat(nps_per) + parseFloat(nhis_per) + parseFloat(longcare_per) + parseFloat(epi_person_per) + parseFloat(epi_company_per) + parseFloat(comwel_per) + parseFloat(annual_pay_per) + parseFloat(income_tax_per) + parseFloat(retire_pay_per);
				eval("document.frm.insure_tot_per.value = insure_tot_per.toFixed(3)");
				eval("document.frm.person_tot_per.value = person_tot_per.toFixed(3)");
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="75%" height="356" valign="top"><table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="6%" >
				          <col width="7%" >
				          <col width="7%" >
				          <col width="7%" >
				          <col width="7%" >
				          <col width="7%" >
				          <col width="7%" >
				          <col width="7%" >
				          <col width="7%" >
				          <col width="7%" >
				          <col width="7%" >
				          <col width="7%" >
				          <col width="*" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th rowspan="3" class="first" scope="col">기준년도</th>
				            <th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3;">4대 보험</th>
				            <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">기타 인건비 요율</th>
				            <th rowspan="3" scope="col">인당계</th>
				            <th rowspan="3" scope="col">비고</th>
			              </tr>
				          <tr>
				            <th rowspan="2" scope="col" style=" border-left:1px solid #e3e3e3;">국민연금</th>
				            <th rowspan="2" scope="col">건강보험</th>
				            <th rowspan="2" scope="col">장기요양</th>
				            <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">고용보험</th>
				            <th rowspan="2" scope="col">산재보험</th>
				            <th rowspan="2" scope="col">보험계</th>
				            <th rowspan="2" scope="col">소득세<br>
			                종업원분</th>
				            <th rowspan="2" scope="col">연차</th>
				            <th rowspan="2" scope="col">퇴직금</th>
			              </tr>
				          <tr>
				            <th scope="col" style=" border-left:1px solid #e3e3e3;">실업급여</th>
				            <th scope="col">고용안정</th>
			              </tr>
			            </thead>
			            <tbody>
						<%
                        do until rs.eof
                        %>
				        <tr>
				          <td class="first"><a href="insure_per_mg.asp?insure_year=<%=rs("insure_year")%>&u_type=<%="U"%>"><%=rs("insure_year")%></a></td>
				          <td class="right"><%=formatnumber(rs("nps_per"),3)%></td>
				          <td class="right"><%=formatnumber(rs("nhis_per"),3)%></td>
				          <td class="right"><%=formatnumber(rs("longcare_per"),3)%></td>
				          <td class="right"><%=formatnumber(rs("epi_person_per"),3)%></td>
				          <td class="right"><%=formatnumber(rs("epi_company_per"),3)%></td>
				          <td class="right"><%=formatnumber(rs("comwel_per"),3)%></td>
				          <td class="right"><%=formatnumber(rs("insure_tot_per"),3)%></td>
				          <td class="right"><%=formatnumber(rs("income_tax_per"),3)%></td>
				          <td class="right"><%=formatnumber(rs("annual_pay_per"),3)%></td>
				          <td class="right"><%=formatnumber(rs("retire_pay_per"),3)%></td>
				          <td class="right"><%=formatnumber(rs("person_tot_per"),3)%></td>
				          <td><%=rs("insure_memo")%></td>
			            </tr>
				        <%
							rs.movenext()
						loop
						%>
			            </tbody>
			          </table>
                      </td>
				      <td width="1%" valign="top">&nbsp;</td>
				      <td width="24%" valign="top"><form method="post" name="frm" action="insure_per_reg_ok.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				          <tbody>
				            <tr>
				              <th width="40%">기준년도</th>
				              <td class="left"><input name="insure_year" type="text" id="insure_year" onKeyUp="checkNum(this);" maxlength="4" value="<%=insure_year%>" style="width:100px;text-align:center"></td>
			                </tr>
				            <tr>
				              <th>국민연금(%)</th>
				              <td class="left"><input name="nps_per" type="text" id="nps_per" value="<%=formatnumber(nps_per,3)%>" onKeyUp="Num_Check(this);" maxlength="6" style="width:100px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>건강보험(%)</th>
				              <td class="left"><input name="nhis_per" type="text" id="nhis_per" value="<%=formatnumber(nhis_per,3)%>" onKeyUp="Num_Check(this);" maxlength="6" style="width:100px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>장기요양(%)</th>
				              <td class="left"><input name="longcare_per" type="text" id="longcare_per" value="<%=formatnumber(longcare_per,3)%>" onKeyUp="Num_Check(this);" maxlength="6" style="width:100px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>실업급여(%)</th>
				              <td class="left"><input name="epi_person_per" type="text" id="epi_person_per" value="<%=formatnumber(epi_person_per,3)%>" onKeyUp="Num_Check(this);" maxlength="6" style="width:100px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>고용안정(%)</th>
				              <td class="left"><input name="epi_company_per" type="text" id="epi_company_per" value="<%=formatnumber(epi_company_per,3)%>" onKeyUp="Num_Check(this);" maxlength="6" style="width:100px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>산재보험(%)</th>
				              <td class="left"><input name="comwel_per" type="text" id="comwel_per" value="<%=formatnumber(comwel_per,3)%>" onKeyUp="Num_Check(this);" maxlength="6" style="width:100px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>4대보험계(%)</th>
				              <td class="left"><input name="insure_tot_per" type="text" id="insure_tot_per" style="width:100px;text-align:right" value="<%=formatnumber(insure_tot_per,3)%>" readonly="true"></td>
			                </tr>
				            <tr>
				              <th>소득세<br>종업원분(%)</th>
				              <td class="left"><input name="income_tax_per" type="text" id="income_tax_per" value="<%=formatnumber(income_tax_per,3)%>" onKeyUp="Num_Check(this);" maxlength="6" style="width:100px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>연차(%)</th>
				              <td class="left"><input name="annual_pay_per" type="text" id="annual_pay_per" value="<%=formatnumber(annual_pay_per,3)%>" onKeyUp="Num_Check(this);" maxlength="6" style="width:100px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>퇴직금(%)</th>
				              <td class="left"><input name="retire_pay_per" type="text" id="retire_pay_per" value="<%=formatnumber(retire_pay_per,3)%>" onKeyUp="Num_Check(this);" maxlength="6" style="width:100px;text-align:right"></td>
			                </tr>
				            <tr>
				              <th>인당계(%)</th>
				              <td class="left"><input name="person_tot_per" type="text" id="person_tot_per" style="width:100px;text-align:right" value="<%=formatnumber(person_tot_per,3)%>" readonly="true"></td>
			                </tr>
				            <tr>
				              <th>비고</th>
				              <td class="left"><input name="insure_memo" type="text" id="insure_memo" onKeyUp="checklength(this,50)" value="<%=insure_memo%>" style="width:150px"></td>
			                </tr>
			              </tbody>
			            </table>
						<br>
				        <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				        <input type="hidden" name="old_insure_year" value="<%=insure_year%>" ID="Hidden1">
				        <div align=center>
                        	<span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                        	<span class="btnType01"><input type="button" value="취소" onclick="javascript:frmcancel();" ID="Button1" NAME="Button1"></span>
                        </div>
			          </form></td>
			        </tr>
				    <tr>
				      <td width="49%">&nbsp;</td>
				      <td width="2%">&nbsp;</td>
				      <td width="49%">&nbsp;</td>
			        </tr>
			      </table>
                </div>
			</div>				
	</div>        				
	</body>
</html>

