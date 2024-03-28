<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim u_type, org_code, view_condi
Dim code_last
Dim org_level, org_date, org_empno, org_empname
Dim org_company, org_bonbu, org_saupbu, org_team, org_reside_place
Dim org_reside_company, org_cost_group, org_cost_center, owner_org
Dim owner_orgname, owner_empno, owner_empname, org_table_org
Dim tel_ddd, tel_no1, tel_no2, org_sido, org_gugun, org_dong, org_addr
Dim org_end_date, org_reg_date, org_reg_user, org_mod_date, org_mod_user
Dim title_line, trade_code
Dim rs_max, max_seq

u_type = Request("u_type")
org_code = Request("org_code")
view_condi = Request("view_condi")

code_last = ""
org_level = ""
org_name = ""
org_date = ""
org_end_date = ""
org_empno = ""
org_empname = ""
org_company = ""
org_bonbu = ""
org_saupbu = ""
org_team = ""
org_reside_place = ""
org_reside_company = ""
org_cost_group = ""
org_cost_center = ""
owner_org = ""
owner_orgname = ""
owner_empno = ""
owner_empname = ""
org_table_org = 0
tel_ddd = ""
tel_no1 = ""
tel_no2 = ""
org_sido = ""
org_gugun = ""
org_dong = ""
org_addr = ""
org_end_date = ""
org_reg_date = ""
org_reg_user = ""
org_mod_date = ""
org_mod_user = ""
trade_code = ""

title_line = " ���� ��� "

'���� ������ ���
If u_type = "U" Then
	Dim rs

	'Sql="select * from emp_org_mst where org_code = '"&org_code&"'"
	objBuilder.Append "SELECT org_level, org_name, org_date, org_end_date, org_empno, "
	objBuilder.Append "org_emp_name, org_company, org_bonbu, org_saupbu, org_team, "
	objBuilder.Append "org_reside_place, org_reside_company, org_cost_group, org_cost_center, "
	objBuilder.Append "org_owner_org, org_owner_empno, org_owner_empname, org_table_org, "
	objBuilder.Append "org_tel_ddd, org_tel_no1, org_tel_no2, org_sido, org_gugun, "
	objBuilder.Append "org_dong, org_addr, org_end_date, org_reg_date, "
	objBuilder.Append "org_reg_user, org_mod_date, org_mod_user, "
	objBuilder.Append "(SELECT org_name FROM emp_org_mst "
	objBuilder.Append "	WHERE org_code = eomt.org_owner_org) AS owner_orgname "
	objBuilder.Append "FROM emp_org_mst AS eomt "
	objBuilder.Append "WHERE org_code = '"&org_code&"'"

	Set rs = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

    org_level = rs("org_level")
    org_name = rs("org_name")
    org_date = rs("org_date")
	org_end_date = rs("org_end_date")
    org_empno = rs("org_empno")
    org_empname = rs("org_emp_name")
    org_company = rs("org_company")
    org_bonbu = rs("org_bonbu")
    org_saupbu = rs("org_saupbu")
    org_team = rs("org_team")
	org_reside_place = rs("org_reside_place")
	org_reside_company = rs("org_reside_company")
	org_cost_group = rs("org_cost_group")
	org_cost_center = rs("org_cost_center")
    owner_org = rs("org_owner_org")
    owner_empno = rs("org_owner_empno")
    owner_empname = rs("org_owner_empname")

	If rs("org_table_org") = "" Or IsNull(rs("org_table_org")) Then
		org_table_org = 0
	Else
		org_table_org = rs("org_table_org")
	End If

    tel_ddd = rs("org_tel_ddd")
    tel_no1 = rs("org_tel_no1")
    tel_no2 = rs("org_tel_no2")
	org_sido = rs("org_sido")
    org_gugun = rs("org_gugun")
    org_dong = rs("org_dong")
    org_addr = rs("org_addr")
    org_end_date = rs("org_end_date")
    org_reg_date = rs("org_reg_date")
	org_reg_user = rs("org_reg_user")
    org_mod_date = rs("org_mod_date")
    org_mod_user = rs("org_mod_user")
	owner_orgname = rs("owner_orgname")

	rs.Close() : Set rs = Nothing

	title_line = " ���� ���� "
Else
	objBuilder.Append "SELECT MAX(org_code) AS max_seq FROM emp_org_mst "

	Set rs_max = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If IsNull(rs_max("max_seq")) Then
		code_last = "0001"
	Else
		max_seq = "000" + CStr((Int(rs_max("max_seq")) + 1))
		code_last = Right(max_seq, 4)
	End If

    rs_max.Close() : Set rs_max = Nothing

    org_code = code_last
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}

			//���� ������
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=org_date%>" );
			});

			//���� �����
			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=org_end_date%>" );
			});

			function goAction(){
			   window.close();
			}

			function frmcheck(){
				if (formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

     		function chkfrm(){
				/*if(document.frm.org_code.value ==""){
					alert('�����ڵ带 �Է����ּ���.');
					frm.org_code.focus();
					return false;
				}*/

				if(document.frm.org_name.value == ""){
					alert('�������� �Է����ּ���.');
					frm.org_name.focus();
					return false;
				}else{
					if(document.frm.org_name.value.indexOf('�ѻ�') > -1){
					  alert('"�ѻ�"�̶�� ���ڴ� �Է� �� �� �����ϴ�.("�ѻ�"->"��ȭ����")');
						frm.org_name.focus();
						return false;
					}
				}

				if(document.frm.org_date.value == ""){
					alert('�����������ڸ� �������ּ���.');
					frm.org_date.focus();
					return false;
				}

				if($('#org_level').val() === "����" && ($('#org_name').val() !== '�濵����' && $('#org_name').val() !== '���������')){
					if(document.frm.org_empno.value ==""){
						alert('���������� �Է����ּ���.');
						frm.org_empno.focus();
						return false;
					}

					if(document.frm.org_empname.value ==""){
						alert('�����强���� �Է����ּ���.');
						frm.org_empname.focus();
						return false;
					}
				}

//				if(document.frm.org_cost_group.value ==""){
//					alert('��뼾Ÿ�׷��� �����ϼ���');
//					frm.org_cost_group.focus();
//					return false;}

				if(document.frm.org_cost_center.value == ""){
					alert('��뱸���� �������ּ���.');
					frm.org_cost_center.focus();
					return false;
				}

				if(document.frm.org_level.value != "ȸ��"){
					if(document.frm.owner_org.value == ""){
						alert('���������� �Է����ּ���.');
						frm.owner_org.focus();
						return false;
					}
				}

				if(document.frm.org_level.value == "����ó"){
					if(document.frm.org_cost_center.value == "����������"){
						if(document.frm.org_reside_place.value == ""){
							alert('����ó�� �Է����ּ���.');
							frm.org_reside_place.focus();
							return false;
						}else{
							if(document.frm.org_reside_place.value.indexOf('�ѻ�') > -1){
								alert('"�ѻ�"�̶�� ���ڴ� �Է� �� �� �����ϴ�.("�ѻ�"->"��ȭ����")');
								frm.org_reside_place.focus();
								return false;
							}
  						}
					}
				}

				if(document.frm.org_cost_center.value == "����������"){
					if(document.frm.org_reside_company.value ==""){
						alert('����óȸ�縦 �Է����ּ���.');
						frm.org_reside_company.focus();
						return false;
					}else{
						if (document.frm.org_cost_center.value.indexOf('�ѻ�') > -1){
						  alert('"�ѻ�"�̶�� ���ڴ� �Է� �� �� �����ϴ�.("�ѻ�"->"��ȭ����")');
								frm.org_cost_center.focus();
								return false;
						}
  					}
				}

				if(document.frm.org_level.value == "����ó"){
					if(document.frm.org_reside_company.value == ""){
						alert('����ó ȸ�縦 �������ּ���.');
						frm.org_reside_company.focus();
						return false;
					}
				}

				if(document.frm.org_level.value == "����ó"){
					if(document.frm.org_cost_center.value != "����������"){
						alert('��뱸�п� ���������� �������ּ���.');
						frm.org_cost_center.focus();
						return false;
					}
				}

				/*if(document.frm.org_cost_group.value =="") {
					alert('����óȸ��(�ŷ�ó)�� �׷��̸��� �����ϴ�.');
					frm.org_reside_company.focus();
					return false;
				}*/

				if(!confirm('���� �Ͻðڽ��ϱ�?')) return false;
				else return true;
				/*
				{
					a=confirm('�Է��Ͻðڽ��ϱ�?')
					if (a==true) {
						return true;
					}
					return false;
				}*/
			}

			function num_chk(txtObj){
				org_to = parseInt(document.frm.org_table_org.value.replace(/,/g,""));

				org_to = String(org_to);
				num_len = org_to.length;
				sil_len = num_len;
				org_to = String(org_to);

				if(org_to.substr(0,1) == "-") sil_len = num_len - 1;
				if(sil_len > 3) org_to = org_to.substr(0,num_len -3) + "," + org_to.substr(num_len -3,3);
				if(sil_len > 6) org_to = org_to.substr(0,num_len -6) + "," + org_to.substr(num_len -6,3) + "," + org_to.substr(num_len -2,3);

				document.frm.org_table_org.value = org_to;
			}

			//�ּ� �˻� �˾�
			function jusoCallBack(roadFullAddr,roadAddrPart1,addrDetail,roadAddrPart2,engAddr,jibunAddr,zipNo,admCd,rnMgtSn,bdMgtSn,detBdNmList,bdNm,bdKdcd,siNm,sggNm,emdNm,liNm,rn,udrtYn,buldMnnm,buldSlno,mtYn,lnbrMnnm,lnbrSlno,emdNo,gubun){
				if(gubun === 'org'){
					$('#org_sido').val(siNm);
					$('#org_gugun').val(sggNm);
					$('#org_dong').val(rn + ' ' + buldMnnm);
					$('#org_addr').val(roadAddrPart2 + ' ' + addrDetail);
					$('#org_zipcode').val(zipNo);
				}else{
					alert('zip_address ���� ����');
				}
			}
		</script>
		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/insa/insa_org_reg_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
							<col width="8%" >
							<col width="17%" >
						</colgroup>
						<tbody>
                            <tr>
                                <th class="first" style="background:#FFFFE6">ȸ���</th>
                                <td colspan="7" class="left" bgcolor="#FFFFE6">
					            <input type="text" name="view_condi" id="view_condi" size="20" value="<%=view_condi%>" class="no-input" readonly/>
                                &nbsp;&nbsp;<span style="color:red;">�� ����ó�� ��뱸���� ������������ ��� �ʼ��� �Է��� �ϼž��մϴ�.</span>
                                </td>
                            </tr>
							<tr>
								<th class="first">���� �ڵ�</th>
                                <td class="left">
									<input type="text" name="org_code" value="<%=org_code%>" class="no-input" readonly/>
								</td>
                                <th>���� ����</th>
                                <td class="left">
                             <%
							 	Dim rsLevel

								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '01' ORDER BY emp_etc_code ASC "

								Set rsLevel = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
 							 %>
                                <select name="org_level" id="org_level" style="width:150px;" value="<%=org_level%>">
                             <%
								Do Until rsLevel.EOF
 			  				 %>
                                <option value='<%=rsLevel("emp_etc_name")%>' <%If org_level = rsLevel("emp_etc_name") Then %>selected<%End If%>><%=rsLevel("emp_etc_name")%></option>
                 			<%
									rsLevel.MoveNext()
								Loop

								rsLevel.Close() : Set rsLevel = Nothing
							%>
            					</select>
            					</td>
                                <th>������<span style="color:red;">*</span></th>
                                <td class="left">
									<input type="text" name="org_name" id="org_name" style="width:150px" value="<%=org_name%>" name="������" onKeyUp="checklength(this, 30);"/>
								</td>
                                <th>����������<span style="color:red;">*</span></th>
                                <td class="left">
									<input type="text" name="org_date" size="10" readonly="true" id="datepicker" style="width:70px;" value="<%=org_date%>"/>
              					</td>
                             </tr>
                             <tr>

								<th class="first">���������ڵ�</th>
                                <td class="left">
									<input type="text" name="owner_org" id="owner_org" size="4" readonly="true" value="<%=owner_org%>"/>
									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_org_select.asp?gubun=owner&mg_level=<%=org_level%>&view_condi=<%=view_condi%>','��������ã��','scrollbars=yes,width=850,height=400')">��������ã��</a>
                                </td>
                                <th>����������</th>
                                <td class="left">
									<input type="text" name="owner_orgname" id="owner_orgname" size="20" readonly="true" value="<%=owner_orgname%>"/>
								</td>
                                <th>�Ҽ�</th>
                                <td colspan="3" class="left">
									<input type="text" name="org_company" id="org_company" style="width:100px" readonly="true" value="<%=org_company%>"/>
									<input type="text" name="org_bonbu" id="org_bonbu" style="width:100px" readonly="true" value="<%=org_bonbu%>"/>
									<input type="text" name="org_saupbu" id="org_saupbu" style="width:100px" readonly="true" value="<%=org_saupbu%>"/>
									<input type="text" name="org_team" id="org_team" style="width:100px" readonly="true" value="<%=org_team%>"/>
                                </td>
                             </tr>
							<tr>
								<th>����������</th>
                                <td class="left">
									<input type="text" name="owner_empno" id="owner_empno" size="7" readonly="true" value="<%=owner_empno%>"/>
								</td>
                                <th>�����������</th>
                                <td class="left">
									<input type="text" name="owner_empname" id="owner_empname" size="20" readonly="true" value="<%=owner_empname%>"/>
								</td>
								<th class="first">������ ���</th>
                                <td class="left">
									<input type="text" name="org_empno" id="org_empno" size="7" readonly="true" value="<%=org_empno%>"/>
									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_emp_select.asp?gubun=<%="orgemp"%>&view_condi=<%=view_condi%>','orgempselect','scrollbars=yes,width=600,height=400')">������ã��</a>
                                </td>
                                <th>�����强��</th>
                                <td class="left">
									<input type="text" name="org_empname" id="org_empname" size="10" readonly="true" value="<%=org_empname%>"/>
                                </td>
                             </tr>
                             <tr>
								<th class="first">��ǥ��ȭ</th>
                                <td class="left">
									<input type="text" name="tel_ddd" id="tel_ddd" style="width:30px;" maxlength="3" value="<%=tel_ddd%>"  oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
									-
                                    <input type="text" name="tel_no1" id="tel_no1" style="width:40px;" maxlength="4" value="<%=tel_no1%>"  oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
                                    -
									<input type="text" name="tel_no2" id="tel_no2" style="width:40px;" maxlength="4" value="<%=tel_no2%>"  oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
								</td>
                                <th>���������</th>
                                <td class="left">
									<input type="text" name="org_end_date" size="10" readonly="true" id="datepicker1" style="width:70px;" value="<%=org_end_date%>"/>
              					</td>
                                <th>����ó</th>
                                <td class="left">
									<input type="text" name="org_reside_place" id="org_reside_place" style="width:150px" value="<%=org_reside_place%>"/>
                                </td>
                                <th class="first">����ó ȸ��</th>
								<td class="left">
									<input type="text" name="org_reside_company" id="org_reside_company" style="width:120px" readonly="true" value="<%=org_reside_company%>"/>
									<a href="#" class="btnType03" onClick="pop_Window('/insa/insa_trade_search.asp?gubun=1','tradesearch','scrollbars=yes,width=600,height=400')">ã��</a>
            					</td>
                             </tr>
                             <tr>
								<th class="first">�ּ�</th>
								<td colspan="5" class="left">
									<input type="text" name="org_sido" id="org_sido" style="width:100px;" readonly="true" value="<%=org_sido%>"/>
									<input type="text" name="org_gugun" id="org_gugun" style="width:150px;" readonly="true" value="<%=org_gugun%>"/>
									<input type="text" name="org_dong" id="org_dong" style="width:150px;" readonly="true" value="<%=org_dong%>"/>
									<input type="text" name="org_addr" id="org_addr" style="width:250px;" onKeyUp="checklength(this,50)" value="<%=org_addr%>"/>
              						<input type="hidden" name="org_zip" id="org_zip" value=""/>
									<a href="#" class="btnType03" onClick="pop_Window('/insa/jusoPopup.asp?gubun=org','�ּ� ��ȸ','scrollbars=yes,width=600,height=400')">�ּ���ȸ</a>
                                </td>
                                <th>��뼾Ÿ�׷�</th>
                                <td class="left">
									<input type="text" name="org_cost_group" id="org_cost_group" style="width:150px" readonly="true" value="<%=org_cost_group%>"/>
            					</td>
                              </tr>
                              <tr>
                                <th>�������</th>
                                <td class="left">
									<input type="text" name="org_reg_date" id="org_reg_date" style="width:150px;" value="<%=org_reg_date%>" class="no-input" readonly/>
                                </td>
                                <th>��������</th>
                                <td class="left">
									<input type="text" name="org_mod_date" id="org_mod_date" style="width:150px;" value="<%=org_mod_date%>" class="no-input" readonly/>
                                </td>
								<th class="first">�����ο�(T.O)</th>
								<td class="left">
									<input type="text" name="org_table_org" id="org_table_org" style="width:90px;text-align:right" value="<%=FormatNumber(org_table_org,0)%>" oninput="this.value=this.value.replace(/[^0-9.]/g, '').replace(/(\..*)\./g, '$1');"/>
            					</td>
                                <th>��� ����<span style="color:red;">*</span></th>
                                <td class="left">
                              <%
								Dim rsCostCenter
								objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
								objBuilder.Append "WHERE emp_etc_type = '70' "
								objBuilder.Append "ORDER BY emp_etc_code ASC "

								Set rsCostCenter = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  %>
								<select name="org_cost_center" id="org_cost_center" style="width:90px">
                                    <option value="" <% If org_cost_center = "" Then %>selected<% End If %>>����</option>
                			  <%
								Do Until rsCostCenter.EOF
			  				  %>
                					<option value='<%=rsCostCenter("emp_etc_name")%>' <%If org_cost_center = rsCostCenter("emp_etc_name") Then %>selected<%End If %>><%=rsCostCenter("emp_etc_name")%></option>
                			  <%
									rsCostCenter.Movenext()
								Loop

								rsCostCenter.Close() : Set rsCostCenter = Nothing
								DBConn.Close() : Set DBConn = Nothing
							  %>
                			     </select>
                                </td>
                              </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();"/></span>
                    <span class="btnType01"><input type="button" value="���" onclick="close_win();"/></span>
                </div>
                <input type="hidden" name="u_type" value="<%=u_type%>"/>
                <input type="hidden" name="mg_level" value="<%=org_level%>"/>
				<%'�ŷ�ó �ڵ� �߰�, �׷��� ���� ��� �ش� �׷��� �ŷ�ó �ڵ� ���%>
				<input type="hidden" name="trade_code" value="<%=trade_code%>"/>
			</form>
		</div>
	</div>
	</body>
</html>