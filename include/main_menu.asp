				<%
				'[�ӿ� ����] �޴� ���� ���� ����[20201221_����ȣ]
				Dim ceoGradeYn : ceoGradeYn = "N"

				'102592 : ������ ���� ��� �߰�
				'user_id = 900001, 100085, 100739, 101100, 101574, 101575, 101664, 102305, 102306 : ���� ����
				If emp_no = "100001" Or user_id = "100740" Or user_id = "102592" Then	'�
					ceoGradeYn = "Y"
				End If

				%>
				<div class="btnLeft">
					<a href="/main/nkp_main.asp"><img src="/image/home.gif" alt="" name="img01" width="65" height="65" border="0"></a>
					<%
						'���ξ���
						If insa_grade < "2" Then
							Response.write "<a href='/person/insa_person_mg.asp' target='_blank'><img src='/image/person_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						Else
							Response.write "<a><img src='/image/person_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						End If

						'���񽺰���
					%>
					<a href="/as_list_ce.asp" target="_blank"><img src="/image/as_mg.gif" alt="" name="img01" width="65" height="65" border="0"></a>
					<%
						'������
						If cost_grade < "7" Then
							Response.write "<a href='/person_cost_report.asp' target='_blank'><img src='/image/cost_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						Else
							Response.write "<a><img src='/image/cost_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						End If

						'�λ����
						If insa_grade = "0" Then
							Response.write "<a href='/insa/insa_report_mg.asp' target='_blank'><img src='/image/insa_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						Else
							Response.write "<a><img src='/image/insa_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						End If

						'�޿�����
						If pay_grade = "0" Then
							Response.write "<a href='/pay/insa_pay_mg.asp' target='_blank'><img src='/image/pay_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						Else
							Response.write "<a><img src='/image/pay_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						End If

						'��������
						If sales_grade < "2" Then
							Response.write "<a href='/sales/sales_report.asp' target='_blank'><img src='/image/sales_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						Else
					  		Response.write "<a><img src='/image/sales_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						End If

						Dim metUrl

						'��ǰ�������
						If met_grade < "4" Then
							Select Case met_grade
								Case "0"
									metUrl = "/met_stock_in_report01.asp"
								Case "2"
									metUrl = "/met_stock_nwin_report01.asp"
								Case Else
									metUrl = "/met_stock_out_ce_mg.asp"
							End Select

							Response.write "<a href='"&metUrl&"' target='_blank'><img src='/image/goods_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						Else
							Response.write "<a><img src='/image/goods_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						End If

						'ȸ�����
						If account_grade = "0" Then
							Response.write "<a href='/finance/card_slip_mg.asp' target='_blank'><img src='/image/account_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						Else
							Response.write "<a><img src='/image/account_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						End If

						'�ӿ�����
						'If emp_no = "100001" Or user_id = "900001" Or user_id = "100085" Or user_id = "102305" Or user_id = "102306" Or user_id = "100739" Or user_id = "100740" Or user_id = "101100" Or user_id = "101574" Or user_id = "101575" Or user_id = "101664" Then
						If ceoGradeYn = "Y" Then
							Response.write "<a href='/ceo_total_info.asp' target='_blank'><img src='/image/eis_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						Else
							Response.write "<a><img src='/image/eis_mg.gif' alt='' name='img01' width='65' height='65' border='0'></a>&nbsp;"
						End If

						'�׷����
					%>
					<a href='http://gw.k-won.co.kr/groupware/login.php' target='_blank'><img src='/image/groupware.gif' alt='' name='img01' width='65' height='65' border='0'></a>
				</div>
