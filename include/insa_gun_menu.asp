			<% 
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
			%>
                <div class="btnRight">
          <a href="insa_gun_mg.asp" class="btnType01">������ ������Ȳ</a>      	
					<a href="insa_year_leave_bat.asp" class="btnType01">�����ް��ϼ�ó��</a>
          <a href="insa_commute_mg.asp" class="btnType01">������ �������Ȳ</a>
          <a href="insa_commute_data_up.asp" class="btnType01">��� ���� ���� ���ε�</a>
                    <%
                     '<a href="insa_gun_mg.asp" class="btnType01">������ ������Ȳ</a>
					 '<a href="insa_gun_list.asp" class="btnType01">���κ� ������Ȳ</a>
                   
                     '<a href="insa_system_popup.asp" class="btnType01">���� ����</a>
                     '<a href="insa_system_popup.asp" class="btnType01">��.�ް���Ȳ</a>

					 '<a href="insa_gun_month_list.asp" class="btnType01">���� ����</a>
                     '<a href="insa_gun_leave_list.asp" class="btnType01">��.�ް���Ȳ</a>
					%>
				</div>
