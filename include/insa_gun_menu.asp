			<% 
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
			%>
                <div class="btnRight">
          <a href="insa_gun_mg.asp" class="btnType01">조직별 근태현황</a>      	
					<a href="insa_year_leave_bat.asp" class="btnType01">연차휴가일수처리</a>
          <a href="insa_commute_mg.asp" class="btnType01">조직별 출퇴근현황</a>
          <a href="insa_commute_data_up.asp" class="btnType01">출근 엑셀 파일 업로드</a>
                    <%
                     '<a href="insa_gun_mg.asp" class="btnType01">조직별 근태현황</a>
					 '<a href="insa_gun_list.asp" class="btnType01">개인별 근태현황</a>
                   
                     '<a href="insa_system_popup.asp" class="btnType01">월별 근태</a>
                     '<a href="insa_system_popup.asp" class="btnType01">병.휴가현황</a>

					 '<a href="insa_gun_month_list.asp" class="btnType01">월별 근태</a>
                     '<a href="insa_gun_leave_list.asp" class="btnType01">병.휴가현황</a>
					%>
				</div>
