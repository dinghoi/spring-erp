				<div class="btnRight">
                <%
				in_name = request.cookies("nkpmg_user")("coo_user_name")
                in_empno = request.cookies("nkpmg_user")("coo_user_id")
				position = request.cookies("nkpmg_user")("coo_position")
				met_grade = request.cookies("nkpmg_user")("coo_insa_grade")
				%>  
					<a href="met_stock_pummok_inout01.asp" class="btnType01">ǰ��/�Ⱓ�� �������Ȳ</a>
                <% 
                    '<a href="met_pummok_inout_mg.asp" class="btnType01">�Ⱓ�� �������Ȳ</a>
                %>
                    <a href="met_goods_pum_jaego_mg.asp" class="btnType01">ǰ�� ���</a>
                    <a href="met_stock_pum_jaego_mg.asp" class="btnType01">â�� ���</a>
                    <a href="met_goods_pum_jaego_report.asp" class="btnType01">���(�ݾ׼�)��Ȳ</a>
                    <a href="met_stock_subul_mg.asp" class="btnType01">������Ȳ</a>
                    <a href="met_stock_jaego_year_trans1.asp" class="btnType01">�����̿�</a>
				<% if in_empno = "900002" then	%>		                    
                    <a href="met_stock_chulgo_amt_data.asp" class="btnType01">���ݾ�</a>
                <% end if	%>                               
                <% '   <a href="meterials_system_popup.asp" class="btnType01">���ǻ�</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">�԰�����</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">�������</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">�ҿ������Ȳ</a>
                    '<a href="meterials_system_popup.asp" class="btnType01">�ҿ�������</a> %>
				</div>
