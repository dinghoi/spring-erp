				<div class="btnRight">
					<a href="as_list_asset.asp" class="btnType01">A/S 총괄 현황</a>
<%
						if emp_no = "999998" then
							Response.write ""
						else
							Response.write "<a href='nh_as_reg.asp' target='_blank'>A/S 신청 접수</a>&nbsp;"
						end if
%>

				</div>
