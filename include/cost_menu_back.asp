				<div class="btnRight">
					<a href="person_cost_report.asp" class="btnType01">개인별비용정산</a>
					<a href="general_cost_mg.asp" class="btnType01">일반경비</a>
				<% if cost_grade = "0" or cost_grade = "3" then  %>
					<a href="others_cost_mg.asp" class="btnType01">비용대행등록</a>
                <% end if	%>
				<% if cost_grade = "6" or cost_grade = "5" or cost_grade < "3" then  %>
					<a href="overtime_mg.asp" class="btnType01">야특근</a>
                <% end if	%>
				<% if cost_grade = "6" or cost_grade = "5" or cost_grade < "3" then  %>
					<a href="transit_cost_mg.asp" class="btnType01">교통비</a>
                <% end if	%>
					<a href="person_card_mg.asp" class="btnType01">개인별카드내역</a>
				<% if cost_grade < "6" then  %>
					<a href="tax_esero_in_mg.asp" class="btnType01">E세로매입세금계산서</a>
					<a href="rent_cost_mg.asp" class="btnType01">임차료</a>
					<a href="outside_cost_mg.asp" class="btnType01">외주비</a>
					<a href="parcel_cost_mg.asp" class="btnType01">운반비</a>
					<a href="etc_cost_mg.asp" class="btnType01">자재및장비</a>
					<a href="tax_bill_cost_mg.asp" class="btnType01">계산서일반비용</a>
                <% end if	%>
				</div>
