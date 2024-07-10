package com.mysite.spring_erp;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.assertArg;

import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import com.mysite.spring_erp.board.entity.EmpBoard;
import com.mysite.spring_erp.board.repository.BoardRepository;

@SpringBootTest
class SpringErpApplicationTests {

	// @Test
	// void contextLoads() {
	// }

	@Autowired // DI(Dependency Injection)
	private BoardRepository boardRepository;

	@Test
	void testJap() {
		// 데이터 입력
		// Board b1 = new Board();
		// b1.setBoardType("0");
		// b1.setBoardTitle("안녕하세요.");
		// b1.setBoardContent("만나서 반갑습니다.^^");
		// b1.setReadCnt(0);
		// b1.setCreatedDate(LocalDateTime.now());
		// b1.setUpdatedDate(LocalDateTime.now());
		// this.boardRepository.save(b1);

		// Board b2 = new Board();
		// b2.setBoardType("1");
		// b2.setBoardTitle("Spring Boot 재미있나요?");
		// b2.setBoardContent("스프링 부트 공주중인데 너무 재미있는 것 같아요.");
		// b2.setReadCnt(0);
		// b2.setCreatedDate(LocalDateTime.now());
		// b2.setUpdatedDate(LocalDateTime.now());
		// this.boardRepository.save(b2);

		// 데이터 조회
		// List<Board> boards = this.boardRepository.findAll(); // 전체 레코드를 가져옴
		// assertEquals(2, boards.size()); // 2개의 레코드가 있어야 함

		// Board b = boards.get(0); // 첫 번째 레코드를 가져옴
		// assertEquals("안녕하세요.", b.getBoardTitle()); // 첫 번째 레코드의 제목은 "안녕하세요."여야 함

		// Optional<Board> oa = this.boardRepository.findById(1L); // 1번 레코드를 가져옴
		// if (oa.isPresent()) { // 1번 레코드가 있으면
		// Board board = oa.get(); // 1번 레코드를 가져옴
		// assertEquals("안녕하세요.", board.getBoardTitle()); // 1번 레코드의 제목은 "안녕하세요."여야 함
		// }

		// And 검색
		// Board b = this.boardRepository.findByBoardTitleAndBoardContent("안녕하세요.", "만나서
		// 반갑습니다.^^");
		// assertEquals(1, b.getBoardSeq());

		// Like 검색
		// List<Board> bList = this.boardRepository.findByBoardTitleLike("안녕%");
		// Board b = bList.get(0);
		// assertEquals("안녕하세요.", b.getBoardTitle());

		// 데이터 수정
		// Optional: 값이 있을 수도 있고 없을 수도 있음, 값이 있으면 get()으로 값을 가져옴
		Optional<EmpBoard> oa = this.boardRepository.findById(1L);
		assertTrue(oa.isPresent());
		EmpBoard b = oa.get();
		b.setBoardTitle("수정된 제목");
		this.boardRepository.save(b);
	}
}
