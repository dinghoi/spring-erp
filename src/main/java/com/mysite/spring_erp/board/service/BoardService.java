package com.mysite.spring_erp.board.service;

// import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;

// import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
// import org.springframework.validation.BindingResult;
// import org.springframework.web.bind.annotation.PostMapping;

import com.mysite.spring_erp.board.entity.EmpBoard;
import com.mysite.spring_erp.board.repository.BoardRepository;
import com.mysite.spring_erp.exception.DataNotFoundException;
import com.mysite.spring_erp.member.entity.EmpMaster;

// import jakarta.validation.Valid;
import lombok.RequiredArgsConstructor;

@RequiredArgsConstructor
@Service
public class BoardService {
    private final BoardRepository boardRepository;

    // @Autowired
    // public BoardService(BoardRepository boardRepository){
    // this.boardRepository = boardRepository;
    // }

    // 게시글 목록 조회
    public List<EmpBoard> getList() {
        return this.boardRepository.findAll();
    }

    // 게시글 상세 조회
    public EmpBoard getBoard(int id) {
        Optional<EmpBoard> board = this.boardRepository.findById(id);
        if (board.isPresent()) {
            return board.get();
        } else {
            throw new DataNotFoundException("게시글이 존재하지 않습니다.");
        }
    }

    // 게시글 저장
    public void saveBoard(String title, String content, EmpMaster name) {
        EmpBoard board = new EmpBoard();
        board.setTitle(title);
        board.setContent(content);
        board.setEmpMaster(name);
        board.setCreated(LocalDateTime.now());
        board.setUpdated(LocalDateTime.now());
        this.boardRepository.save(board);
    }

    // 게시글 조회수 증가
    public void increaseReadCnt(int id) {
        // 게시글 조회수 증가
        EmpBoard board = boardRepository.findById(id)
                .orElseThrow(() -> new IllegalArgumentException("Invalid boardSeq:" + id)); // 게시글이 존재하지 않으면 예외 발생
        board.setReadcnt(board.getReadcnt() + 1); // 조회수 증가
        this.boardRepository.save(board);
    }

    // 게시글 수정
    public void updateBoard(EmpBoard board, String title, String content) {
        board.setTitle(title);
        board.setContent(content);
        board.setUpdated(LocalDateTime.now());
        this.boardRepository.save(board);
    }

    // 게시글 삭제
    public void deleteBoard(EmpBoard board) {
        this.boardRepository.delete(board);
    }
}