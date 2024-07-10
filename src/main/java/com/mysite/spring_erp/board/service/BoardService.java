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
    public EmpBoard getBoard(Long id) {
        Optional<EmpBoard> board = this.boardRepository.findById(id);
        if (board.isPresent()) {
            return board.get();
        } else {
            throw new DataNotFoundException("게시글이 존재하지 않습니다.");
        }
    }

    // 게시글 저장
    public void saveBoard(String title, String content, EmpMaster writer) {
        EmpBoard board = new EmpBoard();
        board.setBoardTitle(title);
        board.setBoardContent(content);
        board.setWriter(writer);
        board.setCreatedDate(LocalDateTime.now());
        board.setUpdatedDate(LocalDateTime.now());
        this.boardRepository.save(board);
    }
}