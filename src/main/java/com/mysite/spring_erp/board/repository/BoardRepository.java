package com.mysite.spring_erp.board.repository;

import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;

import com.mysite.spring_erp.board.entity.EmpBoard;

public interface BoardRepository extends JpaRepository<EmpBoard, Integer> {
    // EmpBoard findByBoardTitleAndBoardContent(String title, String content);

    // List<EmpBoard> findByBoardTitleLike(String title);
}
