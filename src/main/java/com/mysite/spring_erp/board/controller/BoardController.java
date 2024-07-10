package com.mysite.spring_erp.board.controller;

import java.security.Principal;
import java.util.List;

import org.springframework.security.access.prepost.PreAuthorize;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import lombok.RequiredArgsConstructor;
// import org.springframework.web.bind.annotation.RequestParam;
// import org.springframework.web.bind.annotation.RequestParam;

import com.mysite.spring_erp.board.entity.EmpBoard;
import com.mysite.spring_erp.board.form.BoardForm;
import com.mysite.spring_erp.board.service.BoardService;
import com.mysite.spring_erp.member.entity.EmpMaster;
import com.mysite.spring_erp.member.service.MasterService;

import jakarta.validation.Valid;

import org.springframework.web.bind.annotation.PostMapping;
// import org.springframework.web.bind.annotation.RequestBody;

@RequestMapping("/board")
@RequiredArgsConstructor
@Controller
public class BoardController {
    private final BoardService boardService;
    private final MasterService masterService;

    // 게시글 목록 조회
    @GetMapping("/list")
    @PreAuthorize("isAuthenticated()") // 로그인한 사용자만 접근 가능
    public String board_list(Model model) {
        List<EmpBoard> boardList = this.boardService.getList();
        model.addAttribute("boardList", boardList);
        return "board/board_list";
    }

    // 게시글 상세 조회
    @GetMapping("/detail/{id}")
    @PreAuthorize("isAuthenticated()")
    public String board_detail(Model model, @PathVariable("id") Long id) {
        EmpBoard board = this.boardService.getBoard(id);
        model.addAttribute("board", board);
        return "board/board_detail";
    }

    // 게시글 작성 폼
    @GetMapping("/create")
    @PreAuthorize("isAuthenticated()")
    public String board_form(BoardForm boardForm) {
        return "board/board_form";
    }

    // @PostMapping("/create")
    // public String create(@RequestParam(value = "boardTitle") String boardTitle,
    // @RequestParam(value = "boardContent") String boardContent) {
    // this.boardService.saveBoard(boardTitle, boardContent);

    // return "redirect:/board/list"; // 게시글 작성 후 게시글 목록으로 이동
    // }

    // 게시글 저장
    // @Valid 어노테이션을 사용하여 BoardForm 객체에 대한 검증을 수행
    // BindingResult 객체를 사용하여 검증 결과를 확인
    @PostMapping("/create")
    @PreAuthorize("isAuthenticated()")
    public String board_create(@Valid BoardForm boardForm, BindingResult bindingResult, Principal principal) {
        if (bindingResult.hasErrors()) {
            return "board/board_form";
        }
        EmpMaster writer = this.masterService.getEmpMaster(principal.getName());
        this.boardService.saveBoard(boardForm.getBoardTitle(), boardForm.getBoardContent(), writer);

        return "redirect:/board/list"; // 게시글 작성 후 게시글 목록으로 이동
    }
}
