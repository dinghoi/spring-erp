package com.mysite.spring_erp.board.controller;

import java.security.Principal;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.springframework.http.HttpStatus;
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

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpSession;
import jakarta.validation.Valid;

import org.springframework.web.bind.annotation.PostMapping;
// import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.server.ResponseStatusException;

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
    public String board_detail(@PathVariable("id") Integer id, HttpServletRequest request, Model model) {
        // 게시글 조회 시 조회수 증가
        // 세션에 readSet이라는 이름으로 Set 객체를 저장
        HttpSession session = request.getSession();
        Set<Integer> readSet = (Set<Integer>) session.getAttribute("readSet");
        // 세션에 readSet이 없으면 Set 객체를 생성하여 세션에 저장
        if (readSet == null) {
            readSet = new HashSet<>();
            session.setAttribute("readSet", readSet);
        }
        // 게시글을 읽은 게시글 번호를 저장하는 Set 객체에 게시글 번호를 저장
        if (!readSet.contains(id)) {
            readSet.add(id); // 게시글 번호를 Set 객체에 저장
            boardService.increaseReadCnt(id); // 게시글 조회수 증가
        }

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
        EmpMaster id = this.masterService.getEmpMaster(principal.getName());
        this.boardService.saveBoard(boardForm.getBoardTitle(), boardForm.getBoardContent(), id);

        return "redirect:/board/list"; // 게시글 작성 후 게시글 목록으로 이동
    }

    // 게시글 수정 폼
    @GetMapping("/modify/{id}")
    @PreAuthorize("isAuthenticated()")
    public String board_modify(BoardForm boardForm, @PathVariable("id") int id, Principal principal) {
        EmpBoard board = this.boardService.getBoard(id);

        if (!board.getEmpMaster().getNo().equals(principal.getName())) {
            throw new ResponseStatusException(HttpStatus.BAD_REQUEST, "게시글 작성자만 수정할 수 있습니다.");
        }
        boardForm.setBoardTitle(board.getTitle());
        boardForm.setBoardContent(board.getContent());

        return "board/board_form";
    }

    // 게시글 수정
    @PostMapping("/modify/{id}")
    @PreAuthorize("isAuthenticated()")
    public String board_modify(@Valid BoardForm boardForm, BindingResult bindingResult, Principal principal,
            @PathVariable("id") int id) {
        if (bindingResult.hasErrors()) { // 입력값 검증
            return "board/board_form";
        }

        EmpBoard board = this.boardService.getBoard(id);
        if (!board.getEmpMaster().getNo().equals(principal.getName())) {
            throw new ResponseStatusException(HttpStatus.BAD_REQUEST, "게시글 작성자만 수정할 수 있습니다.");
        }

        this.boardService.updateBoard(board, boardForm.getBoardTitle(),
                boardForm.getBoardContent());

        return String.format("redirect:/board/detail/%s", id);

    }

}
