package com.mysite.spring_erp.member.controller;

import org.springframework.dao.DataIntegrityViolationException;
import org.springframework.stereotype.Controller;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import com.mysite.spring_erp.member.form.SignupForm;
import com.mysite.spring_erp.member.service.MasterService;
import com.mysite.spring_erp.member.service.MemberService;

// import groovy.util.logging.Slf4j;
import jakarta.validation.Valid;
import lombok.RequiredArgsConstructor;
// import org.springframework.web.bind.annotation.RequestParam;

@RequestMapping("/member")
@RequiredArgsConstructor
@Controller
// @Slf4j
public class MasterController {
    private final MasterService empMasterService;
    private final MemberService empMemberService;

    // 회원가입 페이지
    @GetMapping("/signup")
    public String form(SignupForm signupForm) {
        return "member/signup";
    }

    @PostMapping("/signup")
    public String form(@Valid SignupForm signupForm, BindingResult bindingResult) {
        // 회원가입 입력값 검증
        if (bindingResult.hasErrors()) {
            return "member/signup";
        }

        // 비밀번호 확인
        // if (signupForm.getPassword() != signupForm.getConfirmPassword()) { // error
        if (!signupForm.getPassword().equals(signupForm.getConfirmPassword())) {
            bindingResult.rejectValue(
                    "confirmPassword",
                    "defferentPassword",
                    "비밀번호가 일치하지 않습니다.");
            return "member/signup";
        }

        // 회원가입 처리
        try {
            this.empMasterService.saveEmpMaster(
                    signupForm.getPassword());

            this.empMemberService.saveEmpMember(
                    signupForm.getMemName(),
                    signupForm.getEngName(),
                    signupForm.getMemEmail());
        } catch (DataIntegrityViolationException e) {
            bindingResult.reject(
                    "alreadyInUser",
                    "중복 이메일, 이미 등록된 이메일입니다.");
            return "member/signup";
        } catch (Exception e) {
            // log.error("회원가입 오류", e);
            e.printStackTrace();
            bindingResult.reject(
                    "unexpectedError",
                    "알 수 없는 에러가 발생했습니다.");
            return "member/signup";
        }
        return "redirect:/";
    }

    @GetMapping("/login")
    public String login() {
        return "member/login";
    }

}
