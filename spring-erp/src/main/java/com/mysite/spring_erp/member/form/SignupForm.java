package com.mysite.spring_erp.member.form;

import jakarta.validation.constraints.NotEmpty;
import jakarta.validation.constraints.Size;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class SignupForm {
    @NotEmpty(message = "이름을 입력해주세요.")
    @Size(min = 2, max = 10, message = "이름은 2자 이상 10자 이하로 입력해주세요.")
    private String memName;

    // @NotEmpty(message = "아이디를 입력해주세요.")
    private String engName;

    @NotEmpty(message = "이메일을 입력해주세요.")
    private String memEmail;

    @NotEmpty(message = "비밀번호를 입력해주세요.")
    private String password;

    @NotEmpty(message = "비밀번호 확인을 입력해주세요.")
    private String confirmPassword;
}
