package com.mysite.spring_erp.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;

@Configuration
public class SpringConfig {
    @Bean
    public PasswordEncoder passwordEncoder() { // BCryptPasswordEncoder 빈 등록
        return new BCryptPasswordEncoder(); // BCryptPasswordEncoder 인스턴스 반환
    }
}
