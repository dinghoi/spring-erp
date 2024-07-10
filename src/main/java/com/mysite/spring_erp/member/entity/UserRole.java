package com.mysite.spring_erp.member.entity;

import lombok.Getter;

@Getter
public enum UserRole {
    ADMIN("ROLE_ADMIN"),
    USER("ROLE_USER");

    UserRole(String code) {
        this.code = code;
    }

    private final String code;
}
