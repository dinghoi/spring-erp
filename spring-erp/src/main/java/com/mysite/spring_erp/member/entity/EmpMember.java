package com.mysite.spring_erp.member.entity;

import java.time.LocalDateTime;

import org.hibernate.annotations.DynamicInsert;
import org.hibernate.annotations.DynamicUpdate;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@Entity
@DynamicInsert
@DynamicUpdate
public class EmpMember {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long memSeq;

    @Column(length = 20, nullable = false)
    private String memName;

    @Column(length = 30)
    private String engName;

    @Column(length = 30, nullable = false, unique = true)
    private String memEmail;

    private LocalDateTime createdDate;
    private LocalDateTime updatedDate;
}
