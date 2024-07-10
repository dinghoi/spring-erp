package com.mysite.spring_erp.member.entity;

import java.time.LocalDateTime;

import org.hibernate.annotations.ColumnDefault;
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
public class EmpMaster {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long empSeq;

    @Column(length = 6, nullable = false, unique = true)
    @ColumnDefault("100000")
    private String empNo;

    @Column
    private String empPwd;

    @ColumnDefault("'N'")
    private String empStatus;

    private LocalDateTime createdDate;
    private LocalDateTime updatedDate;
}
