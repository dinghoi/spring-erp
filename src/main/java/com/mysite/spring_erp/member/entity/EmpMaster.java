package com.mysite.spring_erp.member.entity;

import java.time.LocalDateTime;

import org.hibernate.annotations.ColumnDefault;
import org.hibernate.annotations.DynamicInsert;
// import org.hibernate.annotations.DynamicUpdate;

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
// @DynamicUpdate
public class EmpMaster {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "emp_seq")
    private int id;

    @Column(name = "emp_no", length = 6, nullable = false, unique = true)
    @ColumnDefault("100000")
    private String no;

    @Column(name = "emp_pwd", length = 100, nullable = false)
    private String pwd;

    @Column(name = "emp_name", length = 30, nullable = false)
    private String name;

    @Column(name = "emp_status", length = 1, nullable = false)
    @ColumnDefault("'N'")
    private String status;

    private LocalDateTime createdDate;
    private LocalDateTime updatedDate;
}
