package com.mysite.spring_erp.member.entity;

import java.time.LocalDateTime;

import org.hibernate.annotations.DynamicInsert;
import org.hibernate.annotations.DynamicUpdate;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.JoinColumn;
import jakarta.persistence.OneToOne;
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
    @Column(name = "mem_seq")
    private int id;

    @Column(name = "eng_name", length = 30)
    private String ename;

    @Column(name = "mem_email", length = 30, nullable = false, unique = true)
    private String email;

    private LocalDateTime createdDate;
    private LocalDateTime updatedDate;

    @OneToOne
    @JoinColumn(name = "emp_seq")
    private EmpMaster empMaster;
}
