package com.mysite.spring_erp.board.entity;

import java.time.LocalDateTime;

import org.hibernate.annotations.ColumnDefault;
import org.hibernate.annotations.DynamicInsert;
// import org.hibernate.annotations.DynamicUpdate;

import com.mysite.spring_erp.member.entity.EmpMaster;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.JoinColumn;
import jakarta.persistence.ManyToOne;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@Entity
@DynamicInsert
public class EmpBoard {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "board_seq")
    private int id;

    @Column(name = "board_type", length = 1, nullable = false)
    @ColumnDefault("'1'")
    private String type;

    @Column(name = "board_title", length = 200, nullable = false)
    private String title;

    @Column(name = "board_content", columnDefinition = "TEXT", nullable = false)
    private String content;

    @Column(name = "read_cnt", nullable = false)
    @ColumnDefault("0")
    private int readcnt;

    @Column(name = "att_file", length = 200)
    private String file;

    @Column(name = "created_date", nullable = false)
    private LocalDateTime created;

    @Column(name = "updated_date", nullable = false)
    private LocalDateTime updated;

    @ManyToOne
    @JoinColumn(name = "emp_seq")
    private EmpMaster empMaster;
}
