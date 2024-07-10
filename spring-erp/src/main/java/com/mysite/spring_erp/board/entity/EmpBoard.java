package com.mysite.spring_erp.board.entity;

import java.time.LocalDateTime;

import org.hibernate.annotations.ColumnDefault;
import org.hibernate.annotations.DynamicInsert;
import org.hibernate.annotations.DynamicUpdate;

import com.mysite.spring_erp.member.entity.EmpMaster;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
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
    private Long boardSeq;

    @ColumnDefault("'1'")
    private String boardType;

    @Column(length = 200)
    private String boardTitle;

    @Column(columnDefinition = "TEXT")
    private String boardContent;

    @ColumnDefault("0")
    private Integer readCnt;

    @Column(length = 200)
    private String attFile;

    @ManyToOne
    private EmpMaster writer;

    private LocalDateTime createdDate;
    private LocalDateTime updatedDate;
}
