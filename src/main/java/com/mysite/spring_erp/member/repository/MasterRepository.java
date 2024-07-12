package com.mysite.spring_erp.member.repository;

import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;

import com.mysite.spring_erp.member.entity.EmpMaster;

public interface MasterRepository extends JpaRepository<EmpMaster, Integer> {
    public Optional<EmpMaster> findByNo(String empno);

    // 마지막 사원번호 조회
    public Optional<EmpMaster> findTopByOrderByNoDesc();
}
