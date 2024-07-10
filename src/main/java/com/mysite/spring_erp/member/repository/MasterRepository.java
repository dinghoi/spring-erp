package com.mysite.spring_erp.member.repository;

import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;

import com.mysite.spring_erp.member.entity.EmpMaster;

public interface MasterRepository extends JpaRepository<EmpMaster, Long> {
    public Optional<EmpMaster> findByEmpNo(String empNo);
}
