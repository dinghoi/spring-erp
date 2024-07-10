package com.mysite.spring_erp.member.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import com.mysite.spring_erp.member.entity.EmpMember;

public interface MemberRepository extends JpaRepository<EmpMember, Long> {

}
