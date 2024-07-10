package com.mysite.spring_erp.member.service;

import java.time.LocalDateTime;

import org.springframework.stereotype.Service;

import com.mysite.spring_erp.member.entity.EmpMaster;
import com.mysite.spring_erp.member.entity.EmpMember;
import com.mysite.spring_erp.member.repository.MemberRepository;

import lombok.RequiredArgsConstructor;

@RequiredArgsConstructor
@Service
public class MemberService {
    private final MemberRepository empMemberRepository;

    public EmpMember saveEmpMember(String ename, String email, EmpMaster empMaster) {
        EmpMember member = new EmpMember();
        member.setEname(ename);
        member.setEmail(email);
        member.setCreatedDate(LocalDateTime.now());
        member.setUpdatedDate(LocalDateTime.now());
        member.setEmpMaster(empMaster);
        this.empMemberRepository.save(member);

        return member;
    }
}
