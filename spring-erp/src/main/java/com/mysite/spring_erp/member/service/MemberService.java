package com.mysite.spring_erp.member.service;

import java.time.LocalDateTime;

import org.springframework.stereotype.Service;

import com.mysite.spring_erp.member.entity.EmpMember;
import com.mysite.spring_erp.member.repository.MemberRepository;

import lombok.RequiredArgsConstructor;

@RequiredArgsConstructor
@Service
public class MemberService {
    private final MemberRepository empMemberRepository;

    public EmpMember saveEmpMember(String memName, String engName, String memEmail) {
        EmpMember member = new EmpMember();
        member.setMemName(memName);
        member.setEngName(engName);
        member.setMemEmail(memEmail);
        member.setCreatedDate(LocalDateTime.now());
        member.setUpdatedDate(LocalDateTime.now());
        this.empMemberRepository.save(member);

        return member;
    }
}
