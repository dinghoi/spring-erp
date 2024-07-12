package com.mysite.spring_erp.member.service;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.authority.SimpleGrantedAuthority;
import org.springframework.security.core.userdetails.User;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.stereotype.Service;

import com.mysite.spring_erp.member.entity.EmpMaster;
import com.mysite.spring_erp.member.entity.UserRole;
import com.mysite.spring_erp.member.repository.MasterRepository;

import lombok.RequiredArgsConstructor;

@Service
@RequiredArgsConstructor
public class UserSecurityService implements UserDetailsService {

    private final MasterRepository masterRepository;

    @Override
    public UserDetails loadUserByUsername(String username) throws UsernameNotFoundException {
        Optional<EmpMaster> _user = this.masterRepository.findByNo(username);

        System.out.println(username);

        if (_user.isEmpty()) {
            throw new UsernameNotFoundException("등록되지 않은 사번입니다.");
        }

        EmpMaster empMaster = _user.get();

        List<GrantedAuthority> authorities = new ArrayList<>();
        if ("admin".equals(username)) {
            authorities.add(new SimpleGrantedAuthority(UserRole.ADMIN.getCode()));
        } else {
            authorities.add(new SimpleGrantedAuthority(UserRole.USER.getCode()));
        }

        return new User(empMaster.getNo(), empMaster.getPwd(), authorities);

    }
}
