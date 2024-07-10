package com.mysite.spring_erp.member.service;

import java.time.LocalDateTime;
import java.util.Optional;

// import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Service;

import com.mysite.spring_erp.exception.DataNotFoundException;
import com.mysite.spring_erp.member.entity.EmpMaster;
import com.mysite.spring_erp.member.repository.MasterRepository;

import lombok.RequiredArgsConstructor;

@RequiredArgsConstructor
@Service
public class MasterService {
    private final MasterRepository empMasterRepository;
    private final PasswordEncoder passwordEncoder;

    // 마지막 사원번호 조회
    public EmpMaster getLatestEmpMaster() {
        Optional<EmpMaster> latestMaster = this.empMasterRepository.findTopByOrderByEmpNoDesc();
        if (latestMaster.isPresent()) {
            return latestMaster.get();
        } else {
            throw new DataNotFoundException("user not found");
        }
    }

    // 회원 정보 저장
    public EmpMaster saveEmpMaster(String name, String pwd) {
        EmpMaster master = new EmpMaster();
        // master.setEmpNo(empNo);

        // BCryptPasswordEncoder passwordEncoder = new BCryptPasswordEncoder();
        // master.setEmpPwd(passwordEncoder.encode(emp_pwd));
        master.setName(name);
        master.setPwd(this.passwordEncoder.encode(pwd));

        master.setCreatedDate(LocalDateTime.now());
        master.setUpdatedDate(LocalDateTime.now());
        this.empMasterRepository.save(master);

        return master;
    }

    // 회원 정보 조회
    public EmpMaster getEmpMaster(String username) {
        Optional<EmpMaster> master = this.empMasterRepository.findByEmpno(username);
        if (master.isPresent()) {
            return master.get();
        } else {
            throw new DataNotFoundException("user not found");
        }
    }
}
