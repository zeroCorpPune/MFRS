package com.rcs.mfrs.Services;

import java.util.List;
import java.util.Optional;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.rcs.mfrs.Model.Saftxn;
import com.rcs.mfrs.Repositories.SafRepository;



@Service
public class SafService {
    
    private final SafRepository safRepository;
//private final MistxnRepository mistxnRepository;
    
    @Autowired
    public SafService(SafRepository safRepository) {
        this.safRepository = safRepository;
    }
    public int safSubmitStatus(String year, String month,String ccompcode) {
        return safRepository.safSubmitStatus(year, month, ccompcode);
    }
    
    

    
    public Optional<Saftxn> findByCcompcodeAndCmonthAndCyear(String ccompcode, String cmonth, String cyear) {
        return safRepository.findByCcompcodeAndCmonthAndCyear( ccompcode, cmonth, cyear);
    }


    public Saftxn save(Saftxn saftxn){
        /* 
        if (company.getId() == null){
            company.setCreatedAt(LocalDateTime.now());
        }
        company.setUpdatedAt(LocalDateTime.now());
        */
        
        return safRepository.save(saftxn);
    }
}
