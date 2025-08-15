package com.rcs.mfrs.Services;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Optional;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.rcs.mfrs.Model.Flash_rem;
import com.rcs.mfrs.Repositories.Flash_remRepository;


@Service
public class Flash_remService {
    @Autowired
    private Flash_remRepository flash_remRepository;

    @Autowired
    public Flash_remService(Flash_remRepository flash_remRepository) {
        this.flash_remRepository = flash_remRepository;
    }

  
    
    public Optional<Flash_rem> findByCriteria(String ccompcode, String cmonth, String cyear) {
        System.out.println("fetching remark " ) ;     

        return flash_remRepository.findByCcompcodeAndCmonthAndCyear(ccompcode, cmonth, cyear);
    }

     public Flash_rem save(Flash_rem flash_rem){
        /* 
        if (company.getId() == null){
            company.setCreatedAt(LocalDateTime.now());
        }
        company.setUpdatedAt(LocalDateTime.now());
        */
        flash_rem.setDmodifieddate(LocalDate.now()); // Set the date
        return flash_remRepository.save(flash_rem);
    }
}
