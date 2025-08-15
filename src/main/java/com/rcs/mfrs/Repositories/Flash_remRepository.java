package com.rcs.mfrs.Repositories;

import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.rcs.mfrs.Model.Flash_rem;
import com.rcs.mfrs.Model.Flash_remId;



@Repository
public interface Flash_remRepository extends JpaRepository<Flash_rem, Flash_remId>{
    
    Optional<Flash_rem> findByCcompcodeAndCmonthAndCyear(String ccompcode, String cmonth, String cyear);
    

    
}
