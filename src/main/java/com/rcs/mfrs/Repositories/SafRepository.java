package com.rcs.mfrs.Repositories;

import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import com.rcs.mfrs.Model.Saftxn;
import com.rcs.mfrs.Model.SaftxnId;

@Repository
public interface SafRepository extends JpaRepository<Saftxn, SaftxnId>{

    Optional<Saftxn> findByCcompcodeAndCmonthAndCyear(String ccompcode, String cmonth, String cyear);
    
    @Query("SELECT COUNT(m) FROM Saftxn m WHERE cblock='Y' and  m.cyear = :year AND m.cmonth = :month AND m.ccompcode = :ccompcode")
    int safSubmitStatus(@Param("year") String year, @Param("month") String month, @Param("ccompcode") String ccompcode);
    

}
