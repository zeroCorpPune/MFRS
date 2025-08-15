package com.rcs.mfrs.Repositories;

import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.rcs.mfrs.Model.Dividend_income_new;
import com.rcs.mfrs.Model.Dividend_income_newId;

@Repository
public interface Dividend_income_newRepository extends JpaRepository<Dividend_income_new, Dividend_income_newId> {
    
    List<Dividend_income_new> findByCcompcodeAndCyearAndCmonth(String ccompcode,  String cyear, String cmonth);
    
}
