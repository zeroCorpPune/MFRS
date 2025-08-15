package com.rcs.mfrs.Services;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.rcs.mfrs.Model.Loans_from_to;
import com.rcs.mfrs.Model.Loans_from_toId;
import com.rcs.mfrs.Repositories.Loans_from_toRepository;

@Service
public class Loans_from_toService {
    @Autowired
    private Loans_from_toRepository loans_from_toRepository;

    public Loans_from_toService(Loans_from_toRepository loans_from_toRepository) {
        this.loans_from_toRepository = loans_from_toRepository;
    }

    public void saveAll(List<Loans_from_to> loans_from_tos) {
        loans_from_toRepository.saveAll(loans_from_tos);
    }

    public List<Loans_from_to> findByCcompcodeAndCyearAndCMonth(String ccompcode, String cyear, String cmonth) {
        //return dividend_income_newRepository.findAll();
        //Dividend_income_newId primaryKey = new Dividend_income_newId(ccompcode, cyear, cmonth, description);
        return loans_from_toRepository.findByCcompcodeAndCyearAndCmonth(ccompcode,  cyear, cmonth);
    }

    public List<Loans_from_to> getAllDividendIncomeRecords() {
        return loans_from_toRepository.findAll();
    }

    public void deleteByPrimaryKey(String ccompcode, String cyear, String cmonth, String ccompname) {
        Loans_from_toId primaryKey = new Loans_from_toId(ccompcode, cyear, cmonth, ccompname);
        loans_from_toRepository.deleteById(primaryKey);
    }
}
