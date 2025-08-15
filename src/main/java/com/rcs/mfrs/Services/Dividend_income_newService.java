package com.rcs.mfrs.Services;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.rcs.mfrs.Model.Dividend_income_new;
import com.rcs.mfrs.Model.Dividend_income_newId;
import com.rcs.mfrs.Repositories.Dividend_income_newRepository;

@Service
public class Dividend_income_newService {

    @Autowired
    private Dividend_income_newRepository dividend_income_newRepository;

    public Dividend_income_newService(Dividend_income_newRepository dividend_income_newRepository) {
        this.dividend_income_newRepository = dividend_income_newRepository;
    }

    public void saveAll(List<Dividend_income_new> dividendIncomes) {
        dividend_income_newRepository.saveAll(dividendIncomes);
    }

    public List<Dividend_income_new> findByCcompcodeAndCyearAndCMonth(String ccompcode, String cyear, String cmonth) {
        //return dividend_income_newRepository.findAll();
        //Dividend_income_newId primaryKey = new Dividend_income_newId(ccompcode, cyear, cmonth, description);
        return dividend_income_newRepository.findByCcompcodeAndCyearAndCmonth(ccompcode,  cyear, cmonth);
    }

    public List<Dividend_income_new> getAllDividendIncomeRecords() {
        return dividend_income_newRepository.findAll();
    }

    public void deleteByPrimaryKey(String ccompcode, String cyear, String cmonth, String description) {
    Dividend_income_newId primaryKey = new Dividend_income_newId(ccompcode, cyear, cmonth, description);
    dividend_income_newRepository.deleteById(primaryKey);
    }
}
