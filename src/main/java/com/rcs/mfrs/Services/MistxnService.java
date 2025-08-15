package com.rcs.mfrs.Services;
import java.util.List;

import org.springframework.stereotype.Service;

import com.rcs.mfrs.Model.Mistxn;


//import java.util.Collections;

@Service
public interface MistxnService 
{


    
    List<Mistxn> findByCriteria(String ccompcode, String cmonth, String cyear);
    void saveAll(List<Mistxn> mistxns);

    long countTransactionsByMonthAndYear(String year, String month, String ccompcode);
    
    int flashSubmitStatus(String year, String month, String ccompcode);
    
    int updateMistxnStatus (String year, String month, String ccompcode);
    
}
