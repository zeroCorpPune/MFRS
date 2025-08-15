package com.rcs.mfrs.Services;


import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

import com.rcs.mfrs.Model.Mistxn;
import com.rcs.mfrs.Repositories.MistxnRepository;

@Service
public class MistxnServiceImpl implements MistxnService  {
    private final MistxnRepository mistxnRepository;

    @Autowired
    public MistxnServiceImpl(MistxnRepository mistxnRepository) {
        this.mistxnRepository = mistxnRepository;
    }
 
    @Override
    public List<Mistxn> findByCriteria(String ccompcode, String cmonth, String cyear) {
        return mistxnRepository.findByCcompcodeAndCmonthAndCyear( ccompcode, cmonth, cyear,Sort.by(Sort.Direction.ASC, "ccompcode"));
    }

    @Override
    public void saveAll(List<Mistxn> mistxnListWrapper) {

        if (mistxnListWrapper == null || mistxnListWrapper.contains(null)) {
            throw new IllegalArgumentException("Entities must not be null!");
        }

        System.out.println("MistxnServiceImpl->saveAll");
        mistxnRepository.saveAll(mistxnListWrapper);
        System.out.println("MistxnServiceImpl->saveAll Done");
    }

    @Override
    public long countTransactionsByMonthAndYear(String year, String month,String ccompcode) {
        return mistxnRepository.countByMonthAndYear(year, month, ccompcode);
    }


    @Override
    public int flashSubmitStatus(String year, String month,String ccompcode) {
        return mistxnRepository.flashSubmitStatus(year, month, ccompcode);
    }

    @Override
    public int updateMistxnStatus(String year, String month,String ccompcode) {
       return mistxnRepository.updateMistxnStatus (year, month, ccompcode);
    }
}
