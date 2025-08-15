package com.rcs.mfrs.Services;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.rcs.mfrs.Model.Safcard;
import com.rcs.mfrs.Repositories.SafcardRepository;

@Service
public class SafcardService {
    @Autowired
    private SafcardRepository safcardRepository;

    public List<Safcard> findAll(){
        return safcardRepository.findAll();
    }
}
