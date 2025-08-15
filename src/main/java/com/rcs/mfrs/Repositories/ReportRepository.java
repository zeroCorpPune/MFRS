package com.rcs.mfrs.Repositories;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.rcs.mfrs.Model.Mistxn;
import com.rcs.mfrs.Model.MistxnId;

@Repository
public interface ReportRepository extends JpaRepository<Mistxn, MistxnId>{
    // NOT USED THIS REPOSITORY SO FAR
}
