package com.rcs.mfrs.Repositories;

import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.rcs.mfrs.Model.Loans_from_to;
import com.rcs.mfrs.Model.Loans_from_toId;

@Repository
public interface Loans_from_toRepository extends JpaRepository<Loans_from_to, Loans_from_toId> {
    List<Loans_from_to> findByCcompcodeAndCyearAndCmonth(String ccompcode,  String cyear, String cmonth);
}
