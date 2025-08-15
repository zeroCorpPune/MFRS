package com.rcs.mfrs.Repositories;

import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.rcs.mfrs.Model.Safcard;


@Repository
public interface  SafcardRepository extends JpaRepository<Safcard,String>{
    List<Safcard> findAll ();
}
