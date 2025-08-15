package com.rcs.mfrs.Repositories;

import java.util.List;
import java.util.Optional;

import org.springframework.data.domain.Sort;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.rcs.mfrs.Model.Company;

@Repository
public interface CompanyRepository extends JpaRepository<Company, String>{
    Optional<Company> findByCcompcode(String ccompcode);

    Optional<Company> findByUsermasters(String usermasters);

    List<Company> findAll (Sort sort);
    
}
