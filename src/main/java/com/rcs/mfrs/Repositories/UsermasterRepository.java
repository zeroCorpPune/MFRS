package com.rcs.mfrs.Repositories;

import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;
//import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;
import com.rcs.mfrs.Model.Usermaster;

@Repository
public interface UsermasterRepository extends JpaRepository<Usermaster, String>{
    Optional<Usermaster>  findByCuserid(String cuserid);

    
    List<Usermaster> findByCuseridContaining(String cuserid);

    
    

}
