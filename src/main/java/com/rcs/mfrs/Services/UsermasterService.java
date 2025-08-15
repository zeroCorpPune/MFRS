package com.rcs.mfrs.Services;

import java.util.List;
import java.util.Optional;

import javax.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.authority.SimpleGrantedAuthority;
import org.springframework.security.core.userdetails.User;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
//import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.stereotype.Service;

import com.rcs.mfrs.Model.Company;
//import com.rcs.mfrs.Model.Company;
import com.rcs.mfrs.Model.Usermaster;
import com.rcs.mfrs.Repositories.UsermasterRepository;
import com.rcs.mfrs.Security.CustomUserPrincipal;

import java.util.ArrayList;

import javax.servlet.http.HttpSession;
//import java.util.Collections;
import org.springframework.beans.factory.ObjectFactory;

@Service
public class UsermasterService implements UserDetailsService{
    @Autowired
    private UsermasterRepository usermasterRepository;
    
    //@Autowired
    //private PasswordEncoder passwordEncoder;

    @Autowired
    ObjectFactory<HttpSession> httpSessionFactory;

    public List<Usermaster> getAll(){
        return usermasterRepository.findAll();
    }

    public Optional<Usermaster> getbyCuserid(String cuserid) {
        return usermasterRepository.findByCuserid(cuserid) ;
    }

    public List<Usermaster> getEntitiesBypara(String cuserid) {
        return usermasterRepository.findByCuseridContaining(cuserid);
    }
    
    //public List<Usermaster> getRightsBypara(String cuserid) {
    //    return usermasterRepository.findByCuseridContaining(cuserid);
    //}
    /* 
    public Usermaster save(Usermaster usermaster){
        usermaster.setPassword(passwordEncoder.encode(usermaster.getPassword()));
        return usermasterRepository.save(usermaster);    
    }
    */
   /*
    @Autowired
    public UsermasterService(UsermasterRepository userRepository) {
        this.usermasterRepository = userRepository;
    }
    */
    
    @Override
    public UserDetails loadUserByUsername(String cuserid) throws UsernameNotFoundException {
        System.out.println("username is =============" + cuserid);
        Optional<Usermaster>  user = usermasterRepository.findByCuserid(cuserid);
        
        if (user == null) {
            throw new UsernameNotFoundException("User not found");
        }

        Usermaster account = user.get();
        
       
        List<GrantedAuthority> grantedAuthority = new ArrayList<>();
        grantedAuthority.add(new SimpleGrantedAuthority(account.getUserrole()));

        HttpSession session = httpSessionFactory.getObject();
        session.setAttribute("gblCompanyCode", account.getCompany().getCcompcode());
        System.out.println("password is " + account.getPassword());
        System.out.println("Role is " + account.getUserrole());
        System.out.println("Company name is " + account.getCompany().getCcompname());
        System.out.println("company code is " + account.getCompany().getCcompcode());
        
        // return new User(account.getCuserid(), account.getPassword(),grantedAuthority);
        return new CustomUserPrincipal(account.getCuserid(),
        account.getPassword(),grantedAuthority,
        account.getCanadd(),
        account.getCanmodify(),
        account.getCandelete(), 
        account.getCompany().getCcompname());
        
        
    }
}
