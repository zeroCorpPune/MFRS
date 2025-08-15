package com.rcs.mfrs.Security;

import java.util.Collection;

import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.userdetails.User;

public class CustomUserPrincipal extends User {
    
    private String canModify;
    private String canAdd;
    private String canDelete;
    private String company;
    

    
    public CustomUserPrincipal(String username, String password, Collection<? extends GrantedAuthority> authorities, String canAdd, String canModify, String canDelete, String companynamme) {
        super(username, password, authorities);
        this.canAdd = canAdd;
        this.canModify = canModify;
        this.canDelete = canDelete;
        this.company = companynamme;
        
    }

    

    public String getCompany() {
        return company;
    }

    public String getCanAdd() {
        return canAdd;
    }

    public String getCanModify() {
        return canModify;
    }

    public String getCanDelete() {
        return canDelete;
    }
}
