package com.rcs.mfrs.Util;

public enum Privillages {
    /* 
    RESET_ANY_USER_PASSWORD(1l, "RESET_ANY_USER_PASSWORD"),
    ACCESS_ADMIN_PANEL(2l, "ACCESS_ADMIN_PANEL");
    */

    ACCESS_COMPANY(1l, "ACCESS_COMPANY_LIST"),
    ACCESS_USER(2l, "ACCESS_USER_LIST");

    private Long id;
    private String privillage;
    private Privillages(Long id, String privillage) {
        this.id = id;
        this.privillage = privillage;
    }
    public Long getId() {
        return id;
    }
    public String getPrivillage() {
        return privillage;
    }
    
}
