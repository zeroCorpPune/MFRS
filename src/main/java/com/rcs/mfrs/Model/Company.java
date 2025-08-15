package com.rcs.mfrs.Model;

import java.math.BigDecimal;
import java.time.LocalDate;
//import java.util.List;
import java.util.Set;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.OneToMany;
import javax.validation.constraints.DecimalMax;
import javax.validation.constraints.DecimalMin;
import javax.validation.constraints.NotEmpty;
import javax.validation.constraints.Size;

import com.fasterxml.jackson.annotation.JsonIgnore;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Entity
@Getter
@Setter
@NoArgsConstructor
public class Company {
    
    @Id
    @Column(insertable=false, unique = true,updatable=false, nullable=false)
    @NotEmpty(message = "Company code cannot be blank")
    private String  ccompcode;

    @NotEmpty(message = "Company Name cannot be blank")
    private String  ccompname;

    @Size(max = 25)
    private String CATEGORY;

    @Size(max = 100)
    private String CEMAIL ;

    @Size(max = 1)
    private String CDEFAULT;

    @Size(max = 50)
    private String CCONTACTPERSON;

    @Size(max = 50)
    private String CADDRESS1 ;

    @Size(max = 50)
    private String CADDRESS2 ;

    @Size(max = 50)
    private String CADDRESS3 ;

    @Size(max = 50)
    private String CADDRESS4 ;

    @Size(max = 50)
    private String CADDRESS5 ;

    @Size(max = 50)
    private String CPHONE ;

    @Size(max = 50)
    private String CFAX ;

    @Size(max = 15)
    private String CUSER_ID ;

    @Size(max = 50)
    private String CUSER_NAME;

    @Size(max = 10)
    private String CUSER_TYPE;

    @Size(max = 16)
    private String PASSWORD;

    @Size(max = 1)
    private String CACTIVE ;

    
    private LocalDate DCREATIONDATE ;

    
    private LocalDate DMODIFIEDDATE ;

    @Size(max = 75)
    private String CCOMPNAMESN ;

    @Size(max = 75)
    private String CCOMPNAMESN1 ;

    @Size(max = 1)
    private String SAFYN ;

    @Size(max = 1)
    private String RECPAYYN ;

    @Size(max = 1)
    private String FINPOSYN ;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal STAKE ;

    @Size(max = 1)
    private String TRADERECYN ;

    @Size(max = 1)
    private String BANKBORROWINGSYN ;

    @Size(max = 10)
    private String CCOMPCODE_OLD ;

    @Size(max = 10)
    private String CCOMPCODE_NEW ;

    @Size(max = 1)
    private String WCYN ;

    @Size(max = 2)
    private String SECTORNO ;

    @Size(max = 2)
    private String MISYN ;

    @Size(max = 2)
    private String BUDGETYN ;

    @Size(max = 200)
    private String BUDGET_CEMAIL ;

    @Size(max = 2)
    private String MFRSYN ;

    @Size(max = 75)
    private String CCOMPNAMESN2 ;

    @Size(max = 1)
    private String UPLOADFLASHFROMEXCEL ;

    @Size(max = 1)
    private String INCLUDEINFLASHREPORT ;

    @Size(max = 1)
    private String INCLUDEINBBREPORT ;

    @Size(max = 1)
    private String SECTIONINBBREPORT ;

    @Size(max = 2)
    private String FLEXI_CCOMPCODE ;

    @Size(max = 20)
    private String COMREGNO ;

    @OneToMany(mappedBy = "company")
    private Set<Usermaster> usermasters;
    //private String Set<Usermaster> usermasters;
    
    @OneToMany(mappedBy = "mistxncompany")
    @JsonIgnore
    private Set<Mistxn> mistxn;

    @OneToMany(mappedBy = "dividend_income_newcompany")
    private Set<Dividend_income_new> dividend_income_new;  

    @OneToMany(mappedBy = "saftxncompany")
    @JsonIgnore
    private Set<Saftxn> saftxn;

    @OneToMany(mappedBy = "loans_from_tocompany") 
    @JsonIgnore
    private Set<Loans_from_to> loans_from_tos;

}
