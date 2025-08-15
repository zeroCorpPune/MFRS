package com.rcs.mfrs.Model;
import java.math.BigDecimal;
import java.time.LocalDateTime;

import javax.persistence.Id;
import javax.persistence.IdClass;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.validation.constraints.DecimalMax;
import javax.validation.constraints.DecimalMin;
import javax.persistence.Entity;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Entity
@Getter
@Setter
@IdClass(Loans_from_toId.class)
@NoArgsConstructor
public class Loans_from_to {
    @Id
    private String ccompcode;

    @Id
    private String cmonth;

    @Id
    private String cyear;
    
    @Id
    private String ccompname;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal loan_from  ;
    
    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal loan_to ;
    
    
    private String remarks ; 
    private char cblock ;
    LocalDateTime  dcreationdate ;
    LocalDateTime dmodifieddate ;

    @ManyToOne
    @JoinColumn(name = "ccompcode", insertable = false, updatable = false, nullable = false)
    private Company loans_from_tocompany;
}
