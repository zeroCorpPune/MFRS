package com.rcs.mfrs.Model;

import javax.validation.constraints.NotNull;
import javax.validation.constraints.Size;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.validation.constraints.DecimalMax;
import javax.validation.constraints.DecimalMin;
import java.math.BigDecimal;
import java.text.DecimalFormat;

import java.time.LocalDateTime;

import javax.persistence.Id;
import javax.persistence.IdClass;
import javax.persistence.Entity;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Entity
@Getter
@Setter
@IdClass(SaftxnId.class)
@NoArgsConstructor
public class Saftxn {
    @Id
    private String ccompcode;

    @Id
    private String cmonth;

    @Id
    private String cyear;

    
    private String cash_profit ;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal net_profit_loss;

    private BigDecimal add_depreciation  ;
    private BigDecimal add_amortisation  ;
    private BigDecimal add_prov_for_impairment  ;
    private BigDecimal less_profit_on_si  ;
    private BigDecimal less_profit_on_sfa  ;
    private BigDecimal total_cash_profit  ;
    private String decr_in_wc  ;
    private BigDecimal inventory_decr  ;
    private BigDecimal receivable_decr  ;
    private BigDecimal other_ca_decr  ;
    private BigDecimal other_cr_incr  ;
    private BigDecimal other_cl_incr  ;
    private BigDecimal total_decr_in_wc  ;
    private String incr_in_borrowings  ;
    private BigDecimal recpt_of_term_loan ; 
    private BigDecimal incr_in_od_ltr_bd  ;
    private BigDecimal decr_in_cash_bank  ;
    private BigDecimal total_incr_in_borrowings  ;
    private BigDecimal loan_fm_ogcs  ;
    private BigDecimal loan_fm_othr  ;
    private BigDecimal sale_proceeds_of_fa  ;
    private BigDecimal sale_proceeds_of_invst  ;
    private BigDecimal capital_contribution  ;
    private BigDecimal proprietor_contribution ; 
    private BigDecimal total_sources  ;
    private String cash_loss;
    private BigDecimal net_loss_profit  ;
    private BigDecimal less_depreciation  ;
    private BigDecimal less_amortisation  ;
    private BigDecimal less_prov_for_impairment  ;
    private BigDecimal add_profit_on_si  ;
    private BigDecimal add_profit_on_sfa ;
    private BigDecimal total_cash_loss  ;
    private String incr_in_wc  ;
    private BigDecimal inventory_incr  ;
    private BigDecimal receivable_incr  ;
    private BigDecimal other_ca_incr  ;
    private BigDecimal other_cr_decr  ;
    private BigDecimal other_cl_decr  ;
    private BigDecimal total_incr_in_wc  ;
    private String decr_in_borrowings  ;
    private BigDecimal repymt_of_term_loan  ;
    private BigDecimal decr_in_od_ltr_bd  ;
    private BigDecimal incr_in_cash_bank  ;
    private BigDecimal total_decr_in_borrowings  ;
    private BigDecimal loan_to_ogcs  ;
    private BigDecimal loan_to_othr  ;
    private BigDecimal addn_to_fa  ;
    private BigDecimal purchase_of_invst  ;
    private BigDecimal divnd_paid  ;
    private BigDecimal proprietor_withdrawal  ;
    private BigDecimal total_applications  ;
    private String remarks  ;
    private String cblock  ;
    private LocalDateTime dcreationdate ;
    private LocalDateTime dmodifieddate ;

    @ManyToOne
    @JoinColumn(name = "ccompcode", insertable = false, updatable = false, nullable = false)
    private Company saftxncompany;
}
