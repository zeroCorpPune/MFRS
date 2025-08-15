package com.rcs.mfrs.Model;

import java.time.LocalDate;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Entity
@Getter
@Setter
@NoArgsConstructor
public class Safcard {
    @Id
    private String cash_profit;

     @Column(name = "net_profit_loss")
    private String net_profit_loss_crd;
    private String add_depreciation;
    private String add_amortisation;
    private String add_prov_for_impairment;
    private String less_profit_on_si;
    private String less_profit_on_sfa;
    private String total_cash_profit;
    private String decr_in_wc;
    private String inventory_decr;
    private String receivable_decr;
    private String other_ca_decr;
    private String other_cr_incr;
    private String other_cl_incr;
    private String total_decr_in_wc;
    private String incr_in_borrowings;
    private String recpt_of_term_loan;
    private String incr_in_od_ltr_bd;
    private String decr_in_cash_bank;
    private String total_incr_in_borrowings;
    private String loan_fm_ogcs;
    private String loan_fm_othr;
    private String sale_proceeds_of_fa;
    private String sale_proceeds_of_invst;
    private String capital_contribution;
    private String proprietor_contribution;
    private String total_sources;
    private String cash_loss;
    private String net_loss_profit;
    private String less_depreciation;
    private String less_amortisation;
    private String less_prov_for_impairment;
    private String add_profit_on_si;
    private String add_profit_on_sfa;
    private String total_cash_loss;
    private String incr_in_wc;
    private String inventory_incr;
    private String receivable_incr;
    private String other_ca_incr;
    private String other_cr_decr;
    private String other_cl_decr;
    private String total_incr_in_wc;
    private String decr_in_borrowings;
    private String repymt_of_term_loan;
    private String decr_in_od_ltr_bd;
    private String incr_in_cash_bank;
    private String total_decr_in_borrowings;
    private String loan_to_ogcs;
    private String loan_to_othr;
    private String addn_to_fa;
    private String purchase_of_invst;
    private String divnd_paid;
    private String proprietor_withdrawal;
    private String total_applications;
    private LocalDate dcreationdate;	
    private LocalDate dmodifieddate;
}
