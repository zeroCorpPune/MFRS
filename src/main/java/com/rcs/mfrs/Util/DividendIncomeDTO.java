package com.rcs.mfrs.Util;

import lombok.Getter;
import lombok.Setter;

import java.math.BigDecimal;
import java.time.LocalDate;

@Getter
@Setter
public class DividendIncomeDTO {
    private String ccompcode;
    private String cmonth;
    private String cyear;
    private String desccd;
    private String boldyn;
    private String description;
    private BigDecimal actMtdBal;
    private BigDecimal bgtMtdBal;
    private BigDecimal actMtdBalPy1;
    private BigDecimal actYtdBal;
    private BigDecimal bgtYtdBal;
    private BigDecimal actYtdBalPy1;
    private BigDecimal actYtdBalPy2;
    private BigDecimal actYtdBalPy3;
    private BigDecimal actYtdBalPy4;
    private String remarks;
    private String cblock;
    private String cactive;
    private LocalDate dcreationdate;
    private LocalDate dmodifieddate;
}
