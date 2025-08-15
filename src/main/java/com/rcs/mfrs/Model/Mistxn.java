package com.rcs.mfrs.Model;

import javax.validation.constraints.NotNull;
import javax.validation.constraints.Size;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.validation.constraints.DecimalMax;
import javax.validation.constraints.DecimalMin;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.time.LocalDate;
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
@IdClass(MistxnId.class)
@NoArgsConstructor
public class Mistxn {
    @Id
    private String ccompcode;

    @Id
    private String cmonth;

    @Id
    private String cyear;

    @Id
    private String desccd;

    @Size(max = 1)
    private String boldyn;

    @Size(max = 30)
    private String description;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal actMtdBal;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal bgtMtdBal;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal actMtdBalPy1;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal actYtdBal;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal bgtYtdBal;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal actYtdBalPy1;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal actYtdBalPy2;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal actYtdBalPy3;

    @DecimalMin(value = "-9999999999.999", message = "Value is too small")
    @DecimalMax(value = "9999999999.999", message = "Value is too large")
    private BigDecimal actYtdBalPy4;

    @Size(max = 100)
    private String remarks;

    @Size(max = 1)
    private String cblock;

    @Size(max = 1)
    private String cactive;

    private LocalDate dcreationdate;
    private LocalDate dmodifieddate;

    @ManyToOne
    @JoinColumn(name = "ccompcode", insertable = false, updatable = false, nullable = false)
    private Company mistxncompany;
/*
    
    public String getActMtdBalPy1() {
        if (actMtdBalPy1 == null) {
            return ""; // Return an empty string if the value is null
        }

        // Define the format pattern (e.g., two decimal places, comma as thousand separator)
        DecimalFormat decimalFormat = new DecimalFormat("#,##0.00");

        // Apply the formatting
        return decimalFormat.format(actMtdBalPy1);
    }

    // Setter for actMtdBalPy1
    public void setActMtdBalPy1(BigDecimal actMtdBalPy1) {
        this.actMtdBalPy1 = actMtdBalPy1;
    }

    public String getActYtdBal() {
        if (actYtdBal == null) {
            return ""; // Return an empty string if the value is null
        }

        // Define the format pattern (e.g., two decimal places, comma as thousand separator)
        DecimalFormat decimalFormat = new DecimalFormat("#,##0.00");

        // Apply the formatting
        return decimalFormat.format(actYtdBal);
    }

    // Setter for actMtdBalPy1
    public void setActYtdBal(BigDecimal actYtdBal) {
        this.actYtdBal = actYtdBal;
    }
         */
}
