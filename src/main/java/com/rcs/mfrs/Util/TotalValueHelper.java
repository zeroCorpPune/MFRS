package com.rcs.mfrs.Util;

import java.math.BigDecimal;
import java.text.NumberFormat;
import java.util.Locale;

public class TotalValueHelper {
    
    private BigDecimal totalValue = BigDecimal.ZERO;  // Initialize to 0

    private String formattedValue;

    public BigDecimal getTotalValue() {
       // return totalValue;
       if (totalValue.compareTo(BigDecimal.ZERO) < 0) 
        {
            BigDecimal minus;
            minus = new BigDecimal(-1);
            this.formattedValue = "(" + formatIndianNumber(totalValue.multiply(minus)) + ")";
        }

        return totalValue;

    }


    public String getFormattedValue() {
        if (totalValue.compareTo(BigDecimal.ZERO) < 0) 
        {
            BigDecimal minus;
            minus = new BigDecimal(-1);
            formattedValue = "(" + formatIndianNumber(totalValue.multiply(minus)) + ")";
            
        }
        else
        formattedValue =  formatIndianNumber(totalValue) ;

           // System.out.println("(the TotalValueHelper  getFormattedValue) = "+ formattedValue);
            return formattedValue;
    }

    private String formatIndianNumber(BigDecimal number) {
        NumberFormat indianFormat = NumberFormat.getInstance(new Locale("en", "IN"));
        return indianFormat.format(number);
    }

    public void addToTotal(BigDecimal value) {
        if (value != null) {
            totalValue = totalValue.add(value);
            this.formattedValue = formatIndianNumber(value);
        }
    }


}
 
/*
 import java.math.BigDecimal;

public class TotalValueHelper {
    
    private BigDecimal totalValue = BigDecimal.ZERO;  // Initialize to 0

    

    public BigDecimal getTotalValue() {
        return totalValue;

    }



    public void addToTotal(BigDecimal value) {
        if (value != null) {
            totalValue = totalValue.add(value);

        }
    }


}
    */