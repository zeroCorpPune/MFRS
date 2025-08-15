package com.rcs.mfrs.Controller;

import java.math.BigDecimal;
import java.text.NumberFormat;
import java.util.Locale;


public class ValueWithDecimalAndString {

    private BigDecimal value;
    private String formattedValue;

    public ValueWithDecimalAndString(BigDecimal value) {
        this.value = value;
        this.formattedValue = formatIndianNumber(value);
    }

    private String formatIndianNumber(BigDecimal number) {
        NumberFormat indianFormat = NumberFormat.getInstance(new Locale("en", "IN"));
        return indianFormat.format(number);
    }

    public BigDecimal getValue() {
        return value;
    }

    public String getFormattedValue() {
        if (value.compareTo(BigDecimal.ZERO) < 0) 
        {
           // System.out.println("Comes here " );  
            BigDecimal minus;
            minus = new BigDecimal(-1);
            this.formattedValue = "(" + formatIndianNumber(value.multiply(minus)) + ")";
            //System.out.println("Comes here for " +  formattedValue );  
        }
        //System.out.println("Did it go in the if condition for " +  formattedValue );  
        //System.out.println("(the ValueWithDecimalAndString  getFormattedValue) = "+ formattedValue);
        return formattedValue;
    }

    public void setValue(BigDecimal value) {
        this.value = value;
        this.formattedValue = formatIndianNumber(value);
    }
}