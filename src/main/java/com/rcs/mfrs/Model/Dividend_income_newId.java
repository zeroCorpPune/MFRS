package com.rcs.mfrs.Model;

import java.io.Serializable;
import java.util.Objects;

public class Dividend_income_newId implements Serializable {
    private String ccompcode;
    private String cmonth;
    private String cyear;
    private String description;

    // Default constructor
    public Dividend_income_newId(){}
    
    public Dividend_income_newId(String ccompcode, String cyear, String cmonth, String description) {
        this.ccompcode = ccompcode;
        this.cyear = cyear;
        this.cmonth = cmonth;
        this.description = description;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Dividend_income_newId dividend_income_newIdId = (Dividend_income_newId) o;
        return Objects.equals(ccompcode, dividend_income_newIdId.ccompcode) &&
               Objects.equals(cmonth, dividend_income_newIdId.cmonth) &&
               Objects.equals(cyear, dividend_income_newIdId.cyear) &&
               Objects.equals(description, dividend_income_newIdId.description);
    }

    @Override
    public int hashCode() {
        return Objects.hash(ccompcode, cmonth, cyear, description);
    }

}
