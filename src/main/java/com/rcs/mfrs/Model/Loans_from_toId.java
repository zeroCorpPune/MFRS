package com.rcs.mfrs.Model;

import java.io.Serializable;
import java.util.Objects;


public class Loans_from_toId implements Serializable {
    private String ccompcode;
    private String cmonth;
    private String cyear;
    private String ccompname;
    
    // Default constructor
    public Loans_from_toId(){}

    // Default constructor
    public Loans_from_toId(String ccompcode, String cyear, String cmonth, String ccompname) {
        this.ccompcode = ccompcode;
        this.cyear = cyear;
        this.cmonth = cmonth;
        this.ccompname = ccompname;
    }

    // Getters and Setters

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Loans_from_toId loans_from_toId = (Loans_from_toId) o;
        return Objects.equals(ccompcode, loans_from_toId.ccompcode) &&
               Objects.equals(cmonth, loans_from_toId.cmonth) &&
               Objects.equals(cyear, loans_from_toId.cyear) &&
               Objects.equals(ccompname, loans_from_toId.ccompname);
    }

    @Override
    public int hashCode() {
        return Objects.hash(ccompcode, cmonth, cyear, ccompname);
    }
}
