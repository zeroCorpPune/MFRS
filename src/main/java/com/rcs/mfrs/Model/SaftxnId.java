package com.rcs.mfrs.Model;

import java.io.Serializable;
import java.util.Objects;


public class SaftxnId implements Serializable {
    private String ccompcode;
    private String cmonth;
    private String cyear;
    

    // Default constructor
    public SaftxnId() {}

    // Getters and Setters

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        SaftxnId saftxnId = (SaftxnId) o;
        return Objects.equals(ccompcode, saftxnId.ccompcode) &&
               Objects.equals(cmonth, saftxnId.cmonth) &&
               Objects.equals(cyear, saftxnId.cyear) ;
    }

    @Override
    public int hashCode() {
        return Objects.hash(ccompcode, cmonth, cyear);
    }
}
