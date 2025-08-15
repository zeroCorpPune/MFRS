package com.rcs.mfrs.Model;

import java.io.Serializable;
import java.util.Objects;

public class MistxnId implements Serializable {

    private String ccompcode;
    private String cmonth;
    private String cyear;
    private String desccd;

    // Default constructor
    public MistxnId() {}

    // Getters and Setters

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        MistxnId mistxnId = (MistxnId) o;
        return Objects.equals(ccompcode, mistxnId.ccompcode) &&
               Objects.equals(cmonth, mistxnId.cmonth) &&
               Objects.equals(cyear, mistxnId.cyear) &&
               Objects.equals(desccd, mistxnId.desccd);
    }

    @Override
    public int hashCode() {
        return Objects.hash(ccompcode, cmonth, cyear, desccd);
    }
}