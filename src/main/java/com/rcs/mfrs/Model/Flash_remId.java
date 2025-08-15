package com.rcs.mfrs.Model;

import java.io.Serializable;
import java.util.Objects;

public class Flash_remId implements Serializable{

    private String ccompcode;
    private String cmonth;
    private String cyear;

    public Flash_remId() {}

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Flash_remId flash_remId = (Flash_remId) o;
        return Objects.equals(ccompcode, flash_remId.ccompcode) &&
               Objects.equals(cmonth, flash_remId.cmonth) &&
               Objects.equals(cyear, flash_remId.cyear) ;
    }

    @Override
    public int hashCode() {
        return Objects.hash(ccompcode, cmonth, cyear);
    }
}
