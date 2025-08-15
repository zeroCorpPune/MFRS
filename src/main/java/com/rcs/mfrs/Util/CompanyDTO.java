package com.rcs.mfrs.Util;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class CompanyDTO {
    private String ccompcode;
    private String ccompname;

    public CompanyDTO(String ccompcode, String ccompname) {
        this.ccompcode = ccompcode;
        this.ccompname = ccompname;
    }

    // Getters and setters
}