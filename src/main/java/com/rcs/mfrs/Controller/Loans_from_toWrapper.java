package com.rcs.mfrs.Controller;

import java.util.ArrayList;
import java.util.List;

import com.rcs.mfrs.Model.Loans_from_to;


public class Loans_from_toWrapper {
    private List<Loans_from_to> loans_from_toList = new ArrayList<>();

    public List<Loans_from_to> getLoans_from_toList() {
        return loans_from_toList;
    }

    public void setLoans_from_toList(List<Loans_from_to> loans_from_toList) {
        this.loans_from_toList = loans_from_toList;
    }
}
