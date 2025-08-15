package com.rcs.mfrs.Controller;

import java.util.List;

import com.rcs.mfrs.Model.Mistxn;

public class MistxnListWrapper {
     private List<Mistxn> mistxns;

    public List<Mistxn> getMistxns() {
        //System.out.println("getMistxns");
        return mistxns;
    }

    public void setMistxns(List<Mistxn> mistxns) {
        //System.out.println("setMistxns");
        this.mistxns = mistxns;
    }
}
