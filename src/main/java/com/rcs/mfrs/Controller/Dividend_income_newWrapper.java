package com.rcs.mfrs.Controller;
import java.util.ArrayList;
import java.util.List;

import com.rcs.mfrs.Model.Dividend_income_new;

public class Dividend_income_newWrapper {
    private List<Dividend_income_new> dividendIncomeList = new ArrayList<>();

    public List<Dividend_income_new> getDividendIncomeList() {
        return dividendIncomeList;
    }

    public void setDividendIncomeList(List<Dividend_income_new> dividendIncomeList) {
        this.dividendIncomeList = dividendIncomeList;
    }
}
