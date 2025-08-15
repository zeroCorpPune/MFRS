package com.rcs.mfrs.Controller;

import java.math.BigDecimal;
import java.security.Principal;
import java.time.LocalDate;
import java.time.Month;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javax.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.access.prepost.PreAuthorize;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;

import com.rcs.mfrs.Model.Saftxn;
import com.rcs.mfrs.Model.SaftxnId;
import com.rcs.mfrs.Model.Company;
import com.rcs.mfrs.Model.Flash_rem;
import com.rcs.mfrs.Model.Mistxn;
import com.rcs.mfrs.Model.Safcard;
import com.rcs.mfrs.Security.CustomUserPrincipal;
import com.rcs.mfrs.Services.CompanyService;
import com.rcs.mfrs.Services.SafService;
import com.rcs.mfrs.Services.SafcardService;
import com.rcs.mfrs.Util.MCommon;

@Controller
public class SafController {

    @Autowired
    private CompanyService companyService;

    @Autowired
    private SafService safService;

    @Autowired
    private SafcardService safcardService;

    private final MCommon mcomm;

    @Autowired
    public SafController(MCommon mcomm) {
        this.mcomm = mcomm;
    }

    @PostMapping("/safcancel")
    public String cancel(Model model, Principal principal,HttpSession session) {
         return "redirect:/safmenu";
    }

    @GetMapping("/safmenu")
    public String safmenu( 
        @RequestParam(value = "month", required = false) String month,
        @RequestParam(value = "year", required = false) String year,
        String yearvalue,Model model, Principal principal,HttpSession session) {
        System.out.println("mistxnmenu");
        
        if (month == null) {
        // Get the current month in two-digit format (e.g., "09" for September)
        month = String.format("%02d", LocalDate.now().getMonthValue());
        }
        
        if (year == null) {
            // Get the current month in two-digit format (e.g., "09" for September)
            year =  String.format("%02d", LocalDate.now().getYear());
            }
        System.out.println("default month = " + month);
        System.out.println("default year = " + year);

        List<Integer> years = mcomm.getYearsList(10);
        Map<String, String> months = mcomm.getMonthsList();

        int currentYear = mcomm.getCurrentYear();
        String currentMonth = mcomm.getCurrentMonth();

        model.addAttribute("years", years);
        model.addAttribute("months", months);
        model.addAttribute("currentYear", year);
        model.addAttribute("currentMonth", month);
        model.addAttribute("menu", "");
        Authentication authentication = SecurityContextHolder.getContext().getAuthentication();

        if (authentication != null && authentication.getPrincipal() instanceof CustomUserPrincipal) {
            CustomUserPrincipal customUserPrincipal = (CustomUserPrincipal) authentication.getPrincipal();
            if (customUserPrincipal.getCanModify().equalsIgnoreCase("Y")){
                model.addAttribute("canModifyRecord", true);
            }else{
                model.addAttribute("canModifyRecord", false);
            }
            model.addAttribute("theuserid", principal.getName());

        }
    
        //  Fetch Status 
        String companyCode = (String) session.getAttribute("gblCompanyCode");

        //int countnew = 0; 
       int countnew = safService.safSubmitStatus(year, month,companyCode);

        model.addAttribute("countnew", countnew);
        System.out.println("Status countnew = " + countnew);

   
        return "saf/safmenu";
    }
    
    public static String getMonthName(int monthNumber) {
        if (monthNumber < 1 || monthNumber > 12) {
            throw new IllegalArgumentException("Invalid month number: " + monthNumber);
        }
        return Month.of(monthNumber).name();
    }

/*
    public static String getMonthName(int monthNumber) {
        if (monthNumber < 1 || monthNumber > 12) {
            throw new IllegalArgumentException("Invalid month number: " + monthNumber);
        }
        return Month.of(monthNumber).name().charAt(0) + Month.of(monthNumber).name().substring(1).toLowerCase();
    }
*/




    

    @PreAuthorize("isAuthenticated()")
    @PostMapping("/listsaf")
    public String listsaf(@RequestParam("month") String month,@RequestParam("year") String year,HttpSession session,Model model, Principal principal) {

        
        String companyCode = (String) session.getAttribute("gblCompanyCode");

        model.addAttribute("theuserid", principal.getName());
        model.addAttribute("menu", "SAF");
        Optional<Saftxn> saftxns = safService.findByCcompcodeAndCmonthAndCyear(companyCode, month, year);
        
        List<Safcard> safcards = safcardService.findAll();

        int monthNumber = Integer.parseInt(month); // Example month number
        String monthName = getMonthName(monthNumber);
        System.out.println("Month name: " + monthName);
        model.addAttribute("monthName", monthName);
        model.addAttribute("ccompcode", companyCode);
        model.addAttribute("cmonth", month);
        model.addAttribute("cyear", year);

        
        if (saftxns.isPresent()) {
            System.out.println("Net profit loss =  " + saftxns.get().getNet_profit_loss());
            System.out.println("Net loss profit=  " + saftxns.get().getNet_loss_profit());
            model.addAttribute("safcards", safcards);
            model.addAttribute("saftxns", saftxns);
            
            model.addAttribute("net_profit_loss",saftxns.get().getNet_profit_loss());
            model.addAttribute("net_loss_profit",saftxns.get().getNet_loss_profit());
            
            model.addAttribute("add_depreciation",saftxns.get().getAdd_depreciation());
            model.addAttribute("less_depreciation",saftxns.get().getLess_depreciation());
            model.addAttribute("add_amortisation",saftxns.get().getAdd_amortisation());
            model.addAttribute("less_amortisation",saftxns.get().getLess_amortisation());
            model.addAttribute("add_prov_for_impairment",saftxns.get().getAdd_prov_for_impairment());
            model.addAttribute("less_prov_for_impairment",saftxns.get().getLess_prov_for_impairment());
            model.addAttribute("less_profit_on_si",saftxns.get().getLess_profit_on_si());
            model.addAttribute("add_profit_on_si",saftxns.get().getAdd_profit_on_si());
            model.addAttribute("less_profit_on_sfa",saftxns.get().getLess_profit_on_sfa());
            model.addAttribute("add_profit_on_sfa",saftxns.get().getAdd_profit_on_sfa());
            model.addAttribute("inventory_decr",saftxns.get().getInventory_decr());
            model.addAttribute("inventory_incr",saftxns.get().getInventory_incr());
            model.addAttribute("receivable_decr",saftxns.get().getReceivable_decr());
            model.addAttribute("receivable_incr",saftxns.get().getReceivable_incr());
            model.addAttribute("other_ca_decr",saftxns.get().getOther_ca_decr());
            model.addAttribute("other_ca_incr",saftxns.get().getOther_ca_incr());
            model.addAttribute("other_cr_incr",saftxns.get().getOther_cr_incr());
            model.addAttribute("other_cr_decr",saftxns.get().getOther_cr_decr());
            model.addAttribute("other_cl_incr",saftxns.get().getOther_cl_incr());
            model.addAttribute("other_cl_decr",saftxns.get().getOther_cl_decr());
            model.addAttribute("recpt_of_term_loan",saftxns.get().getRecpt_of_term_loan());
            model.addAttribute("repymt_of_term_loan",saftxns.get().getRepymt_of_term_loan());
            model.addAttribute("incr_in_od_ltr_bd",saftxns.get().getIncr_in_od_ltr_bd());
            model.addAttribute("decr_in_od_ltr_bd",saftxns.get().getDecr_in_od_ltr_bd());

            model.addAttribute("decr_in_cash_bank",saftxns.get().getDecr_in_cash_bank());
            model.addAttribute("incr_in_cash_bank",saftxns.get().getIncr_in_cash_bank());
            model.addAttribute("loan_fm_ogcs",saftxns.get().getLoan_fm_ogcs());
            model.addAttribute("loan_to_ogcs",saftxns.get().getLoan_to_ogcs());

            model.addAttribute("loan_fm_othr",saftxns.get().getLoan_fm_othr());
            model.addAttribute("loan_to_othr",saftxns.get().getLoan_to_othr());
            model.addAttribute("sale_proceeds_of_fa",saftxns.get().getSale_proceeds_of_fa());
            model.addAttribute("addn_to_fa",saftxns.get().getAddn_to_fa());
            model.addAttribute("sale_proceeds_of_invst",saftxns.get().getSale_proceeds_of_invst());
            model.addAttribute("purchase_of_invst",saftxns.get().getPurchase_of_invst());
            model.addAttribute("capital_contribution",saftxns.get().getCapital_contribution());
            model.addAttribute("divnd_paid",saftxns.get().getDivnd_paid());
            model.addAttribute("proprietor_contribution",saftxns.get().getProprietor_contribution());
            model.addAttribute("proprietor_withdrawal",saftxns.get().getProprietor_withdrawal());
            model.addAttribute("remarks",saftxns.get().getRemarks());
            
            
        }
        else
        {
            
            Saftxn newSaftxn = new Saftxn();
            newSaftxn.setCcompcode(companyCode);
            newSaftxn.setCmonth(month);
            newSaftxn.setCyear(year);

            System.out.println("SAF New Record ccompcode =  " + companyCode);
            System.out.println("SAF New Record cmonth =  " + month);
            System.out.println("SAF New Record cyear =  " + year);

            
            

            model.addAttribute("safcards", safcards);
            model.addAttribute("saftxns", newSaftxn);
            model.addAttribute("net_profit_loss",BigDecimal.ZERO);
            model.addAttribute("net_loss_profit",BigDecimal.ZERO);
            
            model.addAttribute("add_depreciation",BigDecimal.ZERO);
            model.addAttribute("less_depreciation",BigDecimal.ZERO);
            model.addAttribute("add_amortisation",BigDecimal.ZERO);
            model.addAttribute("less_amortisation",BigDecimal.ZERO);
            model.addAttribute("add_prov_for_impairment",BigDecimal.ZERO);
            model.addAttribute("less_prov_for_impairment",BigDecimal.ZERO);
            model.addAttribute("less_profit_on_si",BigDecimal.ZERO);
            model.addAttribute("add_profit_on_si",BigDecimal.ZERO);
            model.addAttribute("less_profit_on_sfa",BigDecimal.ZERO);
            model.addAttribute("add_profit_on_sfa",BigDecimal.ZERO);
            model.addAttribute("inventory_decr",BigDecimal.ZERO);
            model.addAttribute("inventory_incr",BigDecimal.ZERO);
            model.addAttribute("receivable_decr",BigDecimal.ZERO);
            model.addAttribute("receivable_incr",BigDecimal.ZERO);
            model.addAttribute("other_ca_decr",BigDecimal.ZERO);
            model.addAttribute("other_ca_incr",BigDecimal.ZERO);
            model.addAttribute("other_cr_incr",BigDecimal.ZERO);
            model.addAttribute("other_cr_decr",BigDecimal.ZERO);
            model.addAttribute("other_cl_incr",BigDecimal.ZERO);
            model.addAttribute("other_cl_decr",BigDecimal.ZERO);
            model.addAttribute("recpt_of_term_loan",BigDecimal.ZERO);
            model.addAttribute("repymt_of_term_loan",BigDecimal.ZERO);
            model.addAttribute("incr_in_od_ltr_bd",BigDecimal.ZERO);
            model.addAttribute("decr_in_od_ltr_bd",BigDecimal.ZERO);

            model.addAttribute("decr_in_cash_bank",BigDecimal.ZERO);
            model.addAttribute("incr_in_cash_bank",BigDecimal.ZERO);
            model.addAttribute("loan_fm_ogcs",BigDecimal.ZERO);
            model.addAttribute("loan_to_ogcs",BigDecimal.ZERO);

            model.addAttribute("loan_fm_othr",BigDecimal.ZERO);
            model.addAttribute("loan_to_othr",BigDecimal.ZERO);
            model.addAttribute("sale_proceeds_of_fa",BigDecimal.ZERO);
            model.addAttribute("addn_to_fa",BigDecimal.ZERO);
            model.addAttribute("sale_proceeds_of_invst",BigDecimal.ZERO);
            model.addAttribute("purchase_of_invst",BigDecimal.ZERO);
            model.addAttribute("capital_contribution",BigDecimal.ZERO);
            model.addAttribute("divnd_paid",BigDecimal.ZERO);
            model.addAttribute("proprietor_contribution",BigDecimal.ZERO);
            model.addAttribute("proprietor_withdrawal",BigDecimal.ZERO);
            model.addAttribute("remarks","");
        }




        return "saf/saf";
        


    }


    @PostMapping("/update-saftxn")
    public String submitSaftxn(@ModelAttribute Saftxn saftxns,BindingResult result,Model model,Principal principal) {
        model.addAttribute("theuserid", principal.getName());
        model.addAttribute("menu", "");
        safService.save(saftxns);
        //return "mistxn/mistxn";
        return "redirect:/safmenu";
    }
}
