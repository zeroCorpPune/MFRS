package com.rcs.mfrs.Controller;

import java.math.BigDecimal;
import java.security.Principal;
import java.time.LocalDate;
import java.util.ArrayList;
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

import com.rcs.mfrs.Model.Mistxn;
import com.rcs.mfrs.Security.CustomUserPrincipal;
import com.rcs.mfrs.Services.MistxnService;
import com.rcs.mfrs.Services.MistxnServiceImpl;
import com.rcs.mfrs.Util.MCommon;

import com.rcs.mfrs.Model.Flash_rem;
import com.rcs.mfrs.Services.CompanyService;
import com.rcs.mfrs.Services.Flash_remService;

@Controller
public class MistxnController {


    @Autowired
    private MistxnService mistxnService;
    
    @Autowired
    private CompanyService companyService;

    @Autowired
    private MistxnServiceImpl mistxnServiceImpl;

    @Autowired
    private Flash_remService flash_remService;

   // @Autowired
   // ObjectFactory<HttpSession> httpSessionFactory;

    private final MCommon mcomm;

    @Autowired
    public MistxnController(MCommon mcomm) {
        this.mcomm = mcomm;
    }


    @PostMapping("/cancel")
    public String cancel(Model model, Principal principal,HttpSession session) {
        model.addAttribute("menu", "");
        return "redirect:/mistxnmenu";
    }

    @GetMapping("/mistxnmenu")
    public String mistxnmenu( 
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

        
        int countnew = mistxnService.flashSubmitStatus(year, month,companyCode);
        model.addAttribute("countnew", countnew);
        System.out.println("Status countnew = " + countnew);

   
        return "mistxn/mistxnmenu";
    }

    @PostMapping("/misentry")
    public String handleForm1(@RequestParam("month") String value, @RequestParam("year") String yearvalue , HttpSession session) {
        session.setAttribute("month", value);
        session.setAttribute("year", yearvalue);
        //System.out.println(session.getAttribute("month"));
        //System.out.println(session.getAttribute("year"));
        return "redirect:/listflash";
    }

    //@GetMapping("/listflash")
    @PreAuthorize("isAuthenticated()")
    @PostMapping("/listflash")
    public String listmistxn(@RequestParam("month") String value,@RequestParam("year") String yearvalue,HttpSession session,Model model, Principal principal) {
        System.out.println("flash get");
        session.setAttribute("month", value);
        session.setAttribute("year", yearvalue);
        
        System.out.println(session.getAttribute("month"));        
        System.out.println((String) session.getAttribute("month"));
        System.out.println((String) session.getAttribute("year"));
        // Usermaster account = new Usermaster();
        // model.addAttribute("useraccount", account);
        
        // This is working for update
        model.addAttribute("theuserid", principal.getName());
        model.addAttribute("menu", "Flash");

        
        
        
        
        
       // HttpSession session1 = httpSessionFactory.getObject();

        System.out.println("gblCompanyCode = " +  session.getAttribute("gblCompanyCode") );     

       // List<Mistxn> mistxns = mistxnService.findByCriteria("RCS", "06", "2024");

        String month = (String) session.getAttribute("month");
        String year = (String) session.getAttribute("year");
        String companyCode = (String) session.getAttribute("gblCompanyCode");

        int numeric_year = Integer.parseInt(year);

        String year_1 = String.valueOf(numeric_year-1);

        long countnew = mistxnService.countTransactionsByMonthAndYear(year, month,companyCode);
        model.addAttribute("countnew", countnew);
        System.out.println("countnew = " + countnew);

        if (countnew>0){
            List<Mistxn> mistxns = mistxnService.findByCriteria(companyCode, month, year);
            //System.out.println("total records fetched " + mistxns.size());
            MistxnListWrapper mistxnListWrapper = new MistxnListWrapper();
            mistxnListWrapper.setMistxns(mistxns); // add your Mistxn objects here if necessary
            model.addAttribute("mistxnListWrapper", mistxnListWrapper);

            Optional<Flash_rem> flash_rems = flash_remService.findByCriteria(companyCode, month, year);

            if (flash_rems.isPresent()) {
                // Display fields from the single Flash_rem object in the console
                Flash_rem rem = flash_rems.get();
                System.out.println(rem.getCcompcode());
                System.out.println(rem.getCmonth());
                System.out.println(rem.getCyear());
                System.out.println(rem.getRemarks());
                // Add other fields you want to display in the console
            } else {
                System.out.println("No records found.*****************************");
            }

            model.addAttribute("flash_rems",flash_rems);
            System.out.println("is the flash remark emply ? - "  + flash_rems.isEmpty());
        }
        else{
            // Fetch data for specific cmonth and cyear for the desired fields
            List<Mistxn> Prev_year_1_fields = mistxnService.findByCriteria(companyCode, month, year_1);
            

            // Combine the data
            List<Mistxn> combinedMistxns = new ArrayList<>();
            for (int i = 0; i < Prev_year_1_fields.size(); i++) {
                Mistxn Prev_year1 = Prev_year_1_fields.get(i);
                

                Mistxn combinedMistxn = new Mistxn();
                combinedMistxn.setCcompcode(companyCode);
                combinedMistxn.setCmonth(month);
                combinedMistxn.setCyear(year);
                combinedMistxn.setDesccd(Prev_year1.getDesccd());
                combinedMistxn.setBoldyn(Prev_year1.getBoldyn());
                combinedMistxn.setDescription(Prev_year1.getDescription());

                combinedMistxn.setActMtdBal(BigDecimal.ZERO);
                combinedMistxn.setBgtMtdBal(BigDecimal.ZERO);
                combinedMistxn.setActMtdBalPy1(Prev_year1.getActMtdBal());
                combinedMistxn.setActYtdBal(BigDecimal.ZERO);
                combinedMistxn.setBgtYtdBal(BigDecimal.ZERO);
                combinedMistxn.setActYtdBalPy1(Prev_year1.getActYtdBal());
                combinedMistxn.setActYtdBalPy2(Prev_year1.getActYtdBalPy1());
                combinedMistxn.setActYtdBalPy3(Prev_year1.getActYtdBalPy2());
                combinedMistxn.setActYtdBalPy4(Prev_year1.getActYtdBalPy3());
                combinedMistxn.setCblock("N");
                combinedMistxn.setCactive(Prev_year1.getCactive());

                combinedMistxn.setDcreationdate(Prev_year1.getDcreationdate());
                combinedMistxn.setDmodifieddate(Prev_year1.getDmodifieddate());

                combinedMistxns.add(combinedMistxn);
            }

            MistxnListWrapper mistxnListWrapper = new MistxnListWrapper();
            mistxnListWrapper.setMistxns(combinedMistxns);
            model.addAttribute("mistxnListWrapper", mistxnListWrapper);
        }

        model.addAttribute("cur_year",session.getAttribute("year"));
        Object yearObj = session.getAttribute("year");
        if (yearObj != null) {
            int yeartodisplay = Integer.parseInt(yearObj.toString());
            model.addAttribute("prev_year", yeartodisplay - 1);
            model.addAttribute("prev_year2", yeartodisplay - 2);
            model.addAttribute("prev_year3", yeartodisplay - 3);
            model.addAttribute("prev_year4", yeartodisplay - 4);
        } else {
            // Handle the case where the year is not present in the session
            model.addAttribute("prev_year", "No year");
        }


        return "mistxn/mistxn";
    }
//  public String updateEntity(@ModelAttribute Mistxn entity) {    
//  public String updateEntity(@ModelAttribute("mistxns") List<Mistxn> mistxns, BindingResult result) {        
    @PostMapping("/update-mistxn")
    public String submitMistxn(@ModelAttribute Flash_rem flash_rems,@ModelAttribute MistxnListWrapper mistxnListWrapper,BindingResult result,Model model) {
        List<Mistxn> mistxns = mistxnListWrapper.getMistxns();
        
        if (result.hasErrors()) {
            System.out.println("has error...............................");
            //model.addAttribute("error", "Validation errors occurred");
            //return "redirect:/mistxnmenu";
        }
        if (mistxns == null || mistxns.contains(null)) {
            // Add error handling here
            model.addAttribute("error", "List of transactions is null or contains null elements");
             return "mistxn/mistxn";
        }
        

        System.out.println("flash save");
        


        System.out.println(flash_rems.getCcompcode());
        System.out.println(flash_rems.getCmonth());
        System.out.println(flash_rems.getCyear());
        System.out.println(flash_rems.getRemarks());




        //entity.getDcreationdate(LocalDateTime.now());
        mistxnService.saveAll(mistxns);
        flash_remService.save(flash_rems);
        //return "mistxn/mistxn";
        return "redirect:/mistxnmenu";
    }

    private int parseInt(Object attribute) {
        throw new UnsupportedOperationException("Not supported yet.");
    }



    @PreAuthorize("isAuthenticated()")
    @PostMapping("/submitFlash")
    public String submitFlash(@RequestParam("month") String value,@RequestParam("year") String yearvalue,HttpSession session,Model model, Principal principal) {
        System.out.println("flash Submit");
        session.setAttribute("month", value);
        session.setAttribute("year", yearvalue);
        String theCcompCode =  (String) session.getAttribute("gblCompanyCode");
        

        int updatedRows = mistxnService.updateMistxnStatus(yearvalue, value,theCcompCode);
        if (updatedRows > 0) {
            System.out.println("Update successful. Rows affected: " + updatedRows);
        } else {
            System.out.println("No records were updated.");
        }
        
        return "redirect:/mistxnmenu";
    }

}
