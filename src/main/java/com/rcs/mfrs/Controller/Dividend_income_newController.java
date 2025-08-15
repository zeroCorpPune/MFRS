package com.rcs.mfrs.Controller;

import java.security.Principal;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.access.prepost.PreAuthorize;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.rcs.mfrs.Model.Company;
import com.rcs.mfrs.Model.Dividend_income_new;
import com.rcs.mfrs.Services.CompanyService;
import com.rcs.mfrs.Services.Dividend_income_newService;
import com.rcs.mfrs.Util.CompanyDTO;
import com.rcs.mfrs.Util.DividendIncomeDTO;

import com.fasterxml.jackson.core.type.TypeReference;

import javax.servlet.http.HttpSession;

import org.springframework.beans.factory.ObjectFactory;

@Controller
public class Dividend_income_newController {
    
    @Autowired
    ObjectFactory<HttpSession> httpSessionFactory;
    
    @Autowired
    private CompanyService companyService;

    @Autowired
    private Dividend_income_newService dividend_income_newService;
    
    @PreAuthorize("isAuthenticated()")
    @PostMapping("/listInterCo")
    public String listInterCo(@RequestParam("month") String value,@RequestParam("year") String yearvalue,Model model, Principal principal) {
        System.out.println("Dividend_income_newController");

        List<Company> companies = companyService.findAll();
       // String code=companies.get(1).getCcompname();
        List<CompanyDTO> companyDTOs = companies.stream()
        .map(c -> new CompanyDTO(c.getCcompcode(), c.getCcompname()))
        .collect(Collectors.toList());
    
        model.addAttribute("companies", companyDTOs);
        // Convert DTOs to JSON string manually
        ObjectMapper objectMapper = new ObjectMapper();
        String companiesJson = "";
        try {
            companiesJson = objectMapper.writeValueAsString(companyDTOs);
        } catch (JsonProcessingException e) {
            e.printStackTrace();  // Handle this exception as per your requirement
        }

        model.addAttribute("companiesJson", companiesJson);

        HttpSession session = httpSessionFactory.getObject();
        session.setAttribute("month", value);
        session.setAttribute("year", yearvalue);
        String month = (String) session.getAttribute("month");
        String year = (String) session.getAttribute("year");
        String companyCode = (String) session.getAttribute("gblCompanyCode");
        System.out.println(companyCode);     
        System.out.println(year);     
        System.out.println(month);     

// Fetch all existing Dividend_income_new records
        //List<Dividend_income_new> dividendIncomeList = dividend_income_newService.getAllDividendIncomeRecords();
        List<Dividend_income_new> dividendIncomeList = dividend_income_newService.findByCcompcodeAndCyearAndCMonth(companyCode,year, month);

        
        
    
        model.addAttribute("dividendIncomeList", dividendIncomeList);
        model.addAttribute("theCompanyCode", companyCode);
        model.addAttribute("theYear", year);
        model.addAttribute("theMonth", month);
        
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

        model.addAttribute("theuserid", principal.getName());
        model.addAttribute("menu", "Flash");

        return "mistxn/interco";
    }

/*
    @PostMapping("/dividend_income/saveAll")
    public String saveInterCo(@ModelAttribute Dividend_income_newWrapper dividendIncomeWrapper,BindingResult result,Model model) {
        // Save the submitted records
        //System.out.println(dividendIncomeWrapper.getDividendIncomeList().get(0).getCcompcode());
        
        if (result.hasErrors()) {
            System.out.println("had error");
        }

        List<Dividend_income_new> records = dividendIncomeWrapper.getDividendIncomeList();
        dividend_income_newService.saveAll(records);

        return "redirect:/mistxnmenu";
    }
*/

    @PostMapping("/dividend_income/saveAll")
    public String saveInterCo(@ModelAttribute Dividend_income_newWrapper dividendIncomeWrapper,
                          @RequestParam(value = "deletedRows", required = false) String deletedRowsJson,
                          BindingResult result, Model model) {
        // Save the submitted records
    
    
    // Handle the deleted rows
    if (deletedRowsJson != null && !deletedRowsJson.isEmpty()) {
        ObjectMapper objectMapper = new ObjectMapper();
        try {
            List<Map<String, String>> deletedRows = objectMapper.readValue(deletedRowsJson, List.class);
            
            // Iterate through deleted rows and delete them from the database
            for (Map<String, String> row : deletedRows) {
                String ccompcode = row.get("ccompcode");
                String cyear = row.get("cyear");
                String cmonth = row.get("cmonth");
                String description = row.get("description");

                System.out.println("ccompcode = " + ccompcode);
                System.out.println("cyear = " + cyear);
                System.out.println("cmonth = " + cmonth);
                System.out.println("description = " + description);
                // Delete the record from the database based on its unique identifiers
                dividend_income_newService.deleteByPrimaryKey(ccompcode, cyear, cmonth,description);
                
            }
        } catch (JsonProcessingException e) {
            e.printStackTrace(); // Handle exception as needed
        }

        List<Dividend_income_new> records = dividendIncomeWrapper.getDividendIncomeList();
        dividend_income_newService.saveAll(records);
    }

    return "redirect:/mistxnmenu";
        
}

@PostMapping("/cancelDividendEntry")
    public String cancelDividendEntry(Model model, Principal principal,HttpSession session) {
        return "redirect:/mistxnmenu";
    }
}
