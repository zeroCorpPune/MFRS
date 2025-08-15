package com.rcs.mfrs.Controller;

import java.security.Principal;
import java.time.Month;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import javax.servlet.http.HttpSession;

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
import com.rcs.mfrs.Model.Loans_from_to;
import com.rcs.mfrs.Services.CompanyService;
import com.rcs.mfrs.Services.Loans_from_toService;
import com.rcs.mfrs.Util.CompanyDTO;
import com.rcs.mfrs.Util.MCommon;


@Controller
public class Loans_from_toController {


    @Autowired
    private Loans_from_toService loans_from_toService;

    @Autowired
    private CompanyService companyService;

    public static String getMonthName(int monthNumber) {
        if (monthNumber < 1 || monthNumber > 12) {
            throw new IllegalArgumentException("Invalid month number: " + monthNumber);
        }
        return Month.of(monthNumber).name();
    }

    @PostMapping("/loansFromAndTo/saveAll")
    public String saveLoans_from_to(@ModelAttribute Loans_from_toWrapper loans_from_toWrapper,
                          @RequestParam(value = "deletedRows", required = false) String deletedRowsJson,
                          BindingResult result, Model model) {
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
                    String ccompname = row.get("ccompname");

                    System.out.println("ccompcode = " + ccompcode);
                    System.out.println("cyear = " + cyear);
                    System.out.println("cmonth = " + cmonth);
                    System.out.println("ccompname = " + ccompname);
                    // Delete the record from the database based on its unique identifiers
                    loans_from_toService.deleteByPrimaryKey(ccompcode, cyear, cmonth,ccompname);
                    
                }
            } catch (JsonProcessingException e) {
                e.printStackTrace(); // Handle exception as needed
            }

            List<Loans_from_to> records = loans_from_toWrapper.getLoans_from_toList();
            loans_from_toService.saveAll(records);
        }

        return "redirect:/safmenu";
    }                        
    
    @PreAuthorize("isAuthenticated()")
    @PostMapping("/loansft")
    public String loansft(@RequestParam("month") String month,@RequestParam("year") String year,HttpSession session,Model model, Principal principal) {
        model.addAttribute("theuserid", principal.getName());
        model.addAttribute("menu", "SAF");

        String companyCode = (String) session.getAttribute("gblCompanyCode");

        int monthNumber = Integer.parseInt(month); // Example month number
        String monthName = getMonthName(monthNumber);
        System.out.println("Month name: " + monthName);
        model.addAttribute("monthName", monthName);
        
        model.addAttribute("theCompanyCode", companyCode);
        model.addAttribute("theYear", year);
        model.addAttribute("theMonth", month);

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

        List<Loans_from_to> loans_from_toList = loans_from_toService.findByCcompcodeAndCyearAndCMonth(companyCode,year, month);
        model.addAttribute("loans_from_toList", loans_from_toList);

        return "saf/loans_from_to";
    }
    
}
