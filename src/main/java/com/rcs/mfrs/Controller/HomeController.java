package com.rcs.mfrs.Controller;

import java.security.Principal;
import java.util.List;
import java.util.stream.Collectors;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.rcs.mfrs.Model.Company;
import com.rcs.mfrs.Model.Usermaster;
import com.rcs.mfrs.Security.CustomUserPrincipal;
import com.rcs.mfrs.Services.CompanyService;
import com.rcs.mfrs.Services.UsermasterService;
import com.rcs.mfrs.Util.CompanyDTO;

@Controller
public class HomeController {
    @Autowired
    private CompanyService companyService;

    @Autowired
    private UsermasterService usermasterService;

    // @Autowired
    // private CustomUserPrincipal customUserPrincipal;

    //@GetMapping("/index")
    @GetMapping("/companylist")
    public String index(Model model, Principal principal) {
       
        List<Company> companies = companyService.getAll();
        model.addAttribute("companylisting", companies);
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
            e.printStackTrace();  // Handle this exce as per your requirement
        }

        model.addAttribute("companiesJson", companiesJson);

    
        

        Authentication authentication = SecurityContextHolder.getContext().getAuthentication();

        if (authentication != null && authentication.getPrincipal() instanceof CustomUserPrincipal) {
            CustomUserPrincipal customUserPrincipal = (CustomUserPrincipal) authentication.getPrincipal();
            if (customUserPrincipal.getCanModify().equalsIgnoreCase("Y")){
                model.addAttribute("canModifyRecord", true);
            }else{
                model.addAttribute("canModifyRecord", false);
            }
            model.addAttribute("theuserid", principal.getName());
            model.addAttribute("canAdd", customUserPrincipal.getCanAdd());     
            model.addAttribute("canModify", customUserPrincipal.getCanModify());
            model.addAttribute("canDelete", customUserPrincipal.getCanDelete());
            //model.addAttribute("ccompcode", customUserPrincipal.getCcompcode());
            List<Usermaster> userlistAll = usermasterService.getAll();
            List<Usermaster> selecteduserlist = usermasterService.getEntitiesBypara(principal.getName());
    
            model.addAttribute("userlistingAll", userlistAll);
            model.addAttribute("selecteduserlisting", selecteduserlist);
        }

        model.addAttribute("menu", "");
        model.addAttribute("theuserid", principal.getName());
        
        //return "index";
        return "company/companylist";
    }


    @GetMapping("/body")
    public String body(Model model, Principal principal) {
        System.out.println("calling body.html ");
        return "body";
    }

    @GetMapping("/homeexample")
    public String homeexample(Model model, Principal principal) {
        System.out.println("calling homeexample.html ");
        return "homeexample";
    }
    
    @GetMapping("/about")
    public String about(Model model, Principal principal) {
        System.out.println("calling body.html ");
        return "about";
    }

    @GetMapping("/products")
    public String products(Model model, Principal principal) {
        System.out.println("calling products.html ");
        return "products";
    }

    @GetMapping("/bodynew")
    public String bodynew(Model model, Principal principal) {
        System.out.println("calling bodynew.html ");
        return "bodynew";
    }

    @GetMapping("/home_old")
    public String home_old(Model model, Principal principal) {
        System.out.println("calling home_old.html ");
        return "home_old";
    }

    @GetMapping("/home")
    public String home(Model model, Principal principal) {
        System.out.println("calling home i.e. home.html............");
        List<Company> companies = companyService.getAll();
        model.addAttribute("companylisting", companies);
        

        Authentication authentication = SecurityContextHolder.getContext().getAuthentication();

        if (authentication != null && authentication.getPrincipal() instanceof CustomUserPrincipal) {
            CustomUserPrincipal customUserPrincipal = (CustomUserPrincipal) authentication.getPrincipal();
            if (customUserPrincipal.getCanModify().equalsIgnoreCase("Y")){
                model.addAttribute("canModifyRecord", true);
            }else{
                model.addAttribute("canModifyRecord", false);
            }
            model.addAttribute("theuserid", principal.getName());
            model.addAttribute("canAdd", customUserPrincipal.getCanAdd());     
            model.addAttribute("canModify", customUserPrincipal.getCanModify());
            model.addAttribute("canDelete", customUserPrincipal.getCanDelete());
            List<Usermaster> userlistAll = usermasterService.getAll();
            List<Usermaster> selecteduserlist = usermasterService.getEntitiesBypara(principal.getName());
    
            model.addAttribute("userlistingAll", userlistAll);
            model.addAttribute("selecteduserlisting", selecteduserlist);
            model.addAttribute("menu", "");
        }

       

        
        return "home";
    }    
    
    
    @GetMapping("/login")
    public String login(Model model) {
        System.out.println("Account Controller - get");
        // Usermaster account = new Usermaster();
        // model.addAttribute("useraccount", account);
        return "login";
    }

    




    
        
       
    
}