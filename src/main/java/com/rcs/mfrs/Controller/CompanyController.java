package com.rcs.mfrs.Controller;

import java.io.FileNotFoundException;
import java.security.Principal;
import java.util.List;
import java.util.Optional;

import javax.validation.Valid;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.security.access.prepost.PreAuthorize;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.servlet.ModelAndView;

import com.rcs.mfrs.Model.Company;
import com.rcs.mfrs.Model.Usermaster;
import com.rcs.mfrs.Security.CustomUserPrincipal;
import com.rcs.mfrs.Services.CompanyService;
import com.rcs.mfrs.Services.UsermasterService;

import net.sf.jasperreports.engine.JRException;



@Controller
public class CompanyController {
    @Autowired
    private CompanyService companyService;

    @Autowired
    private UsermasterService usermasterService;


    @PostMapping("/report/{format}")
    public ResponseEntity<String> generateReport(@PathVariable String format) throws  FileNotFoundException, JRException{   
       // return companyService.exportReport(format);
       String result = companyService.exportReport(format);

       return ResponseEntity.ok(result);
    }

    /*
    @GetMapping("/editcompany/{id}")
    public String getPost(@PathVariable String id, Model model) {
        System.out.println("Company edit - 10");
        Optional<Company> optionalCompany = companyService.getbyCcompcode(id);
        System.out.println("Company edit - 10.1");
        if (optionalCompany.isPresent()) {
            System.out.println("Company edit - 10.2");
            Company company = optionalCompany.get();
            System.out.println("Company edit - 10.3");
            model.addAttribute("company", company);
            System.out.println("Company edit - 10.4");
            return "company/editcompany";
        } else {
            return "404";
        }
    }
      */ 
    @PostMapping("/editcompany")
    @PreAuthorize("isAuthenticated()")
    public String updatePost(@Valid  @ModelAttribute Company post, BindingResult bindingResult){
    //public String updatePost(@Valid  @ModelAttribute Company post, BindingResult bindingResult){        
    /*
        System.out.println("Company edit - 11");
        
       // System.out.println("the id is " + id);
        if (bindingResult.hasErrors()){
            System.out.println("error is there");
            return "company/editcompany";
        }
        System.out.println("Company edit - 11.2");
        Optional<Company> optionalPost = companyService.getbyCcompcode(post.getCcompcode());
        System.out.println("Company edit - 11.3"); 
        if(optionalPost.isPresent()){
            System.out.println("Company edit - 11.4");
            Company existingPost = optionalPost.get();
            System.out.println("Company edit - 11.5");
            //existingPost.setCcompname(post.getCcompname());
            existingPost.setCEMAIL(post.getCEMAIL());
           // existingPost.setUseflash(post.getUseflash());
           // System.out.println(existingPost.getCEMAIL());
            companyService.save(existingPost);
            System.out.println("Company edit - 11.7");
        }
        
        //return "redirect:/posts/"+post.getId();
        
        return "redirect:/companylist";
        */
        if (bindingResult.hasErrors()) {
            System.out.println("Validation errors occurred");
            return "company/editcompany";
        }
    
        Optional<Company> optionalPost = companyService.getbyCcompcode(post.getCcompcode());
        if (optionalPost.isPresent()) {
            Company existingPost = optionalPost.get();
            System.out.println("New CEMAIL from form: " + post.getCEMAIL());
            
            // Overwrite the existing CEMAIL with the new value from the form
            //optionalPost.get().setCEMAIL("");
            optionalPost.get().setCEMAIL(post.getCEMAIL());
            optionalPost.get().setBUDGET_CEMAIL(post.getBUDGET_CEMAIL());
            

            System.out.println("Old CEMAIL: " + existingPost.getCEMAIL());
    
            // Save the updated entity
            companyService.save(existingPost);
    
        }
    
        return "redirect:/companylist";

    }
            
    

    @PostMapping("/company")
    @PreAuthorize("isAuthenticated()")
    public String updatePost1(@ModelAttribute Company company, Model model ,  Principal principal){
        System.out.println("Company edit - 12");
       
        System.out.println("Company edit - 12.0");

        System.out.println("the code is " + company.getCcompcode());

        Optional<Usermaster> optionalusermaster = usermasterService.getbyCuserid(principal.getName());
        Optional<Company> optionalAccount = companyService.getbyCcompcode(company.getCcompcode());
        System.out.println("Company edit - 12.1");
        if(optionalusermaster.isPresent()){
            System.out.println("Company edit - 12.2");
            //Company company = new Company();
            //System.out.println(company.getCcompname());
           System.out.println("Company edit - 12.3");
             
            System.out.println("Company edit - 12.4");
            model.addAttribute("company", optionalAccount);
            model.addAttribute("CEMAIL", company.getCEMAIL());

            System.out.println(optionalAccount.get().getCEMAIL());
            model.addAttribute("budget_cemail", optionalAccount.get().getBUDGET_CEMAIL());
            System.out.println(optionalAccount.get().getBUDGET_CEMAIL());

            model.addAttribute("CATEGORY", optionalAccount.get().getCATEGORY());

            
            
            model.addAttribute("CEMAIL", optionalAccount.get().getCEMAIL());
            model.addAttribute("CDEFAULT", optionalAccount.get().getCDEFAULT());
            model.addAttribute("CCONTACTPERSON", optionalAccount.get().getCCONTACTPERSON());
            model.addAttribute("CADDRESS1", optionalAccount.get().getCADDRESS1());
            model.addAttribute("CADDRESS2", optionalAccount.get().getCADDRESS2());
            model.addAttribute("CADDRESS3", optionalAccount.get().getCADDRESS3());
            model.addAttribute("CADDRESS4", optionalAccount.get().getCADDRESS4());
            model.addAttribute("CADDRESS5", optionalAccount.get().getCADDRESS5());
            
            model.addAttribute("CPHONE", optionalAccount.get().getCPHONE());
            model.addAttribute("CFAX", optionalAccount.get().getCFAX());
            model.addAttribute("CUSER_ID", optionalAccount.get().getCUSER_ID());
            model.addAttribute("CUSER_NAME", optionalAccount.get().getCUSER_NAME());
            model.addAttribute("CUSER_TYPE", optionalAccount.get().getCUSER_TYPE());
            model.addAttribute("PASSWORD", optionalAccount.get().getPASSWORD());
            model.addAttribute("CACTIVE", optionalAccount.get().getCACTIVE());
            
            model.addAttribute("DCREATIONDATE", optionalAccount.get().getDCREATIONDATE());
            model.addAttribute("DMODIFIEDDATE", optionalAccount.get().getDMODIFIEDDATE());
            
            model.addAttribute("CCOMPNAMESN", optionalAccount.get().getCCOMPNAMESN());
            model.addAttribute("CCOMPNAMESN1", optionalAccount.get().getCCOMPNAMESN1());
            model.addAttribute("SAFYN", optionalAccount.get().getSAFYN());
            model.addAttribute("RECPAYYN", optionalAccount.get().getRECPAYYN());
            model.addAttribute("FINPOSYN", optionalAccount.get().getFINPOSYN());
            model.addAttribute("STAKE", optionalAccount.get().getSTAKE());
            
            model.addAttribute("TRADERECYN", optionalAccount.get().getTRADERECYN());
            model.addAttribute("BANKBORROWINGSYN", optionalAccount.get().getBANKBORROWINGSYN());
            model.addAttribute("CCOMPCODE_OLD", optionalAccount.get().getCCOMPCODE_OLD());
            model.addAttribute("CCOMPCODE_NEW", optionalAccount.get().getCCOMPCODE_NEW());
            model.addAttribute("WCYN", optionalAccount.get().getWCYN());
            model.addAttribute("SECTORNO", optionalAccount.get().getSECTORNO());
            model.addAttribute("MISYN", optionalAccount.get().getMISYN());
            model.addAttribute("MFRSYN", optionalAccount.get().getMFRSYN());
            model.addAttribute("CCOMPNAMESN2", optionalAccount.get().getCCOMPNAMESN2());
            model.addAttribute("UPLOADFLASHFROMEXCEL", optionalAccount.get().getUPLOADFLASHFROMEXCEL());
            model.addAttribute("INCLUDEINFLASHREPORT", optionalAccount.get().getINCLUDEINFLASHREPORT());
            model.addAttribute("INCLUDEINBBREPORT", optionalAccount.get().getINCLUDEINBBREPORT());
            model.addAttribute("SECTIONINBBREPORT", optionalAccount.get().getSECTIONINBBREPORT());
            model.addAttribute("FLEXI_CCOMPCODE", optionalAccount.get().getFLEXI_CCOMPCODE());
            model.addAttribute("COMREGNO", optionalAccount.get().getCOMREGNO());
            /* */
            System.out.println("Company edit - 12.5");
            model.addAttribute("theuserid", principal.getName());
            model.addAttribute("menu", "Company");
            //return "company/editcompany";
           
            return "company/company";

        }else{
            return "redirect:/";
        }
        //return "redirect:/posts/"+post.getId();
        //return "redirect:/company/editwithformfield";

    }

    @PostMapping("/company/add")
    @PreAuthorize("isAuthenticated()")
    public String addCompany(Model model, Principal principal) {
        String authUser = "email";
        System.out.println("Company Add - 1");
        if(principal != null){
            authUser = principal.getName();
            System.out.println("Company Add - 2");
            
        }
        System.out.println(authUser);
        
        Optional<Usermaster> optionalusermaster = usermasterService.getbyCuserid(principal.getName());
        
        
        System.out.println("The company" + optionalusermaster.get().getCompany().getCcompcode());
        Optional<Company> optionalAccount = companyService.getbyCcompcode(optionalusermaster.get().getCompany().getCcompcode());
        System.out.println("Company Add - 3");
        if(optionalusermaster.isPresent()){
            Company company = new Company();
            //company.setCcompcode("");

            System.out.println("Company Add - 4");
            //Company company = new Company();
            //System.out.println(company.getCcompname());
            System.out.println("Company Add - 5");
             
            System.out.println("Company Add - 6");
            model.addAttribute("company", company);
            System.out.println("Company Add - 7");
            return "company/addcompany";

        }else{
            return "redirect:/";
        }
    }

     


    @PostMapping("/company/addnewcompany")
    @PreAuthorize("isAuthenticated()")
    public String addNewCompanyHandler(@Validated @ModelAttribute Company company, BindingResult bindingResult ,Principal principal, Model model){
        System.out.println("Now saving new company");
        if (bindingResult.hasErrors()){
            return "company/addcompany";
        }
        //if(company.findByCodeNumber(users.getCodeNumber())) {

        //Optional<Company> optionalAccount = companyService.getbyCcompcode(company.getCcompcode()));

        //System.out.println("User entered the code " + company.getCcompcode());
        Optional<Company> optionalAccount =  companyService.getbyCcompcode(company.getCcompcode());
        if (optionalAccount.isPresent())
        {
            model.addAttribute("customerror" , "Company code Record already exist");
            
            
            return "company/addcompany";
        }
        if(principal != null){
            companyService.save(company);
            return "redirect:/index";
        }
        else
            return "redirect:/?error";
        /*
        if (company.getAccount().getEmail().compareToIgnoreCase(authUser) < 0){
            return "redirect:/?error";
        }
        */
        
       
    }

/* 

    @GetMapping("/posts/{id}/delete")
    @PreAuthorize("isAuthenticated()")
    public String deletePost(@PathVariable Long id){
        Optional<Post> optionalPost = postService.getById(id);
        if(optionalPost.isPresent()){
            Post post = optionalPost.get();
            postService.delete(post);
            return "redirect:/";
        }else{
            return "redirect:/?error";
        }

    }

    */
}