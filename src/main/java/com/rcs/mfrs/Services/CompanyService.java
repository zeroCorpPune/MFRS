package com.rcs.mfrs.Services;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;
import org.springframework.util.ResourceUtils;

import com.rcs.mfrs.Model.Company;
import com.rcs.mfrs.Repositories.CompanyRepository;

import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;

@Service
public class CompanyService {
    @Autowired
    private CompanyRepository companyRepository;

    public List<Company> getAll(){
        return companyRepository.findAll();
    }

    public List<Company> findAll() {
        //return companyRepository.findAll(Sort.by(Sort.Direction.ASC, "ccompcode"));
        return companyRepository.findAll(Sort.by(Sort.Direction.ASC, "CATEGORY").and(Sort.by(Sort.Direction.ASC, "ccompcode")));

    }
    
    public String exportReport (String reportFormat) throws FileNotFoundException, JRException
    {
        String path = "D:\\MFRS\\";
        List<Company> companies = companyRepository.findAll();
        System.out.println("Total records fetched: " + companies.size());
    
        //System.out.println(companies.get(1).getCcompname());
        // Load the JRXML file from the classpath
        File file = ResourceUtils.getFile("classpath:company.jrxml");
        
        // Compile the Jasper report
        JasperReport jasperReport = JasperCompileManager.compileReport(file.getAbsolutePath());
        
        // Create a JRBeanCollectionDataSource from the list of Company entities
        JRBeanCollectionDataSource dataSource = new JRBeanCollectionDataSource(companies);
        
        // Parameters for the report
        Map<String, Object> parameters = new HashMap<>();
        parameters.put("CreatedBy", "Vipul");
        
        // Fill the report with data
        JasperPrint jasperPrint = JasperFillManager.fillReport(jasperReport, parameters, dataSource);
    
        // Export the report to HTML or PDF
        if (reportFormat.equalsIgnoreCase("html")) {
            JasperExportManager.exportReportToHtmlFile(jasperPrint, path + "companies.html");
        }
        if (reportFormat.equalsIgnoreCase("pdf")) {
            JasperExportManager.exportReportToPdfFile(jasperPrint, path + "companies.pdf");
        }
    
        return "Report Generated in "  + path;
       
    }

    public Optional<Company> getbyCcompcode(String ccompcode) {
        return companyRepository.findByCcompcode(ccompcode) ;
    }

    public Optional<Company> getbyUsermasters(String cuserid) {
        return companyRepository.findByUsermasters(cuserid) ;
    }

    public Company save(Company company){
        /* 
        if (company.getId() == null){
            company.setCreatedAt(LocalDateTime.now());
        }
        company.setUpdatedAt(LocalDateTime.now());
        */
        
        return companyRepository.save(company);
    }
}
