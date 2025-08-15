package com.rcs.mfrs.Util;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.rcs.mfrs.Controller.ValueWithDecimalAndString;
import com.rcs.mfrs.Repositories.MistxnRepository;


public class GroupResult {
  

    private MistxnRepository mistxnRepository;

    // ✅ Constructor to accept MistxnRepository
    public GroupResult(MistxnRepository mistxnRepository) {
        this.mistxnRepository = mistxnRepository;
    }

    public Map<String, Object> getPage1A(String month, String year) 
    {
        // Fetch the results
        System.out.println("year = " + year);
        System.out.println("month = " + month);
        List<Object[]> result = mistxnRepository.FlashMISPage1A(year, month);
    
        System.out.println("dddddddddddddddddddddddddddddddddddddddddddddddddddddddddddd.....................");
        TotalValueHelper totalMTD   = new TotalValueHelper();  // To accumulate total for A5, D5, E6
        TotalValueHelper totalMBDG  = new TotalValueHelper();  // To accumulate total for A5, D5, E6
        TotalValueHelper totalVAR1  = new TotalValueHelper();  // To accumulate total for A5, D5, E6
        
        TotalValueHelper totalLY    = new TotalValueHelper();  // To accumulate total for A5, D5, E6
        TotalValueHelper totalVAR   = new TotalValueHelper();  // To accumulate total for A5, D5, E6
        
        
        Map<String, Map<String, ValueWithDecimalAndString>> P1S1desccdValues = new HashMap<>();
        String[] metrics = { "MTD", "MBDG", "VAR1", "VARP", "LY", "VAR", "VARLY"};
        
        for (String desccd : Arrays.asList("A5", "D5", "E6", "R5","E5","K5","M5","CALC1","CALC2","CALC3","CALC4")) {
        
            Map<String, ValueWithDecimalAndString> valuesMap = new HashMap<>();
            for (String metric : metrics) {
                valuesMap.put(metric, new ValueWithDecimalAndString(BigDecimal.ZERO));
            }
            P1S1desccdValues.put(desccd, valuesMap);
        }
    
            
        Map<String, String> P1S1desccdDescriptions = new HashMap<>();
        
        for (Object[] row : result) {
            P1S1desccdDescriptions.put((String) row[0], (String) row[1]);
        
            String desccd    = (String) row[0];
            BigDecimal mtd   = (BigDecimal) row[2];  // Assuming MTD is in the third column
            BigDecimal mbdg  = (BigDecimal) row[3];  // Assuming MBDG is in the fourth column
            BigDecimal var1  = (BigDecimal) row[4];  // Assuming VAR1 is in the fifth column
            BigDecimal ly    = (BigDecimal) row[6];  // Assuming LY is in the seventh column
            BigDecimal var   = (BigDecimal) row[7];  // Assuming VAR is in the eighth column
            
            
            mtd = mtd.setScale(0, RoundingMode.HALF_UP);
            mbdg = mbdg.setScale(0, RoundingMode.HALF_UP);
            var1 = var1.setScale(0, RoundingMode.HALF_UP);
            ly = ly.setScale(0, RoundingMode.HALF_UP);
            var = var.setScale(0, RoundingMode.HALF_UP);
            
            if (desccd != null && (desccd.trim().equals("A5") || desccd.trim().equals("D5") || desccd.trim().equals("E6"))) {
            
                Map<String, ValueWithDecimalAndString> valuesMap = P1S1desccdValues.get(desccd.trim());

                
                
                valuesMap.put("MTD",new ValueWithDecimalAndString(mtd));
                valuesMap.put("MBDG",new ValueWithDecimalAndString(mbdg));
                valuesMap.put("VAR1",new ValueWithDecimalAndString(var1));

                
                if (mbdg.compareTo(BigDecimal.ZERO) != 0) { // Prevent division by zero
                    BigDecimal result1 = var1.divide(mbdg, 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
                
                    valuesMap.put("VARP", new ValueWithDecimalAndString(result1));
                } else {
                    valuesMap.put("VARP", new ValueWithDecimalAndString(BigDecimal.ZERO));
                }
                
                valuesMap.put("LY",new ValueWithDecimalAndString(ly));
                valuesMap.put("VAR",new ValueWithDecimalAndString(var));
                

                if (ly.compareTo(BigDecimal.ZERO) != 0) { // Prevent division by zero
                    BigDecimal result1 = var.divide(ly, 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
                
                    valuesMap.put("VARLY", new ValueWithDecimalAndString(result1));
                } else {
                    valuesMap.put("VARLY", new ValueWithDecimalAndString(BigDecimal.ZERO));
                }
                

                totalMTD.addToTotal(mtd);  
                totalMBDG.addToTotal(mbdg);  
                totalVAR1.addToTotal(var1);  
                totalLY.addToTotal(ly);  
                totalVAR.addToTotal(var);  
                
          

            }
            
            

            if (desccd != null && (desccd.trim().equals("E5") || desccd.trim().equals("M5") || desccd.trim().equals("K5") || desccd.trim().equals("R5"))) {
                //Map<String, ValueWithDecimalAndString>    valuesMap = P1S1desccdValues.get("R5");
                Map<String, ValueWithDecimalAndString> valuesMap = P1S1desccdValues.get(desccd.trim());
                valuesMap.put("MTD",new ValueWithDecimalAndString(mtd));
                valuesMap.put("MBDG",new ValueWithDecimalAndString(mbdg));
                valuesMap.put("VAR1",new ValueWithDecimalAndString(var1));

                if (mbdg.compareTo(BigDecimal.ZERO) != 0) { 
                    BigDecimal result1 = var1.divide(mbdg, 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
                
                    valuesMap.put("VARP", new ValueWithDecimalAndString(result1));
                } else {
                    valuesMap.put("VARP", new ValueWithDecimalAndString(BigDecimal.ZERO));
                }
                valuesMap.put("LY",new ValueWithDecimalAndString(ly));
                valuesMap.put("VAR",new ValueWithDecimalAndString(var));

                if (ly.compareTo(BigDecimal.ZERO) != 0) { // Prevent division by zero
                    BigDecimal result1 = var.divide(ly, 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
                

                    valuesMap.put("VARLY", new ValueWithDecimalAndString(result1));
                } else {

                    valuesMap.put("VARLY", new ValueWithDecimalAndString(BigDecimal.ZERO));
                }
            }
            
        }
        
        //System.out.println ("descValues K5 MTD 1= " + P1S1desccdValues.get("R5").get("MTD").getValue().multiply(BigDecimal.valueOf(100)));

        //System.out.println ("descValues K5 MTD 2 = " + P1S1desccdValues.get("R5").get("MTD").getValue().multiply(BigDecimal.valueOf(100)).divide(totalMTD.getTotalValue(), 3, RoundingMode.HALF_UP));
        

        //Map<String, Map<String, ValueWithDecimalAndString>> P1S1desccdValues = new HashMap<>();
        //P1S1desccdValues.putIfAbsent("CALC2", new HashMap<>());  
        //P1S1desccdValues.get("CALC2").put("MTD", new ValueWithDecimalAndString(new BigDecimal("123.45")));
        P1S1desccdDescriptions.put((String) "CALC2", "PBT % on Total Revenue");
        P1S1desccdValues.get("CALC2").put("MTD", new ValueWithDecimalAndString(P1S1desccdValues.get("R5").get("MTD").getValue().multiply(BigDecimal.valueOf(100)).divide(totalMTD.getTotalValue(), 1, RoundingMode.HALF_UP)));
        P1S1desccdValues.get("CALC2").put("MBDG", new ValueWithDecimalAndString(P1S1desccdValues.get("R5").get("MBDG").getValue().multiply(BigDecimal.valueOf(100)).divide(totalMBDG.getTotalValue(), 1, RoundingMode.HALF_UP)));
        P1S1desccdValues.get("CALC2").put("VAR1", new ValueWithDecimalAndString( P1S1desccdValues.get("CALC2").get("MTD").getValue().subtract(P1S1desccdValues.get("CALC2").get("MBDG").getValue())     ));
        P1S1desccdValues.get("CALC2").put("LY", new ValueWithDecimalAndString(P1S1desccdValues.get("R5").get("LY").getValue().multiply(BigDecimal.valueOf(100)).divide(totalLY.getTotalValue(), 1, RoundingMode.HALF_UP)));
        P1S1desccdValues.get("CALC2").put("VAR", new ValueWithDecimalAndString( P1S1desccdValues.get("CALC2").get("MTD").getValue().subtract(P1S1desccdValues.get("CALC2").get("LY").getValue())     ));


        
        P1S1desccdDescriptions.put((String) "CALC3", "Cash PBT - Operations");
        P1S1desccdValues.get("CALC3").put("MTD", new ValueWithDecimalAndString(P1S1desccdValues.get("M5").get("MTD").getValue().subtract(P1S1desccdValues.get("E5").get("MTD").getValue())));
        P1S1desccdValues.get("CALC3").put("MBDG", new ValueWithDecimalAndString(P1S1desccdValues.get("M5").get("MBDG").getValue().subtract(P1S1desccdValues.get("E5").get("MBDG").getValue())));
        P1S1desccdValues.get("CALC3").put("VAR1", new ValueWithDecimalAndString(P1S1desccdValues.get("M5").get("VAR1").getValue().subtract(P1S1desccdValues.get("E5").get("VAR1").getValue())));

        if (P1S1desccdValues.get("CALC3").get("MBDG").getValue().compareTo(BigDecimal.ZERO) != 0) { 
            BigDecimal result1 = P1S1desccdValues.get("CALC3").get("VAR1").getValue().divide(P1S1desccdValues.get("CALC3").get("MBDG").getValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            P1S1desccdValues.get("CALC3").put("VARP", new ValueWithDecimalAndString(result1));
        } else 
        {
            P1S1desccdValues.get("CALC3").put("VARP", new ValueWithDecimalAndString(BigDecimal.ZERO));
            
        }

        P1S1desccdValues.get("CALC3").put("LY", new ValueWithDecimalAndString(P1S1desccdValues.get("M5").get("LY").getValue().subtract(P1S1desccdValues.get("E5").get("LY").getValue())));
        P1S1desccdValues.get("CALC3").put("VAR", new ValueWithDecimalAndString(P1S1desccdValues.get("M5").get("VAR").getValue().subtract(P1S1desccdValues.get("E5").get("VAR").getValue())));

        if (P1S1desccdValues.get("CALC3").get("LY").getValue().compareTo(BigDecimal.ZERO) != 0) { 
            BigDecimal result1 = P1S1desccdValues.get("CALC3").get("VAR").getValue().divide(P1S1desccdValues.get("CALC3").get("LY").getValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            P1S1desccdValues.get("CALC3").put("VARLY", new ValueWithDecimalAndString(result1));
        } else 
        {
            P1S1desccdValues.get("CALC3").put("VARLY", new ValueWithDecimalAndString(BigDecimal.ZERO));
            
        }

        
        P1S1desccdDescriptions.put((String) "CALC4", "Cash PBT % on Total Revenue");


        if (totalMTD.getTotalValue().compareTo(BigDecimal.ZERO) != 0) { 
            BigDecimal result1 = P1S1desccdValues.get("CALC3").get("MTD").getValue().divide(totalMTD.getTotalValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            P1S1desccdValues.get("CALC4").put("MTD", new ValueWithDecimalAndString(result1));
        } else 
        {
            P1S1desccdValues.get("CALC4").put("MTD", new ValueWithDecimalAndString(BigDecimal.ZERO));
            
        }

        if (totalMBDG.getTotalValue().compareTo(BigDecimal.ZERO) != 0) { 
            BigDecimal result1 = P1S1desccdValues.get("CALC3").get("MBDG").getValue().divide(totalMBDG.getTotalValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            P1S1desccdValues.get("CALC4").put("MBDG", new ValueWithDecimalAndString(result1));
        } else 
        {
            P1S1desccdValues.get("CALC4").put("MBDG", new ValueWithDecimalAndString(BigDecimal.ZERO));
            
        }

        P1S1desccdValues.get("CALC4").put("VAR1", new ValueWithDecimalAndString(P1S1desccdValues.get("CALC4").get("MTD").getValue().subtract(P1S1desccdValues.get("CALC4").get("MBDG").getValue())));
        
        if (totalLY.getTotalValue().compareTo(BigDecimal.ZERO) != 0) { 
            BigDecimal result1 = P1S1desccdValues.get("CALC3").get("LY").getValue().divide(totalLY.getTotalValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            P1S1desccdValues.get("CALC4").put("LY", new ValueWithDecimalAndString(result1));
        } else 
        {
            P1S1desccdValues.get("CALC4").put("LY", new ValueWithDecimalAndString(BigDecimal.ZERO));
            
        }
        P1S1desccdValues.get("CALC4").put("VAR", new ValueWithDecimalAndString(P1S1desccdValues.get("CALC4").get("MTD").getValue().subtract(P1S1desccdValues.get("CALC4").get("LY").getValue())));

        


        


        P1S1desccdValues.get("CALC1").put("MTD", new ValueWithDecimalAndString(totalMTD.getTotalValue()));
        P1S1desccdValues.get("CALC1").put("MBDG", new ValueWithDecimalAndString(totalMBDG.getTotalValue()));
        P1S1desccdValues.get("CALC1").put("VAR1", new ValueWithDecimalAndString(totalVAR1.getTotalValue()));


        //model.addAttribute("totalMTD",     totalMTD);
        //model.addAttribute("totalMBDG",    totalMBDG);
        //model.addAttribute("totalVAR1",    totalVAR1);



        if (totalMBDG.getTotalValue().compareTo(BigDecimal.ZERO) != 0) { // Prevent division by zero
            BigDecimal result1 = totalVAR1.getTotalValue().divide(totalMBDG.getTotalValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            //model.addAttribute("totaVARP",     new ValueWithDecimalAndString(result1));
            P1S1desccdValues.get("CALC1").put("VARP", new ValueWithDecimalAndString(result1));
        } else {
            //model.addAttribute("totaVARP",   0  );
            P1S1desccdValues.get("CALC1").put("VARP", new ValueWithDecimalAndString(BigDecimal.ZERO));

        }

        //model.addAttribute("totalLY",      totalLY);
        //model.addAttribute("totalVAR",     totalVAR);
        P1S1desccdValues.get("CALC1").put("LY", new ValueWithDecimalAndString(totalLY.getTotalValue()));
        P1S1desccdValues.get("CALC1").put("VAR", new ValueWithDecimalAndString(totalVAR.getTotalValue()));

        if (totalLY.getTotalValue().compareTo(BigDecimal.ZERO) != 0) { // Prevent division by zero
            BigDecimal result1 = totalVAR.getTotalValue().divide(totalLY.getTotalValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            //model.addAttribute("totalVARLY",     new ValueWithDecimalAndString(result1));
            P1S1desccdValues.get("CALC1").put("VARLY", new ValueWithDecimalAndString(result1));
        } else {
            //model.addAttribute("totalVARLY",   0  );
            P1S1desccdValues.get("CALC1").put("VARLY", new ValueWithDecimalAndString(BigDecimal.ZERO));
        }

        Map<String, Object> response = new HashMap<>();
        response.put("P1S1desccdValues", P1S1desccdValues);
        response.put("P1S1desccdDescriptions", P1S1desccdDescriptions);

        return response;
    }


    public Map<String, Object> getPage1B(String month, String year) 
    {
        // Fetch the results
        System.out.println("year = " + year);
        System.out.println("month = " + month);
        List<Object[]> result = mistxnRepository.FlashMISPage1B(year, month);
    
        System.out.println("deeeeeeeeeeeeeeeeeeeeeeeeeeeeeee.....................");
        TotalValueHelper totalMTD   = new TotalValueHelper();  // To accumulate total for A5, D5, E6
        TotalValueHelper totalMBDG  = new TotalValueHelper();  // To accumulate total for A5, D5, E6
        TotalValueHelper totalVAR1  = new TotalValueHelper();  // To accumulate total for A5, D5, E6
        
        TotalValueHelper totalLY    = new TotalValueHelper();  // To accumulate total for A5, D5, E6
        TotalValueHelper totalVAR   = new TotalValueHelper();  // To accumulate total for A5, D5, E6
        
        
        Map<String, Map<String, ValueWithDecimalAndString>> P1S1desccdValues = new HashMap<>();
        String[] metrics = { "MTD", "MBDG", "VAR1", "VARP", "LY", "VAR", "VARLY"};
        
        for (String desccd : Arrays.asList("A5", "D5", "E6", "R5","E5","K5","M5","CALC1","CALC2","CALC3","CALC4")) {
        
            Map<String, ValueWithDecimalAndString> valuesMap = new HashMap<>();
            for (String metric : metrics) {
                valuesMap.put(metric, new ValueWithDecimalAndString(BigDecimal.ZERO));
            }
            P1S1desccdValues.put(desccd, valuesMap);
        }
    
            
        Map<String, String> P1S1desccdDescriptions = new HashMap<>();
        
        for (Object[] row : result) {
            P1S1desccdDescriptions.put((String) row[0], (String) row[1]);
        
            String desccd    = (String) row[0];
            BigDecimal mtd   = (BigDecimal) row[2];  // Assuming MTD is in the third column
            BigDecimal mbdg  = (BigDecimal) row[3];  // Assuming MBDG is in the fourth column
            BigDecimal var1  = (BigDecimal) row[4];  // Assuming VAR1 is in the fifth column
            BigDecimal ly    = (BigDecimal) row[6];  // Assuming LY is in the seventh column
            BigDecimal var   = (BigDecimal) row[7];  // Assuming VAR is in the eighth column
            
            
            mtd = mtd.setScale(0, RoundingMode.HALF_UP);
            mbdg = mbdg.setScale(0, RoundingMode.HALF_UP);
            var1 = var1.setScale(0, RoundingMode.HALF_UP);
            ly = ly.setScale(0, RoundingMode.HALF_UP);
            var = var.setScale(0, RoundingMode.HALF_UP);
            
            if (desccd != null && (desccd.trim().equals("A5") || desccd.trim().equals("D5") || desccd.trim().equals("E6"))) {
            
                Map<String, ValueWithDecimalAndString> valuesMap = P1S1desccdValues.get(desccd.trim());

                
                
                valuesMap.put("MTD",new ValueWithDecimalAndString(mtd));
                valuesMap.put("MBDG",new ValueWithDecimalAndString(mbdg));
                valuesMap.put("VAR1",new ValueWithDecimalAndString(var1));

                
                if (mbdg.compareTo(BigDecimal.ZERO) != 0) { // Prevent division by zero
                    BigDecimal result1 = var1.divide(mbdg, 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
                
                    valuesMap.put("VARP", new ValueWithDecimalAndString(result1));
                } else {
                    valuesMap.put("VARP", new ValueWithDecimalAndString(BigDecimal.ZERO));
                }
                
                valuesMap.put("LY",new ValueWithDecimalAndString(ly));
                valuesMap.put("VAR",new ValueWithDecimalAndString(var));
                

                if (ly.compareTo(BigDecimal.ZERO) != 0) { // Prevent division by zero
                    BigDecimal result1 = var.divide(ly, 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
                
                    valuesMap.put("VARLY", new ValueWithDecimalAndString(result1));
                } else {
                    valuesMap.put("VARLY", new ValueWithDecimalAndString(BigDecimal.ZERO));
                }
                

                totalMTD.addToTotal(mtd);  
                totalMBDG.addToTotal(mbdg);  
                totalVAR1.addToTotal(var1);  
                totalLY.addToTotal(ly);  
                totalVAR.addToTotal(var);  
                
          

            }
            
            

            if (desccd != null && (desccd.trim().equals("E5") || desccd.trim().equals("M5") || desccd.trim().equals("K5") || desccd.trim().equals("R5"))) {
                //Map<String, ValueWithDecimalAndString>    valuesMap = P1S1desccdValues.get("R5");
                Map<String, ValueWithDecimalAndString> valuesMap = P1S1desccdValues.get(desccd.trim());
                valuesMap.put("MTD",new ValueWithDecimalAndString(mtd));
                valuesMap.put("MBDG",new ValueWithDecimalAndString(mbdg));
                valuesMap.put("VAR1",new ValueWithDecimalAndString(var1));

                if (mbdg.compareTo(BigDecimal.ZERO) != 0) { 
                    BigDecimal result1 = var1.divide(mbdg, 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
                
                    valuesMap.put("VARP", new ValueWithDecimalAndString(result1));
                } else {
                    valuesMap.put("VARP", new ValueWithDecimalAndString(BigDecimal.ZERO));
                }
                valuesMap.put("LY",new ValueWithDecimalAndString(ly));
                valuesMap.put("VAR",new ValueWithDecimalAndString(var));

                if (ly.compareTo(BigDecimal.ZERO) != 0) { // Prevent division by zero
                    BigDecimal result1 = var.divide(ly, 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
                

                    valuesMap.put("VARLY", new ValueWithDecimalAndString(result1));
                } else {

                    valuesMap.put("VARLY", new ValueWithDecimalAndString(BigDecimal.ZERO));
                }
            }
            
        }
        System.out.println ("descValues MTD = " + P1S1desccdValues.get("A5").get("MTD").getValue());
        //System.out.println ("descValues K5 MTD 1= " + P1S1desccdValues.get("R5").get("MTD").getValue().multiply(BigDecimal.valueOf(100)));

        //System.out.println ("descValues K5 MTD 2 = " + P1S1desccdValues.get("R5").get("MTD").getValue().multiply(BigDecimal.valueOf(100)).divide(totalMTD.getTotalValue(), 3, RoundingMode.HALF_UP));
        

        //Map<String, Map<String, ValueWithDecimalAndString>> P1S1desccdValues = new HashMap<>();
        //P1S1desccdValues.putIfAbsent("CALC2", new HashMap<>());  
        //P1S1desccdValues.get("CALC2").put("MTD", new ValueWithDecimalAndString(new BigDecimal("123.45")));
        P1S1desccdDescriptions.put((String) "CALC2", "PBT % on Total Revenue");
        P1S1desccdValues.get("CALC2").put("MTD", new ValueWithDecimalAndString(P1S1desccdValues.get("R5").get("MTD").getValue().multiply(BigDecimal.valueOf(100)).divide(totalMTD.getTotalValue(), 1, RoundingMode.HALF_UP)));
        P1S1desccdValues.get("CALC2").put("MBDG", new ValueWithDecimalAndString(P1S1desccdValues.get("R5").get("MBDG").getValue().multiply(BigDecimal.valueOf(100)).divide(totalMBDG.getTotalValue(), 1, RoundingMode.HALF_UP)));
        P1S1desccdValues.get("CALC2").put("VAR1", new ValueWithDecimalAndString( P1S1desccdValues.get("CALC2").get("MTD").getValue().subtract(P1S1desccdValues.get("CALC2").get("MBDG").getValue())     ));
        P1S1desccdValues.get("CALC2").put("LY", new ValueWithDecimalAndString(P1S1desccdValues.get("R5").get("LY").getValue().multiply(BigDecimal.valueOf(100)).divide(totalLY.getTotalValue(), 1, RoundingMode.HALF_UP)));
        P1S1desccdValues.get("CALC2").put("VAR", new ValueWithDecimalAndString( P1S1desccdValues.get("CALC2").get("MTD").getValue().subtract(P1S1desccdValues.get("CALC2").get("LY").getValue())     ));


        
        P1S1desccdDescriptions.put((String) "CALC3", "Cash PBT - Operations");
        P1S1desccdValues.get("CALC3").put("MTD", new ValueWithDecimalAndString(P1S1desccdValues.get("M5").get("MTD").getValue().subtract(P1S1desccdValues.get("E5").get("MTD").getValue())));
        P1S1desccdValues.get("CALC3").put("MBDG", new ValueWithDecimalAndString(P1S1desccdValues.get("M5").get("MBDG").getValue().subtract(P1S1desccdValues.get("E5").get("MBDG").getValue())));
        P1S1desccdValues.get("CALC3").put("VAR1", new ValueWithDecimalAndString(P1S1desccdValues.get("M5").get("VAR1").getValue().subtract(P1S1desccdValues.get("E5").get("VAR1").getValue())));

        if (P1S1desccdValues.get("CALC3").get("MBDG").getValue().compareTo(BigDecimal.ZERO) != 0) { 
            BigDecimal result1 = P1S1desccdValues.get("CALC3").get("VAR1").getValue().divide(P1S1desccdValues.get("CALC3").get("MBDG").getValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            P1S1desccdValues.get("CALC3").put("VARP", new ValueWithDecimalAndString(result1));
        } else 
        {
            P1S1desccdValues.get("CALC3").put("VARP", new ValueWithDecimalAndString(BigDecimal.ZERO));
            
        }

        P1S1desccdValues.get("CALC3").put("LY", new ValueWithDecimalAndString(P1S1desccdValues.get("M5").get("LY").getValue().subtract(P1S1desccdValues.get("E5").get("LY").getValue())));
        P1S1desccdValues.get("CALC3").put("VAR", new ValueWithDecimalAndString(P1S1desccdValues.get("M5").get("VAR").getValue().subtract(P1S1desccdValues.get("E5").get("VAR").getValue())));

        if (P1S1desccdValues.get("CALC3").get("LY").getValue().compareTo(BigDecimal.ZERO) != 0) { 
            BigDecimal result1 = P1S1desccdValues.get("CALC3").get("VAR").getValue().divide(P1S1desccdValues.get("CALC3").get("LY").getValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            P1S1desccdValues.get("CALC3").put("VARLY", new ValueWithDecimalAndString(result1));
        } else 
        {
            P1S1desccdValues.get("CALC3").put("VARLY", new ValueWithDecimalAndString(BigDecimal.ZERO));
            
        }

        
        P1S1desccdDescriptions.put((String) "CALC4", "Cash PBT % on Total Revenue");


        if (totalMTD.getTotalValue().compareTo(BigDecimal.ZERO) != 0) { 
            BigDecimal result1 = P1S1desccdValues.get("CALC3").get("MTD").getValue().divide(totalMTD.getTotalValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            P1S1desccdValues.get("CALC4").put("MTD", new ValueWithDecimalAndString(result1));
        } else 
        {
            P1S1desccdValues.get("CALC4").put("MTD", new ValueWithDecimalAndString(BigDecimal.ZERO));
            
        }

        if (totalMBDG.getTotalValue().compareTo(BigDecimal.ZERO) != 0) { 
            BigDecimal result1 = P1S1desccdValues.get("CALC3").get("MBDG").getValue().divide(totalMBDG.getTotalValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            P1S1desccdValues.get("CALC4").put("MBDG", new ValueWithDecimalAndString(result1));
        } else 
        {
            P1S1desccdValues.get("CALC4").put("MBDG", new ValueWithDecimalAndString(BigDecimal.ZERO));
            
        }

        P1S1desccdValues.get("CALC4").put("VAR1", new ValueWithDecimalAndString(P1S1desccdValues.get("CALC4").get("MTD").getValue().subtract(P1S1desccdValues.get("CALC4").get("MBDG").getValue())));
        
        if (totalLY.getTotalValue().compareTo(BigDecimal.ZERO) != 0) { 
            BigDecimal result1 = P1S1desccdValues.get("CALC3").get("LY").getValue().divide(totalLY.getTotalValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            P1S1desccdValues.get("CALC4").put("LY", new ValueWithDecimalAndString(result1));
        } else 
        {
            P1S1desccdValues.get("CALC4").put("LY", new ValueWithDecimalAndString(BigDecimal.ZERO));
            
        }
        P1S1desccdValues.get("CALC4").put("VAR", new ValueWithDecimalAndString(P1S1desccdValues.get("CALC4").get("MTD").getValue().subtract(P1S1desccdValues.get("CALC4").get("LY").getValue())));

        


        


        P1S1desccdValues.get("CALC1").put("MTD", new ValueWithDecimalAndString(totalMTD.getTotalValue()));
        P1S1desccdValues.get("CALC1").put("MBDG", new ValueWithDecimalAndString(totalMBDG.getTotalValue()));
        P1S1desccdValues.get("CALC1").put("VAR1", new ValueWithDecimalAndString(totalVAR1.getTotalValue()));


        //model.addAttribute("totalMTD",     totalMTD);
        //model.addAttribute("totalMBDG",    totalMBDG);
        //model.addAttribute("totalVAR1",    totalVAR1);



        if (totalMBDG.getTotalValue().compareTo(BigDecimal.ZERO) != 0) { // Prevent division by zero
            BigDecimal result1 = totalVAR1.getTotalValue().divide(totalMBDG.getTotalValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            //model.addAttribute("totaVARP",     new ValueWithDecimalAndString(result1));
            P1S1desccdValues.get("CALC1").put("VARP", new ValueWithDecimalAndString(result1));
        } else {
            //model.addAttribute("totaVARP",   0  );
            P1S1desccdValues.get("CALC1").put("VARP", new ValueWithDecimalAndString(BigDecimal.ZERO));

        }

        //model.addAttribute("totalLY",      totalLY);
        //model.addAttribute("totalVAR",     totalVAR);
        P1S1desccdValues.get("CALC1").put("LY", new ValueWithDecimalAndString(totalLY.getTotalValue()));
        P1S1desccdValues.get("CALC1").put("VAR", new ValueWithDecimalAndString(totalVAR.getTotalValue()));

        if (totalLY.getTotalValue().compareTo(BigDecimal.ZERO) != 0) { // Prevent division by zero
            BigDecimal result1 = totalVAR.getTotalValue().divide(totalLY.getTotalValue(), 3, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100)); 
            //model.addAttribute("totalVARLY",     new ValueWithDecimalAndString(result1));
            P1S1desccdValues.get("CALC1").put("VARLY", new ValueWithDecimalAndString(result1));
        } else {
            //model.addAttribute("totalVARLY",   0  );
            P1S1desccdValues.get("CALC1").put("VARLY", new ValueWithDecimalAndString(BigDecimal.ZERO));
        }

        Map<String, Object> response = new HashMap<>();
        response.put("P1S2desccdValues", P1S1desccdValues);
        response.put("P1S2desccdDescriptions", P1S1desccdDescriptions);

        return response;
    }
}
