package com.rcs.mfrs.Repositories;

import java.util.List;

import javax.transaction.Transactional;

import org.springframework.data.domain.Sort;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import com.rcs.mfrs.Model.Mistxn;
import com.rcs.mfrs.Model.MistxnId;

@Repository
public interface MistxnRepository extends JpaRepository<Mistxn, MistxnId>{
    
    List<Mistxn> findByCcompcodeAndCmonthAndCyear(String ccompcode, String cmonth, String cyear ,Sort sort);

    @Query("SELECT COUNT(m) FROM Mistxn m WHERE m.cyear = :year AND m.cmonth = :month AND m.ccompcode = :ccompcode")
    long countByMonthAndYear(@Param("year") String year, @Param("month") String mont, @Param("ccompcode") String ccompcode);


    @Query("SELECT COUNT(m) FROM Mistxn m WHERE cblock='Y' and  m.cyear = :year AND m.cmonth = :month AND m.ccompcode = :ccompcode")
    int flashSubmitStatus(@Param("year") String year, @Param("month") String mont, @Param("ccompcode") String ccompcode);

    /*
    sSql = " select "
	sSql = sSql & " DESCCD, DESCRIPTION,"
	sSql = sSql & " SUM(ACT_MTD_BAL) MTD,"
	sSql = sSql & " SUM(BGT_MTD_BAL) MBDG,"
	sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL) VAR1,"
	sSql = sSql & " SUM(ACT_MTD_BAL-BGT_MTD_BAL)/SUM(DECODE(BGT_MTD_BAL,0,1,BGT_MTD_BAL)) *100 VARP,"
	sSql = sSql & " SUM(ACT_MTD_BAL_PY1) LY,"
	sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1) VAR,"
	sSql = sSql & " SUM(ACT_MTD_BAL-ACT_MTD_BAL_PY1)/SUM(DECODE(ACT_MTD_BAL_PY1,0,1,ACT_MTD_BAL_PY1))*100 VARLY"
	sSql = sSql & " from mistxn"
	sSql = sSql & " where"
	sSql = sSql & " cyear='" & TheReportYear & "' and cmonth = '" & TheReportMonth & "' AND"
	sSql = sSql & " CCOMPCODE NOT IN ('OMHC','OMC','OFCC')"
	sSql = sSql & " AND DESCCD IN ('A5', 'D5','E5','E6','M5','K5','R5')"
	sSql = sSql & " GROUP BY CYEAR, CMONTH,"
	sSql = sSql & " DESCCD, DESCRIPTION"
	sSql = sSql & " ORDER BY 1"
     */

     /*@Query("select " +
     " m.desccd, m.description, " +                                 // Select DESCCD and DESCRIPTION
     " SUM(m.actMtdBal) as MTD, " +                                  // SUM(ACT_MTD_BAL) as MTD
     " SUM(m.bgtMtdBal) as MBDG, " +                                 // SUM(BGT_MTD_BAL) as MBDG
     " SUM(m.actMtdBal - m.bgtMtdBal) as VAR1, " +                   // SUM(ACT_MTD_BAL - BGT_MTD_BAL) as VAR1
     " SUM(m.actMtdBal - m.bgtMtdBal) / SUM(CASE WHEN m.bgtMtdBal = 0 THEN 1 ELSE m.bgtMtdBal END) * 100 as VARP, " +  // VARP calculation using CASE WHEN
     " SUM(m.actMtdBalPy1) as LY, " +                                // SUM(ACT_MTD_BAL_PY1) as LY
     " SUM(m.actMtdBal - m.actMtdBalPy1) as VAR, " +                 // SUM(ACT_MTD_BAL - ACT_MTD_BAL_PY1) as VAR
     " SUM(m.actMtdBal - m.actMtdBalPy1) / SUM(CASE WHEN m.actMtdBalPy1 = 0 THEN 1 ELSE m.actMtdBalPy1 END) * 100 as VARLY " + // VARLY calculation using CASE WHEN
     "from Mistxn m " +                                              // From table mistxn (represented by entity `Mistxn`)
     "where m.cyear = :year AND m.cmonth = :month " +                // Filter by year and month
     "AND m.ccompcode NOT IN ('OMHC', 'OMC', 'OFCC') " +             // Exclude specific company codes
     "AND m.desccd IN ('A5', 'D5', 'E5', 'E6', 'M5', 'K5', 'R5') " + // Filter by DESCCD values
     "GROUP BY m.cyear, m.cmonth, m.desccd, m.description " +        // Group by year, month, DESCCD, and description
     "ORDER BY m.desccd")                                            // Order by DESCCD
List<Object[]> FlashMISPage1(@Param("year") String year, @Param("month") String month);
*/

@Query("select " +
" m.desccd, m.description, " +                                 // Select DESCCD and DESCRIPTION
" SUM(m.actMtdBal) as MTD, " +                                  // SUM(ACT_MTD_BAL) as MTD
" SUM(m.bgtMtdBal) as MBDG, " +                                 // SUM(BGT_MTD_BAL) as MBDG
" SUM(m.actMtdBal - m.bgtMtdBal) as VAR1, " +                   // SUM(ACT_MTD_BAL - BGT_MTD_BAL) as VAR1
" SUM(m.actMtdBal - m.bgtMtdBal) / SUM(CASE WHEN m.bgtMtdBal = 0 THEN 1 ELSE m.bgtMtdBal END) * 100 as VARP, " +  // VARP calculation using CASE WHEN
" SUM(m.actMtdBalPy1) as LY, " +                                // SUM(ACT_MTD_BAL_PY1) as LY
" SUM(m.actMtdBal - m.actMtdBalPy1) as VAR, " +                 // SUM(ACT_MTD_BAL - ACT_MTD_BAL_PY1) as VAR
" SUM(m.actMtdBal - m.actMtdBalPy1) / SUM(CASE WHEN m.actMtdBalPy1 = 0 THEN 1 ELSE m.actMtdBalPy1 END) * 100 as VARLY " + // VARLY calculation using CASE WHEN
"from Mistxn m " +                                              // From table mistxn (represented by entity `Mistxn`)
"where m.cyear = :year AND m.cmonth = :month " +                // Filter by year and month
"AND m.ccompcode NOT IN ('OMHC', 'OMC', 'OFCC') " +             // Exclude specific company codes
"AND m.desccd IN ('A5', 'D5', 'E5', 'E6', 'M5', 'K5', 'R5') " + // Filter by DESCCD values
"GROUP BY m.cyear, m.cmonth, m.desccd, m.description " +        // Group by year, month, DESCCD, and description
"ORDER BY m.desccd")                                            // Order by DESCCD
List<Object[]> FlashMISPage1A(String year, String month);


 
@Query("select " +
" m.desccd, m.description, " +
" SUM(m.actYtdBal) as MTD, " +
" SUM(m.bgtYtdBal) as MBDG, " +
" SUM(m.actYtdBal - m.bgtYtdBal) as VAR1 ," +
" SUM(m.actYtdBal - m.bgtYtdBal) / SUM(CASE WHEN m.bgtYtdBal = 0 THEN 1 ELSE m.bgtYtdBal END) * 100 as VARP, " +  // VARP calculation using CASE WHEN
" SUM(m.actYtdBalPy1) as LY, " +
" SUM(m.actYtdBal - m.actYtdBalPy1) as VAR, " +
" SUM(m.actYtdBal - m.actYtdBalPy1) / SUM(CASE WHEN m.actYtdBalPy1 = 0 THEN 1 ELSE m.actYtdBalPy1 END) * 100 as VARLY " + // VARLY calculation using CASE WHEN
" from Mistxn m " +
" where " +
" m.cyear = :year AND m.cmonth = :month " +
" AND m.ccompcode NOT IN ('OMHC','OMC','OFCC') " +
" AND m.desccd IN ('A5', 'D5','E5','E6','M5','K5','R5') " +
" GROUP BY m.cyear, m.cmonth,m.desccd, m.description " +
" ORDER BY m.desccd")
List<Object[]> FlashMISPage1B(String year, String month);

    //@Query("Update Mistxn m Set cblock='Y' WHERE m.cyear = :year AND m.cmonth = :month AND m.ccompcode = :ccompcode")
    //void submitFlash(@Param("year") String year, @Param("month") String mont, @Param("ccompcode") String ccompcode);

    @Modifying
    @Transactional
    @Query("UPDATE Mistxn m SET m.cblock = 'Y' WHERE m.cyear = :year AND m.cmonth = :month AND m.ccompcode = :ccompcode")
    int updateMistxnStatus(@Param("year") String year, @Param("month") String month, @Param("ccompcode") String ccompcode);
/*

SELECT a.CCOMPCODE,a.CMONTH,a.CYEAR,a.desccd,a.BOLDYN,a.description,a.ACT_MTD_BAL,0,0,a.ACT_YTD_BAL,0,
(SELECT ACT_YTD_BAL  FROM MISTXN b
WHERE Upper(Trim(cMonth)) = '09' And Upper(Trim(cYear)) = '2022' 
And Upper(Trim(cCompCode)) = a.ccompcode
and a.desccd = b.desccd
) as act_ytd_bal_py1_2022,
(SELECT ACT_YTD_BAL  FROM MISTXN b
WHERE Upper(Trim(cMonth)) = '09' And Upper(Trim(cYear)) = '2021' 
And Upper(Trim(cCompCode)) = a.ccompcode
and a.desccd = b.desccd
) as act_ytd_bal_py2_2021,
(SELECT ACT_YTD_BAL  FROM MISTXN b
WHERE Upper(Trim(cMonth)) = '09' And Upper(Trim(cYear)) = '2020' 
And Upper(Trim(cCompCode)) = a.ccompcode
and a.desccd = b.desccd
) as act_ytd_bal_py3_2020
 FROM mistxn a
 WHERE a.cMonth = '09' And a.cYear = '2023' 
And a.cCompCode = 'RCS';
    @Query("SELECT COUNT(m) FROM Mistxn m WHERE m.cyear = :year AND m.cmonth = :month AND m.ccompcode = :ccompcode")
    long getByMonthAndYearandCompanycode(@Param("year") String year, @Param("month") String mont, @Param("ccompcode") String ccompcode);
*/
}
