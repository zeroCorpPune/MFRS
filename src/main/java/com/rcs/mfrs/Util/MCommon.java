package com.rcs.mfrs.Util;

import java.time.Month;
import java.time.Year;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.springframework.stereotype.Component;


@Component
public class MCommon {

    public List<Integer> getYearsList(int range) {
        int currentYear = Year.now().getValue();
        return IntStream.rangeClosed(currentYear - range, currentYear)
                        .boxed()
                        .collect(Collectors.toList());
    }

    public Map<String, String> getMonthsList() {
        return IntStream.rangeClosed(1, 12)
                        .boxed()
                        .collect(Collectors.toMap(
                            month -> String.format("%02d", month), // Prefix with 0 if necessary
                            month -> Month.of(month).name(),
                            (oldValue, newValue) -> oldValue,
                            LinkedHashMap::new // To maintain order
                        ));
    }

    public int getCurrentYear() {
        return Year.now().getValue();
    }

    public String getCurrentMonth() {
        return String.format("%02d", java.time.LocalDate.now().getMonthValue());
    }
}
