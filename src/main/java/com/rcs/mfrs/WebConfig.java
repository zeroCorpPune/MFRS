package com.rcs.mfrs;

import org.springframework.context.annotation.Configuration;
import org.springframework.format.FormatterRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;


@Configuration
public class WebConfig implements WebMvcConfigurer {

    @Override
    public void addFormatters(FormatterRegistry registry) {
        registry.addConverter(new StringToLocalDateConverter());
        registry.addFormatter(new LocalDateFormatter());
    }

    public class StringToLocalDateConverter implements org.springframework.core.convert.converter.Converter<String, LocalDate> {

        private DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

        @Override
        public LocalDate convert(String source) {
            try {
                return LocalDate.parse(source, formatter);
            } catch (DateTimeParseException e) {
                return null; // or throw an exception
            }
        }
    }

    public class LocalDateFormatter implements org.springframework.format.Formatter<LocalDate> {

        private DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

        @Override
        public LocalDate parse(String text, java.util.Locale locale) throws java.text.ParseException {
            try {
                return LocalDate.parse(text, formatter);
            } catch (DateTimeParseException e) {
                throw new java.text.ParseException(e.getMessage(), 0);
            }
        }

        @Override
        public String print(LocalDate object, java.util.Locale locale) {
            return object.format(formatter);
        }
    }
}
