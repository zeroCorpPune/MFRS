package com.rcs.mfrs.Model;

import javax.validation.constraints.NotNull;
import javax.validation.constraints.Size;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.persistence.PrePersist;
import javax.persistence.PreUpdate;
import javax.validation.constraints.DecimalMax;
import javax.validation.constraints.DecimalMin;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

import javax.persistence.Id;
import javax.persistence.IdClass;
import javax.persistence.Entity;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Entity
@Getter
@Setter
@IdClass(Flash_remId.class)
@NoArgsConstructor
public class Flash_rem {
    @Id
    private String ccompcode;

    @Id
    private String cmonth;

    @Id
    private String cyear;

    @Size(max = 1500)
    private String remarks;

    private LocalDate dcreationdate;
    private LocalDate dmodifieddate;

    @ManyToOne
    @JoinColumn(name = "ccompcode", insertable = false, updatable = false, nullable = false)
    private Company flash_remcompany;


    @PrePersist
    public void prePersist() {
        dcreationdate = LocalDate.now();  // Set creation date before insert
    }

    @PreUpdate
    public void preUpdate() {
        dmodifieddate = LocalDate.now();  // Set modification date before update
    }
}
