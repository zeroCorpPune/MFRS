package com.rcs.mfrs.Model;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Entity
@Getter
@Setter
@NoArgsConstructor
public class Usermaster {
    
    
    @Id
    @Column(name = "cuserid", insertable=false, unique = true,updatable=false, nullable=false, length = 10)
    private String cuserid;
  
    

    @Column(name = "password", nullable = false)
    private String password;

    @Column(name = "canadd", length = 1)
    private String canadd;

    @Column(name = "canmodify", length = 1)
    private String canmodify;

    @Column(name = "candelete", length = 1)
    private String candelete;

    @Column(name = "canview", length = 1)
    private String canview;

    @Column(name = "cactive", length = 1)
    private String cactive;

    @Column(name = "userrole", length = 15)
    private String userrole;
    
    
   @ManyToOne
   @JoinColumn(name = "ccompcode", nullable = false)
   private Company company;

}
