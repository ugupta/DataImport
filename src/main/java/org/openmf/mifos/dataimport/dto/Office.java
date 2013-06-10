package org.openmf.mifos.dataimport.dto;

import java.util.Date;


public class Office {
    
    private final transient Integer rowIndex;
    
    private final String name;
    
    private final Integer perentId;
    
    private final String externalId;
    
    private final Date openingDate;

    public Office(String name, Integer perentId, String externalId, Date openingDate, Integer rowIndex) {
        if(name == null || name.trim().equals("")) {
            throw new IllegalArgumentException("name can not be empty");
        }
        this.name = name;
        this.perentId = perentId;
        this.externalId = externalId;
        this.openingDate = (Date) openingDate.clone();
        this.rowIndex = rowIndex;
    }

    public String getName() {
        return this.name;
    }

    public Integer getPerentId() {
        return this.perentId;
    }

    public String getExternalId() {
        return this.externalId;
    }

    public Date getOpeningDate() {
        return (Date) this.openingDate.clone();
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

}
