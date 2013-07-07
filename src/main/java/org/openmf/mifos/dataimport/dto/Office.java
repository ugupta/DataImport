package org.openmf.mifos.dataimport.dto;

import java.util.Locale;


public class Office {
    
    private final transient Integer rowIndex;
    
    private final String name;
    
    private final String parentId;
    
    private final String dateFormat="dd MMMM yyyy";
    
    private final Locale locale;
    
    private final String openingDate;
    
    private final String externalId;

    public Office(String name, String parentId, String externalId, String openingDate, Locale locale, Integer rowIndex ) {
        if(name == null || name.trim().equals("")) {
            throw new IllegalArgumentException("name can not be empty");
        }
        this.name = name;
        this.parentId = parentId;
        this.externalId = externalId;
        this.openingDate = openingDate;
        this.rowIndex = rowIndex;
        this.locale = locale;
    }

    public String getName() {
        return this.name;
    }

    public String getParentId() {
        return this.parentId;
    }

    public String getExternalId() {
        return this.externalId;
    }

    public String getOpeningDate() {
        return this.openingDate;
    }
    
    public Locale getLocale() {
    	return locale;
    }
    
    public String getDateFormat() {
    	return dateFormat;
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

}
