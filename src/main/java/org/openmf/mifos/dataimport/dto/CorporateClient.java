package org.openmf.mifos.dataimport.dto;

import java.util.Locale;

public class CorporateClient {

	 private final transient Integer rowIndex;
	    
	    private final String dateFormat;
	    
	    private final Locale locale;
	    
	    private final String officeId;
	    
	    private final String staffId;
	    
	    private final String fullname;
	    
	    private final String externalId;
	    
	    private final String active;
	    
	    private final String activationDate;
	    
	    
	    public CorporateClient(String fullname, String activationDate, String active, String externalId, String officeId, String staffId, Integer rowIndex ) {
	        if(fullname == null || fullname.trim().equals("")) {
	            throw new IllegalArgumentException("Name can not be empty.");
	        }
	        this.fullname = fullname;
	        this.activationDate = activationDate;
	        this.active = active;
	        this.externalId = externalId;
	        this.officeId = officeId;
	        this.staffId = staffId;
	        this.rowIndex = rowIndex;
	        this.dateFormat = "dd MMMM yyyy";
	        this.locale = Locale.ENGLISH;
	    }
	    
	    public String getFullName() {
	        return this.fullname;
	    }
	    
	    public String getActivationDate() {
	        return this.activationDate;
	    }
	    
	    public String isActive() {
	        return this.active;
	    }
	    
	    public String getExternalId() {
	        return this.externalId;
	    }
	    
	    public String getOfficeId() {
	        return this.officeId;
	    }
	    
	    public String getStaffId() {
	        return this.staffId;
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
