package org.openmf.mifos.dataimport.dto;

import java.util.Locale;

public class LoanApproval {
	
	private final transient Integer rowIndex;

	private final String approvedOnDate;
	
	private final String dateFormat="dd MMMM yyyy";
	
	 private final Locale locale = Locale.ENGLISH;
	 
	 private final String note = "";
	 
	 public LoanApproval(String approvedOnDate, Integer rowIndex ) {
	        this.approvedOnDate = approvedOnDate;
	        this.rowIndex = rowIndex;
	    }
	 
	  public String getApprovedOnDate() {
		    return approvedOnDate;
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
	    
	    public String getNote() {
	    	return note;
	    }
	    
}
