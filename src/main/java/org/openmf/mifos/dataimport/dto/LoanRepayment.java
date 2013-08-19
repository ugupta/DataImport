package org.openmf.mifos.dataimport.dto;

import java.util.Locale;

public class LoanRepayment {

	 private final transient Integer rowIndex;
	 
	 private final String transactionAmount;

	 private final String transactionDate;
	
	 private final String paymentTypeId;
	
	 private final String dateFormat="dd MMMM yyyy";
	
	 private final Locale locale = Locale.ENGLISH;
	 
	 private final String note = "";
	 
	 private final String accountNumber = "";
	 
	 private final String routingCode = "";
	 
	 private final String receiptNumber = "";
	 
	 private final String bankNumber = "";
	 
	 private final String checkNumber = "";
	 
	 public LoanRepayment(String transactionAmount, String transactionDate, String paymentTypeId, Integer rowIndex) {
		    this.transactionAmount = transactionAmount;
	        this.transactionDate = transactionDate;
	        this.paymentTypeId = paymentTypeId;
	        this.rowIndex = rowIndex;
	    }
	 
	    public String getTransactionAmount() {
	    	return transactionAmount;
	    }
	 
	    public String getTransactionDate() {
		    return transactionDate;
	    }
	 
	    public String getPaymentTypeId() {
		 return paymentTypeId;
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
	    
	    public String getAccountNumber() {
	    	return this.accountNumber;
	    }
	    
	    public String getRoutingCode() {
	    	return this.routingCode;
	    }
	    
	    public String getReceiptNumber() {
	    	return this.receiptNumber;
	    }
	    
	    public String getBankNumber() {
	    	return this.bankNumber;
	    }
	    
	    public String getCheckNumber() {
	    	return this.checkNumber;
	    }
}
