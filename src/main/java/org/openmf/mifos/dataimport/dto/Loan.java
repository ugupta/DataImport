package org.openmf.mifos.dataimport.dto;

import java.util.Locale;

public class Loan {

	private final String amortizationType;
	
	private final String clientId;
	
	private final String dateFormat = "dd MMMM yyyy";
	
	private final String expectedDisbursementDate;
	
	private final String fundId;
	
	private final String graceOnInterestCharged;
	
	private final String graceOnInterestPayment;
	
	private final String graceOnPrincipalPayment;
	
	private final String inArrearsTolerance;
	
	private final String interestCalculationPeriodType;
	
	private final String interestChargedFromDate;
	
	private final String interestRatePerPeriod;
	
	private final String interestType;
	
	private final String loanOfficerId;
	
	private final String loanTermFrequency;
	
	private final String loanTermFrequencyType;
	
	private final String loanType;
	
	private final Locale locale = Locale.ENGLISH;
	
	private final String numberOfRepayments;
	
	private final String principal;
	
	private final String productId;
	
	private final String repaymentEvery;
	
	private final String repaymentEveryFrequencyType;
	
	private final String repaymentsStartingFromDate;
	
	private final String submittedOnDate;
	
	private final String transactionProcessingStrategyId;
	
	public Loan(String amortizationType, String clientId, String expectedDisbursementDate, String fundId, String graceOnInterestCharged, String graceOnInterestPayment, String graceOnPrincipalPayment,
			String inArrearsTolerance, String interestCalculationPeriodType, String interestChargedFromDate, String interestRatePerPeriod, String interestType, String loanOfficerId,
			String loanTermFrequency, String loanTermFrequencyType, String loanType, String numberOfRepayments, String principal, String productId, String repaymentEvery,
			String repaymentEveryFrequencyType, String repaymentsStartingFromDate, String submittedOnDate, String transactionProcessingStrategyId) {
		this.amortizationType = amortizationType;
		this.clientId = clientId;
		this.expectedDisbursementDate = expectedDisbursementDate;
		this.fundId = fundId;
		this.graceOnInterestCharged = graceOnInterestCharged;
		this.graceOnInterestPayment = graceOnInterestPayment;
		this.graceOnPrincipalPayment = graceOnPrincipalPayment;
		this.inArrearsTolerance = inArrearsTolerance;
		this.interestCalculationPeriodType = interestCalculationPeriodType;
		this.interestChargedFromDate = interestChargedFromDate;
		this.interestRatePerPeriod = interestRatePerPeriod;
		this.interestType = interestType;
		this.loanOfficerId = loanOfficerId;
		this.loanTermFrequency = loanTermFrequency;
		this.loanTermFrequencyType = loanTermFrequencyType;
		this.loanType = loanType;
		this.numberOfRepayments = numberOfRepayments;
		this.principal = principal;
		this.productId = productId;
		this.repaymentEvery = repaymentEvery;
		this.repaymentEveryFrequencyType = repaymentEveryFrequencyType;
		this.repaymentsStartingFromDate = repaymentsStartingFromDate;
		this.submittedOnDate = submittedOnDate;
		this.transactionProcessingStrategyId = transactionProcessingStrategyId;
	}
	
	public String getAmortizationType() {
		return amortizationType;
	}
	
	public String getClientId() {
		return clientId;
	}
	
	public String getExpectedDisbursementDate() {
		return expectedDisbursementDate;
	}
	
	public String getFundId() {
		return fundId;
	}
	
	public String getGraceOnInterestCharged() {
		return graceOnInterestCharged;
	}
	
	public String getGraceOnInterestPayment() {
		return graceOnInterestPayment;
	}
	
	public String getGraceOnPrincipalPayment() {
		return graceOnPrincipalPayment;
	}
	
	public String getInArrearsTolerance() {
		return inArrearsTolerance;
	}
	
	public String getInterestCalculationPeriodType() {
		return interestCalculationPeriodType;
	}
	
	public String getInterestChargedFromDate() {
		return interestChargedFromDate;
	}
	
	public String getInterestRatePerPeriod() {
		return interestRatePerPeriod;
	}
	
	public String getInterestType() {
		return interestType;
	}
	
	public String getLoanOfficerId() {
		return loanOfficerId;
	}
	
	public String getLoanTermFrequencyType() {
		return loanTermFrequencyType;
	}
	
	public String getLoanTermFrequency() {
		return loanTermFrequency;
	}
	
	public String getLoanType() {
		return loanType;
	}
	
	public String getNumberOfRepayments() {
		return numberOfRepayments;
	}
	
	public String getPrincipal() {
		return principal;
	}
	
	public String getProductId() {
		return productId;
	}
	
	public String getRepaymentEvery() {
		return repaymentEvery;
	}
	
	public String getRepaymentEveryFrequencyType() {
		return repaymentEveryFrequencyType;
	}
	
	public String getRepaymentsStartingFromDate() {
		return repaymentsStartingFromDate;
	}
	
	public String getSubmittedOnDate() {
		return submittedOnDate;
	}
	
	public String getTransactionProcessingStrategyId() {
		return transactionProcessingStrategyId;
	}
	
	public String getDateFormat() {
		return dateFormat;
	}
	
	public Locale getLocale() {
		return locale;
	}
	
}
