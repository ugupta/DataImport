package org.openmf.mifos.dataimport.dto;

import com.google.gson.annotations.SerializedName;

public class Product {

	@SerializedName("id")
    private final Integer id;
    
	@SerializedName("name")
    private final String name;
	
	@SerializedName("fundName")
	private final String fundName;
	
	@SerializedName("status")
	private final String status;
	
	public Product(Integer id, String name, String fundName, String status ) {
		this.id = id;
		this.name = name;
		this.fundName = fundName;
		this.status = status;
	}
	
	public Integer getId() {
    	return this.id;
    }

    public String getName() {
        return this.name;
    }
    
    public String getFundName() {
        return this.fundName;
    }
    
    public String getStatus() {
    	return this.status;
    }
}
