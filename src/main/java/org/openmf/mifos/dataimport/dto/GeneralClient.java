package org.openmf.mifos.dataimport.dto;

import java.util.ArrayList;
import com.google.gson.annotations.SerializedName;


public class GeneralClient {
	
	@SerializedName("id")
    private final Integer id;
	
	@SerializedName("displayName")
    private final String displayName;
    
	@SerializedName("officeName")
    private final String officeName;
    
	@SerializedName("activationDate")
    private final ArrayList<Integer> activationDate;

	public GeneralClient(Integer id, String displayName,  String officeName, ArrayList<Integer> activationDate) {
		this.id = id;
        this.displayName = displayName;
        this.activationDate = activationDate;
        this.officeName = officeName;
    }
	
	public Integer getId() {
		return this.id;
	}
	
	public String getDisplayName() {
        return this.displayName;
    }
    
    public ArrayList<Integer> getActivationDate() {
        return this.activationDate;
    }
    
    public String getOfficeName() {
        return this.officeName;
    }
}
