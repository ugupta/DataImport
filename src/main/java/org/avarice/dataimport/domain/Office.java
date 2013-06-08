package org.avarice.dataimport.domain;

import java.io.Serializable;
import java.util.Date;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "m_office")
public class Office implements Serializable{
	
	 /**
	 * 
	 */
	private static final long serialVersionUID = 12345L;

	@Id
	 @Column(name = "id")
	 @GeneratedValue
	 private Long id;
	  
	 @Column(name = "parent_id")
	 private Long parentId;
	  
	 @Column(name = "hierarchy")
	 private String hierarchy;
	  
	 @Column(name = "external_id")
	 private Long externalId;
	  
	 @Column(name = "name")
	 private String name;
	  
	 @Column(name = "opening_date")
	 private Date openingDate;
	 
	     public Long getId() {
		  return id;
		 }
		 
		 public void setId(Long id) {
		  this.id = id;
		 }
		 
		 public String getName() {
		  return name;
		 }
		 
		 public void setName(String name) {
		  this.name = name;
		 }
		 
		 public Long getParentId() {
		  return parentId;
		 }
		 
		 public void setParentId(Long parentId ) {
		  this.parentId = parentId;
		 }
		 
		 public Long getExternalId() {
		  return externalId;
		 }
		 
		 public void setExternalId(Long externalId) {
		  this.externalId = externalId;
		 }
		 
		 public String getHierarchy() {
		  return hierarchy;
		 }
		 
		 public void setHierarchy(String hierarchy) {
		  this.hierarchy = hierarchy;
		 }
		 
		 public Date getOpeningDate() {
		  return openingDate;
		 }
		 
		 public void setOpeningDate(Date openingDate) {
		  this.openingDate = openingDate;
		 }

}
