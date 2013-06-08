package org.avarice.dataimport.domain;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;


@Service("officeBo")
public class OfficeBoImpl implements OfficeBo{
   
	@Autowired
	OfficeDao officeDao;
	 
	public void setofficeDao(OfficeDao officeDao) {
		this.officeDao = officeDao;
	}
	
	public void save(Office office){
		officeDao.save(office);
	}
	
	public Office getOfficeById(long id){
		return officeDao.getOfficeById(id);
	}
	
	public Office getOfficeByName(String name){
		return officeDao.getOfficeByName(name);
	}
}
