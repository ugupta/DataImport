package org.avarice.dataimport.domain;


public interface OfficeBo {
	void save(Office office);
	Office getOfficeById(long id);
	Office getOfficeByName(String name);
}
