package org.avarice.dataimport.domain;

import org.hibernate.Session;

public interface OfficeDao {
	
	Session currentSession();
	void save(Office office);
	Office getOfficeById(long id);
	Office getOfficeByName(String name);

}
