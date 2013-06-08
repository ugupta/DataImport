package org.avarice.dataimport.domain;

import java.util.List;

import org.hibernate.Query;
import org.hibernate.Session;
import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;

@Repository("officeDao")
public class OfficeDaoImpl implements OfficeDao{
	
	private SessionFactory sessionFactory;
	@Autowired
    public OfficeDaoImpl(SessionFactory sessionFactory)
    {
        this.sessionFactory=sessionFactory;
    }
	
	public Session currentSession(){
		return sessionFactory.getCurrentSession();
	}
	
	public void save(Office office){
		currentSession().save(office);
		currentSession().flush();
	}

	
	public Office getOfficeById(long id){
		return (Office)currentSession().get(Office.class, id);
	}
	
	public Office getOfficeByName(String name){
		Query query=currentSession().createQuery("from Office where name = '"+name+"'");
		List<Office> list = query.list();
		return (Office)list.get(0);
	}
}
