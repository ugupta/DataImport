package org.avarice.dataimport;

import java.io.IOException;
import java.util.Date;

import javax.annotation.Resource;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.avarice.dataimport.domain.Office;
import org.avarice.dataimport.domain.OfficeBo;
import org.hibernate.SessionFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartException;
import org.springframework.web.multipart.MultipartFile;

@Service("saveContents")
@Transactional
public class SaveContents {
	private static final Logger logger = LoggerFactory.getLogger(SaveContents.class);
	
	 @Resource(name="sessionFactory")
	 private SessionFactory sessionFactory;
	 
	 @Resource(name="officeBo")
	 private OfficeBo officeBo;
	
	public void saveOfficeContents(MultipartFile file)throws IOException{
		try{
			HSSFWorkbook offices= new HSSFWorkbook(file.getInputStream());
			HSSFSheet worksheet = offices.getSheet("Offices");
			HSSFRow entry;
			Integer noOfEntries=1;
			//getLastRowNum and getPhysicalNumberOfRows showing false values sometimes.
			while(worksheet.getRow(noOfEntries)!=null){
				noOfEntries++;
			}
			logger.info(noOfEntries.toString());
			for(int rowIndex=1;rowIndex<noOfEntries;rowIndex++){
			    entry=worksheet.getRow(rowIndex);
			    Integer externalId=((Double)entry.getCell(0).getNumericCellValue()).intValue();
			    Office parent=officeBo.getOfficeByName(entry.getCell(1).getStringCellValue());
			    Long parentId=parent.getId();
			    String name=entry.getCell(2).getStringCellValue();
			    Date openingDate=entry.getCell(3).getDateCellValue();
//			    Date date = new SimpleDateFormat("MM/dd/yyyy", Locale.ENGLISH).parse(openingDate);
			    logger.info("Row Contents:"+parentId+" "+externalId+" "+name+" "+openingDate);
			    Office office=new Office();
			    office.setParentId((long)parentId);
			    office.setExternalId((long)externalId);
			    office.setName(name);
			    office.setOpeningDate(openingDate);
			    String parentHierarchy=parent.getHierarchy();
			    //Pre save to generate id for use in hierarchy
			    officeBo.save(office);
			    office.setHierarchy(parentHierarchy+office.getId()+".");
			    officeBo.save(office);
			}
			
			
		}catch(Exception e){
		    logger.info(e.getMessage()+" "+e.getCause());
			throw new MultipartException("Constraints Violated");
		}
	}
	
	
}
