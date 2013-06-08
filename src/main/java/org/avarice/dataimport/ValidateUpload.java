package org.avarice.dataimport;

import org.springframework.web.multipart.MultipartException;
import org.springframework.web.multipart.MultipartFile;

public class ValidateUpload {
	
	
   public static void validateOfficeData(MultipartFile file){
   	 if(!file.getContentType().equals("application/vnd.ms-excel"))
   		 throw new MultipartException("Only excel files accepted!");
    }
}
