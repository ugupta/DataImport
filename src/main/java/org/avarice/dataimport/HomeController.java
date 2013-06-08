package org.avarice.dataimport;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;

import org.avarice.dataimport.service.DownloadService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartException;
import org.springframework.web.multipart.MultipartFile;

/**
 * Handles requests for the application home page.
 */
@Controller
public class HomeController {
	
	private static final Logger logger = LoggerFactory.getLogger(HomeController.class);
	
	@Resource(name="downloadService")
	 private DownloadService downloadService;
	@Resource(name="saveContents")
	 private SaveContents saveContents;
	
	@RequestMapping(value = "/", method = RequestMethod.GET)
	public String home(Model model) {
		
		Map<String, String> entities = new LinkedHashMap<String, String>();
		entities.put("office", "Offices");
		entities.put("client", "Clients");
		entities.put("loan", "Loans");
		entities.put("loanProduct", "Loan Products");
		entities.put("staff", "Staff");
		
		model.addAttribute("entities", entities );
		
		return "home";
	}
	
	@RequestMapping(value = "/download", method = RequestMethod.GET)
	public void home(@RequestParam("entity") String viewType,HttpServletResponse response, Model model)throws ClassNotFoundException {
		logger.info("The client wants to view {}.", viewType);
		
		downloadService.downloadXLS(response,viewType);
	}
	
	@RequestMapping(value = "/form", method = RequestMethod.POST)
    public String handleFormUpload(@RequestParam("file") MultipartFile file)throws IOException {
		try{
        if (!file.isEmpty()) {
        	ValidateUpload.validateOfficeData(file);
        	saveContents.saveOfficeContents(file);
            logger.info("Upload successful!");
            }
		}catch(MultipartException e){
			logger.info("Upload failed!");
		}
		return "redirect:/";
    }
	
	
}
