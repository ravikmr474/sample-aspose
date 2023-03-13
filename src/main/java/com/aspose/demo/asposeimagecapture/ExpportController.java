package com.aspose.demo.asposeimagecapture;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ExpportController {

	@Autowired
	private ExportService exportService;
	
	@GetMapping("/download")
	public void download(HttpServletResponse response) {
		try {
			exportService.download(response);
			System.out.println("Done");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
