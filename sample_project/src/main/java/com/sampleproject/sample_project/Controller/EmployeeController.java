package com.sampleproject.sample_project.Controller;

import com.sampleproject.sample_project.Service.EmployeeService;
import org.apache.poi.util.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.IOException;

@RestController
public class EmployeeController {

    @Autowired
    private EmployeeService employeeService;

    @RequestMapping(value = "/EmployeePage")
    public String EmployeePage() {
        return "EmployeePage";
    }

    @GetMapping("/fetchAllEmployeeData/user-data.xlsx")
    public void fetchAllEmployeeData(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=user-data.xlsx");
        ByteArrayInputStream inputStream = employeeService.downloadAllUserData();
        IOUtils.copy(inputStream, response.getOutputStream());
    }

}
