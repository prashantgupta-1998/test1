package com.sampleproject.sample_project.Service;

import com.sampleproject.sample_project.Entity.Employee;
import com.sampleproject.sample_project.Interface.EmployeeInterface;
import com.sampleproject.sample_project.Repository.EmployeeRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;

@Service
public class EmployeeService implements EmployeeInterface {

    @Autowired
    private EmployeeRepository employeeRepository;

    @Override
    public ByteArrayInputStream downloadAllUserData() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (Workbook workbook = new XSSFWorkbook()) {
            List<Employee> employeeList = employeeRepository.findAll();
            populateAllUserData(Optional.ofNullable(employeeList), workbook);
            workbook.write(out);
            workbook.close();
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException io) {
            io.printStackTrace();
            return null;
        }
    }

    public void populateAllUserData(Optional<List<Employee>> optionalEmployees, Workbook workbook) {
        if (!optionalEmployees.isPresent()) {
            System.out.println("No User Data Present on DB");
        }
        Sheet sheet = workbook.createSheet("USERS");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        Row row = sheet.createRow(0);
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        headerCellStyle.setFont(font);

        // Creating header cells
        String[] headers = {"Name", "Department", "Salary"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerCellStyle);
        }

        int index = 1;

        // Creating data row for each of objects
        for (Employee employee : optionalEmployees.get()) {
            Row dataRow = sheet.createRow(index); // index to exclude the header row
            dataRow.createCell(0).setCellValue(employee.getName());
            dataRow.createCell(1).setCellValue(employee.getDept());
            dataRow.createCell(2).setCellValue(employee.getSalary().doubleValue());
            index = index + 1;
        }

        // Making sure the size of excel cell auto resize to fit the data
        for (int colSize = 0; colSize < headers.length; colSize++) {
            sheet.autoSizeColumn(colSize);
        }
    }

}
