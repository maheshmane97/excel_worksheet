package com.hc.excel.controller;

import com.hc.excel.service.CustomerService;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.support.ServletUriComponentsBuilder;

import java.io.IOException;
import java.util.Map;

@RestController
@RequestMapping("/customer")
public class CustomerController {

    @Autowired
    private CustomerService service;

    @GetMapping("/excel")
    public void exportToExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheetl");
        String headerKey="Content-Disposition";
        String headerValue="attachment; filename=Customer_Information.xlsx";
        response.setHeader(headerKey,headerValue);
        service.ExportToExcel(response);
    }
    @PostMapping("/save")
    public ResponseEntity<?> UploadCustomer(@RequestParam("file") MultipartFile file){
        this.service.saveCustomerFromExcelToDb(file);
        String downloadUri= ServletUriComponentsBuilder.fromCurrentContextPath().path("/excel/")
                            .toUriString();
        System.out.println(downloadUri);
        return ResponseEntity.ok(Map.of("Customers Data Uploaded and saved successfully..!", downloadUri));
    }
}
