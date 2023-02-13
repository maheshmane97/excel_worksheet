package com.hc.excel.service;

import com.hc.excel.domain.Customer;
import com.hc.excel.repository.CustomerRepository;
import com.hc.excel.utils.ExcelExportUtil;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
@Service
public class CustomerService {
    @Autowired
    private CustomerRepository customerRepository;

    public List<Customer> ExportToExcel(HttpServletResponse response) throws IOException {
        List<Customer> customerList=customerRepository.findAll();
        ExcelExportUtil util=new ExcelExportUtil(customerList);
        util.exportToExcel(response);
        return customerList;
    }

    public void saveCustomerFromExcelToDb(MultipartFile file){
        System.out.println(ExcelExportUtil.isValidateExcelFile(file));
        if(ExcelExportUtil.isValidateExcelFile(file)){
            List<Customer>  list=null;
            try {
                 list = ExcelExportUtil.getCustomerFromExcel(file.getInputStream());
            } catch (IOException e) {
                throw new IllegalArgumentException("This file is not Excel File..!");
            }
            this.customerRepository.saveAll(list);
        }
    }

    public List<Customer> getAllCustomers(){
        return customerRepository.findAll();
    }
}
