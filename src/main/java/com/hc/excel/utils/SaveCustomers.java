package com.hc.excel.utils;

import com.hc.excel.domain.Address;
import com.hc.excel.domain.Customer;
import com.hc.excel.repository.CustomerRepository;
import jakarta.annotation.PostConstruct;
import org.springframework.stereotype.Component;

import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;

//@Component
public class SaveCustomers {

    private CustomerRepository customerRepository;

    public SaveCustomers(CustomerRepository customerRepository) {
        this.customerRepository = customerRepository;
    }

    private List<Customer> list= Arrays.asList(
            new Customer("Mahesh","Mane",10, new Address("India","Sangli","416 419")),
            new Customer("Manoj","Mane",20, new Address("India","KM","416 419")),
            new Customer("Rohit","More",30, new Address("India","hingangaon","416 413")),
            new Customer("Yogesh","Mane",40, new Address("India","Sangli","416 419")),
            new Customer("Mahesh","Mane",50, new Address("India","Sangli","416 419"))
    );
    @PostConstruct
    public void saveData(){
        customerRepository.saveAll(list);
    }
}
