package com.hc.excel.domain;

import jakarta.persistence.*;
import lombok.*;

@Entity
@Data
@Setter
@Getter
@NoArgsConstructor
public class Customer {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Integer id;
    private String firstName;
    private String lastName;
    private Integer conNumber;
    @Embedded
    private Address address;

    public Customer(String firstName, String lastName, Integer conNumber, Address address) {
        this.firstName = firstName;
        this.lastName = lastName;
        this.conNumber = conNumber;
        this.address = address;
    }

    public Integer getConNumber(Integer intValue) {
        return this.conNumber=intValue;
    }
}
