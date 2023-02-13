package com.hc.excel.utils;

import com.hc.excel.domain.Address;
import com.hc.excel.domain.Customer;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;

public class ExcelExportUtil {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    List<Customer> customerList;

    public ExcelExportUtil(List<Customer> customerList) {
        this.customerList = customerList;
        workbook=new XSSFWorkbook();
    }

    public void createCell(Row row, int columnCounter,Object value, CellStyle style){
        sheet.autoSizeColumn(columnCounter);
        Cell cell= row.createCell(columnCounter);
        if(value instanceof Integer){
            cell.setCellValue((Integer) value);
        }if(value instanceof Double){
            cell.setCellValue((Double) value);
        }if(value instanceof Boolean){
            cell.setCellValue((Boolean) value);
        }if(value instanceof Long){
            cell.setCellValue((Long) value);
        }else{
            cell.setCellValue((String) value.toString());
        }
        cell.setCellStyle(style);
    }

    public void createHeaderRow(){
        sheet=workbook.createSheet("Customer Information");
        Row row= sheet.createRow(0);
        CellStyle style=workbook.createCellStyle();
        XSSFFont font=workbook.createFont();
        font.setBold(true);
        font.setFontHeight(24);
        font.setItalic(true);
        font.setColor(IndexedColors.BRIGHT_GREEN.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());


        createCell(row,0,"Customer Information",style);
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,6));
        font.setFontHeight((short) 10);

        row= sheet.createRow(1);
        font.setBold(true);
        font.setFontHeight(16);
        style.setFont(font);
        createCell(row, 0,"ID",style);
        createCell(row, 1,"First Name",style);
        createCell(row, 2,"Last Name",style);
        createCell(row, 3,"Contact Number",style);
        createCell(row, 4,"Country",style);
        createCell(row, 5,"City",style);
        createCell(row, 6,"Pin Code",style);
    }

    public void writeToExcel(){
        int rowCounter=2;
        CellStyle style= workbook.createCellStyle();
        XSSFFont font=workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());

        for(Customer customer:customerList){
            Row roe=sheet.createRow(rowCounter++);
            int columnCounter=0;
            createCell(roe,columnCounter++,customer.getId(),style);
            createCell(roe,columnCounter++,customer.getFirstName(),style);
            createCell(roe,columnCounter++,customer.getLastName(),style);
            //Data enterd in excel is text format not in number format
            createCell(roe,columnCounter++,(Integer) customer.getConNumber(),style);
            createCell(roe,columnCounter++,customer.getAddress().getCountry(),style);
            createCell(roe,columnCounter++,customer.getAddress().getCity(),style);
            createCell(roe,columnCounter++,customer.getAddress().getPinCode(),style);
        }
    }

    public void exportToExcel(HttpServletResponse response) throws IOException {
        createHeaderRow();
        writeToExcel();
        ServletOutputStream outputStream=response.getOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    public static boolean isValidateExcelFile(MultipartFile file){
        System.out.println(file.getContentType());
        return file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    }

    public static List<Customer> getCustomerFromExcel(InputStream inputStream){
        List<Customer> list=new ArrayList<>();
        try{
            XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream);
            XSSFSheet sheet=workbook1.getSheet("Customer Information");

            int rowIndex=0;
            for(Row row:sheet){
                if(rowIndex==0 || rowIndex==1){
                    rowIndex++;
                    continue;
                }
                Iterator<Cell> cellIterator = row.iterator();
                int cellIndex = 0;
                Customer customer = new Customer();
                Address address=new Address();
                while (cellIterator.hasNext()){
                    Cell cell1 = cellIterator.next();
                    switch (cellIndex){
                        case 1 : customer.setFirstName(cell1.getStringCellValue());
                            System.out.println(cell1.getStringCellValue());
                            break;
                        case 2 : customer.setLastName(cell1.getStringCellValue());
                            System.out.println(cell1.getStringCellValue());
                            break;
                        case 3 : customer.getConNumber( (Integer) ((Double) cell1.getNumericCellValue()).intValue());
                            System.out.println((Integer) ((Double) cell1.getNumericCellValue()).intValue());
                            break;
                        case 4 : address.setCountry(cell1.getStringCellValue());
                            System.out.println(cell1.getStringCellValue());
                            break;
                        case 5 : address.setCity(cell1.getStringCellValue());
                            System.out.println(cell1.getStringCellValue());
                            break;
                        case 6 : address.setPinCode(cell1.getStringCellValue());
                            System.out.println(cell1.getStringCellValue());
                            break;
                        default :
                            break;
                    }
                    cellIndex++;
                }
                customer.setAddress(address);
                list.add(customer);
            }

        }catch(Exception e){
            e.getStackTrace();
        }
        return list;
    }
}
