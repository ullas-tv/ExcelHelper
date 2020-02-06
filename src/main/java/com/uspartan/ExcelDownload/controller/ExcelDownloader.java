package com.uspartan.ExcelDownload.controller;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.uspartan.ExcelDownload.model.Contact;
import com.uspartan.ExcelDownload.model.Person;
import com.uspartan.ExcelDownload.model.User;
import com.uspartan.ExcelDownload.service.GenericExcelService;

public class ExcelDownloader {

  public static void main(String[] args) throws IOException {

    /*
     * Workbook workbook = new XSSFWorkbook();
     * 
     * Sheet sheet = workbook.createSheet("Persons"); // sheet.setColumnWidth(0,
     * 6000); // sheet.setColumnWidth(1, 4000); sheet.setDefaultColumnWidth(6000);
     * 
     * Row header = sheet.createRow(0);
     * 
     * CellStyle headerStyle = workbook.createCellStyle();
     * headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
     * headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
     * 
     * XSSFFont font = ((XSSFWorkbook) workbook).createFont();
     * font.setFontName("Arial"); font.setFontHeightInPoints((short) 16);
     * font.setBold(true); headerStyle.setFont(font);
     * 
     * String[] columns = { "Name", "Age", "Phone", "Email" }; // Create cells for
     * (int i = 0; i < columns.length; i++) { Cell cell = header.createCell(i);
     * cell.setCellValue(columns[i]); cell.setCellStyle(headerStyle); }
     * 
     * // Style for Date Cell CellStyle dateCellStyle = workbook.createCellStyle();
     * org.apache.poi.ss.usermodel.CreationHelper createHelper =
     * workbook.getCreationHelper();
     * dateCellStyle.setDataFormat(createHelper.createDataFormat().
     * getFormat("dd/mm/yyyy h:mm"));
     * 
     * // General Cell Style CellStyle cellStyle = workbook.createCellStyle();
     * cellStyle.setWrapText(true);
     * 
     * File currDir = new File("."); String path = currDir.getAbsolutePath(); String
     * fileLocation = path.substring(0, path.length() - 1) +
     * "ExcelDownloadDemo.xlsx";
     * 
     * FileOutputStream outputStream = new FileOutputStream(fileLocation);
     * workbook.write(outputStream); workbook.close();
     */
    List<Person> persons = new ArrayList<>();
    Person p1 = new Person("A", "a@roytuts.com", "Kolkata");
    Person p2 = new Person("B", "b@roytuts.com", "Mumbai");
    Person p3 = new Person("C", "c@roytuts.com", "Delhi");
    Person p4 = new Person("D", "d@roytuts.com", "Chennai");
    Person p5 = new Person("E", "e@roytuts.com", "Bangalore");
    Person p6 = new Person("F", "f@roytuts.com", "Hyderabad");
    persons.add(p1);
    persons.add(p2);
    persons.add(p3);
    persons.add(p4);
    persons.add(p5);
    persons.add(p6);
    List<User> users = new ArrayList<>();
    ;
    User u1 = new User("u1", "pwd1");
    User u2 = new User("u2", "pwd2");
    User u3 = new User("u3", "pwd3");
    User u4 = new User("u4", "pwd4");
    User u5 = new User("u5", "pwd5");
    users.add(u1);
    users.add(u2);
    users.add(u3);
    users.add(u4);
    users.add(u5);
    List<Contact> contacts = new ArrayList<>();
    Contact c1 = new Contact("9478512354", "24157853", "24578613");
    Contact c2 = new Contact("9478512354", "24157853", "24578613");
    Contact c3 = new Contact("9478512354", "24157853", "24578613");
    Contact c4 = new Contact("9478512354", "24157853", "24578613");
    contacts.add(c1);
    contacts.add(c2);
    contacts.add(c3);
    contacts.add(c4);
    
    File currDir = new File(".");
    String path = currDir.getAbsolutePath();
    String fileLocation = path.substring(0, path.length() - 1) + "ExcelDownloadDemo.xlsx";
    
    GenericExcelService.writeToExcelInMultiSheets(fileLocation, "Person Details", persons);
    GenericExcelService.writeToExcelInMultiSheets(fileLocation, "User Details", users);
    GenericExcelService.writeToExcelInMultiSheets(fileLocation, "Contact Details", contacts);
  }

}
