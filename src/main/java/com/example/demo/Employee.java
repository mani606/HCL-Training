package com.example.demo;

import java.io.*;
import java.sql.*;
import java.util.*;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class Employee {

	public static void main(String[] args) {
		 String jdbcURL = "jdbc:mysql://localhost:3306/db";
	        String username = "root";
	        String password = "mysql";
	 
	        String excelFilePath = "E:\\emp.xlsx";
	 
	        int batchSize = 20;
	 
	        Connection connection = null;
	 
	        try {
	            long start = System.currentTimeMillis();
	             
	            FileInputStream inputStream = new FileInputStream(excelFilePath);
	 
	            Workbook workbook = new XSSFWorkbook(inputStream);
	 
	            Sheet firstSheet = (Sheet) workbook.getSheetAt(0);
	            Iterator<Row> rowIterator = firstSheet.iterator();
	 
	            connection = DriverManager.getConnection(jdbcURL, username, password);
	            connection.setAutoCommit(false);
	  
	            String sql = "INSERT INTO employee (ename, salary, designation,gender) VALUES (?, ?, ?,?)";
	            PreparedStatement statement = connection.prepareStatement(sql);    
	             
	            int count = 0;
	             
	            rowIterator.next(); // skip the header row
	             
	            while (rowIterator.hasNext()) {
	                Row nextRow = rowIterator.next();
	                Iterator<Cell> cellIterator = nextRow.cellIterator();
	 
	                while (cellIterator.hasNext()) {
	                    Cell nextCell = cellIterator.next();
	 
	                    int columnIndex = nextCell.getColumnIndex();
	 
	                    switch (columnIndex) {
	                    case 0:
	                        String name = nextCell.getStringCellValue();
	                        statement.setString(1, name);
	                        break;
	                    case 1:
	                    	int salary= (int) nextCell.getNumericCellValue();
	                        statement.setInt(2, salary);
	                        break;
	                    case 2:
	                    	String designation = nextCell.getStringCellValue();
	                        statement.setString(3, designation);
	                        break;
	                    case 3:
	                        String gender = nextCell.getStringCellValue();
	                        statement.setString(4, gender);
	                    }
	 
	                }
	                 
	                statement.addBatch();
	                 
	                if (count % batchSize == 0) {
	                    statement.executeBatch();
	                }              
	 
	            }
	 
	            workbook.close();
	             
	            // execute the remaining queries
	            statement.executeBatch();
	  
	            connection.commit();
	            connection.close();
	             
	            long end = System.currentTimeMillis();
	            System.out.printf("Import done in %d ms\n", (end - start));
	             
	        } catch (IOException ex1) {
	            System.out.println("Error reading file");
	            ex1.printStackTrace();
	        } catch (SQLException ex2) {
	            System.out.println("Database error");
	            ex2.printStackTrace();
	        }
	 
	    }

	}

//code java.net
