package com.nanda.file.apache.poi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Scanner;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileOpeartions {

	public static void writeIntoXLSX() throws FileNotFoundException, IOException {

		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet = workBook.createSheet("Customers");

		Map<String, String> customers = getInput();

		Iterator<Entry<String, String>> mapEntry = customers.entrySet().iterator();
		int i = 0;
		while (mapEntry.hasNext()) {
			Row row = sheet.createRow(i++);
			Cell name = row.createCell(0);
			Cell number = row.createCell(1);
			Entry<String, String> entry = mapEntry.next();
			name.setCellValue(entry.getKey());
			number.setCellValue(entry.getValue());
		}

		workBook.write(new FileOutputStream(new File("./src/test/resources/CustomerBook.xlsx")));

		workBook.close();

	}

	public static void writeIntoXLS() throws FileNotFoundException, IOException {
		
		HSSFWorkbook workBook = new HSSFWorkbook();
		
		HSSFSheet workSheet = workBook.createSheet("Customers");
		
		Map<String, String> customers = getInput();
		
		Iterator<Entry<String, String>> mapEntry = customers.entrySet().iterator();
		int i = 0;
		
		while (mapEntry.hasNext()) {
			Row row = workSheet.createRow(i++);
			Cell name = row.createCell(0);
			Cell number = row.createCell(1);
			Entry<String, String> entry = mapEntry.next();
			name.setCellValue(entry.getKey());
			number.setCellValue(entry.getValue());
		}

		workBook.write(new FileOutputStream(new File("./src/test/resources/CustomerBook.xls")));

		workBook.close();

	}

	public static Map<String, String> getInput() {

		Map<String, String> customers = new TreeMap<String, String>();

		Scanner scanInput = new Scanner(System.in);
		while (scanInput.hasNext()) {
			String customer = scanInput.next();
			customers.put(customer.split(":")[0], customer.split(":")[1]);
		}
		scanInput.close();

		return customers;
	}

	public static void main(String[] args) throws InvalidFormatException, IOException {
		
		writeIntoXLS();

	}

}
