package com.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.solr.client.solrj.SolrServer;
import org.apache.solr.client.solrj.impl.HttpSolrServer;
import org.apache.solr.common.SolrInputDocument;

public class GavelCalanderUpdate {
	private static String SOLR_URL = "http://172.19.4.4:8085/solr/";
	private static final String CI_APPLICATIONS = SOLR_URL
			+ "gavel-customer-calendar";

	public static void main(String[] args) {
		String fileName = "D:\\GAVS\\2017.xlsx";
		String clientName = "DEFAULT";
		String year = "2017";
		System.out.println("fileName  " + fileName);
		SolrServer server = new HttpSolrServer(CI_APPLICATIONS);
		try {
			File file = new File(fileName);
			FileInputStream fIP = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(fIP);
			int count = 1;
			XSSFSheet sheet = workbook.getSheetAt(0);
			for (Row row : sheet) {
				SolrInputDocument monthObjectdcoument = new SolrInputDocument();
				SolrInputDocument weekObjectdcoument = new SolrInputDocument();
				SolrInputDocument dayObjectdcoument = new SolrInputDocument();
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 3) {
						if (row.getCell(cn) != null) {
							// Month
							if (cn == 3) {
								monthObjectdcoument.addField("CustomerID",
										clientName);
								monthObjectdcoument.addField("CalendarType",
										"DEFAULT");
								monthObjectdcoument.addField("CalendarYear",
										year);
								monthObjectdcoument.addField("TermsType",
										"MONTH");
								monthObjectdcoument.addField("TermKey", String
										.valueOf(row.getCell(cn)
												.getStringCellValue()));
								// monthObjectdcoument.addField("TermLabel",
								// row.getCell(cn).getStringCellValue() +
								// "-2016");
								monthObjectdcoument.addField("TermLabel",
										"month"
												+ (int) row.getCell(cn - 1)
														.getNumericCellValue());
								System.out.println("TermLabel-->"
										+ "month"
										+ (int) row.getCell(cn - 1)
												.getNumericCellValue());
								monthObjectdcoument.addField(
										"CustomerID-Type-Year-Key", clientName
												+ "-DEFAULT-"
												+ year
												+ "-MONTH-"
												+ row.getCell(cn)
														.getStringCellValue());
							}
							if (cn == 10) {
								Date date = HSSFDateUtil.getJavaDate(row
										.getCell(cn).getNumericCellValue());
								monthObjectdcoument.addField("StartDate",
										new SimpleDateFormat(
												"yyyy-MM-dd'T'00:00:00'Z'")
												.format(date));
							}
							if (cn == 11) {
								Date date = HSSFDateUtil.getJavaDate(row
										.getCell(cn).getNumericCellValue());
								monthObjectdcoument.addField("EndDate",
										new SimpleDateFormat(
												"yyyy-MM-dd'T'23:59:59'Z'")
												.format(date));
								if (monthObjectdcoument
										.getFieldValue("CustomerID-Type-Year-Key") != null) {
									server.add(monthObjectdcoument);
									monthObjectdcoument.clear();
									monthObjectdcoument = new SolrInputDocument();
								}
							}
							// Week
							if (cn == 7) {
								weekObjectdcoument.addField("CustomerID",
										clientName);
								weekObjectdcoument.addField("CalendarType",
										"DEFAULT");
								weekObjectdcoument.addField("CalendarYear",
										year);
								weekObjectdcoument
										.addField("TermsType", "WEEK");
								weekObjectdcoument.addField(
										"TermKey",
										"W"
												+ ((Double) row.getCell(cn)
														.getNumericCellValue())
														.intValue());
								weekObjectdcoument.addField(
										"TermLabel",
										"W"
												+ ((Double) row.getCell(cn)
														.getNumericCellValue())
														.intValue());
								weekObjectdcoument.addField(
										"CustomerID-Type-Year-Key",
										clientName
												+ "-DEFAULT-"
												+ year
												+ "-WEEK-"
												+ ((Double) row.getCell(cn)
														.getNumericCellValue())
														.intValue());
							}
							if (cn == 8) {
								Date date = HSSFDateUtil.getJavaDate(row
										.getCell(cn).getNumericCellValue());
								weekObjectdcoument.addField("StartDate",
										new SimpleDateFormat(
												"yyyy-MM-dd'T'00:00:00'Z'")
												.format(date));
							}
							if (cn == 9) {
								Date date = HSSFDateUtil.getJavaDate(row
										.getCell(cn).getNumericCellValue());
								weekObjectdcoument.addField("EndDate",
										new SimpleDateFormat(
												"yyyy-MM-dd'T'23:59:59'Z'")
												.format(date));
								if (weekObjectdcoument
										.getFieldValue("CustomerID-Type-Year-Key") != null) {
									server.add(weekObjectdcoument);
									weekObjectdcoument.clear();
									weekObjectdcoument = new SolrInputDocument();
								}
							}
							// Day
							if (cn == 1) {
								dayObjectdcoument.addField("CustomerID",
										clientName);
								dayObjectdcoument.addField("CalendarType",
										"DEFAULT");
								dayObjectdcoument
										.addField("CalendarYear", year);
								dayObjectdcoument.addField("TermsType", "DAY");
								dayObjectdcoument.addField("TermKey", "DAY"
										+ (count - 4));
								dayObjectdcoument.addField(
										"CustomerID-Type-Year-Key", clientName
												+ "-DEFAULT-" + year + "-DAY"
												+ (count - 4));
								if (dayObjectdcoument
										.getFieldValue("CustomerID-Type-Year-Key") != null) {
									server.add(dayObjectdcoument);
									dayObjectdcoument.clear();
									dayObjectdcoument = new SolrInputDocument();
								}
							}
							if (cn == 0) {
								Date date = HSSFDateUtil.getJavaDate(row
										.getCell(cn).getNumericCellValue());
								dayObjectdcoument.addField("StartDate",
										new SimpleDateFormat(
												"yyyy-MM-dd'T'00:00:00'Z'")
												.format(date));
								dayObjectdcoument.addField("EndDate",
										new SimpleDateFormat(
												"yyyy-MM-dd'T'23:59:59'Z'")
												.format(date));
								dayObjectdcoument.addField(
										"TermLabel",
										new SimpleDateFormat("dd/MMM").format(
												date).toUpperCase());
							}
						}
					}
				}
				count++;
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
