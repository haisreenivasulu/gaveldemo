package com.gavs;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.solr.client.solrj.SolrServer;
import org.apache.solr.client.solrj.SolrServerException;
import org.apache.solr.client.solrj.impl.HttpSolrServer;
import org.apache.solr.common.SolrInputDocument;

import com.gavs.constant.GavelConstant;

/**
 * @author sreenivasulu.s
 *
 */

public class CPUPrediction {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-prediction-cpu";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"CPUPrediction.xlsx";
		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 1;
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			String OnBehalfOf = "";
			int id = 0;
			String multiAssignedGroup = "";
			int ptime = 0;
			String pvalue = "";
			String deviceValue = "";
			int ldate = 0;

			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 0) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								multiAssignedGroup = row.getCell(cn).getStringCellValue();
								master.addField("Spike", multiAssignedGroup);
							}
							if (cn == 1) {
								id = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", id);
							}
							if (cn == 2) {
								multiAssignedGroup = row.getCell(cn).getStringCellValue();
								master.addField("AtRisk", multiAssignedGroup);
							}
							if (cn == 3) {
								pvalue = row.getCell(cn).getStringCellValue();
								master.addField("Alert", pvalue);
							}
							if (cn == 4) {
								deviceValue = row.getCell(cn).getStringCellValue();
								master.addField("NumberOfPredictions", deviceValue);
							}
							if (cn == 5) {
								ldate = (int) row.getCell(cn).getNumericCellValue();
								master.addField("LatestDateTime", ldate);
							}
						}

					}

				}
				//System.out.println("before : " + master);
				if (!master.isEmpty()) {
					serverMaster.add(master);
				}
				serverMaster.commit();
				//System.out.println("after: " + master);

				count++;
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}
		System.out.println("CPUPrediction loaded successfully");
	}

	public static String getDate(String createdDate) {

		try {

			Date createDate = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(createdDate);
			return new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'").format(createDate);
		} catch (java.text.ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}

	public static String getTomorrowDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 1);
		return dateFormat.format(cal.getTime());
	}

	public static String getTodayDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		return dateFormat.format(cal.getTime());
	}

	public static void deleteSolrData() {
		HttpSolrServer solr = new HttpSolrServer(master);
		try {
			solr.deleteByQuery("*:*");
		} catch (SolrServerException e) {
			throw new RuntimeException("Failed to delete data in Solr. " + e.getMessage(), e);
		} catch (IOException e) {
			throw new RuntimeException("Failed to delete data in Solr. " + e.getMessage(), e);
		}
	}

}
