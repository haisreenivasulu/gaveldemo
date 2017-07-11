package com.gavs;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.DataFormatter;
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
public class BMSMonitoringCPU {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-monitoring-cpu";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
 		SolrServer serverMaster = new HttpSolrServer(master);
		String fileName = GavelConstant.FILE_PATH+"BSMCPU.xlsx";

		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 0;
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			int processorLoadAvg = 0;
			int processorLoadMax = 0;
			String NetworkName = "";
			String Description = "";
			int DeviceID = 0;
			int Index = 0;
			String BestState = "";
			int NetworkInterfaceID = 0;
			String NetworkAddress = "";
			double TimeDelta = 0.0;
			String WorstState = "";
			int StatisticalCpuID = 0;
			String PollTime = "";
			int StatisticalDiskIdentificationID = 0;
			String DeviceName = "";
			int processorLoadMin = 0;

			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 0 ) {

						if (row.getCell(cn) != null) {
							if (cn == 0) {
								processorLoadMax = (int) row.getCell(cn).getNumericCellValue();
								master.addField("ProcessorLoadMax", processorLoadMax);
							}
							if (cn == 1) {
								NetworkName = row.getCell(cn).getStringCellValue();
								master.addField("NetworkName", NetworkName);
							}
							if (cn == 2) {
								Description = row.getCell(cn).getStringCellValue();
								master.addField("Description", Description);
							}
							if (cn == 3) {
								DeviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", DeviceID);
							}
							if (cn == 4) {
								Index = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Index", Index);
							}
							if (cn == 5) {
								BestState = row.getCell(cn).getStringCellValue();
								master.addField("BestState", BestState);
							}
							if (cn == 6) {
								NetworkInterfaceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("NetworkInterfaceID", NetworkInterfaceID);
							}

							if (cn == 7) {
								StatisticalDiskIdentificationID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalCpuIdentificationID", StatisticalDiskIdentificationID);
							}
							if (cn == 8) {
								processorLoadAvg = (int) row.getCell(cn).getNumericCellValue();
								master.addField("ProcessorLoadAvg", processorLoadAvg);
							}
							if (cn == 9) {
								NetworkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", NetworkAddress);
							}
							if (cn == 10) {
								TimeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", TimeDelta);
							}
							if (cn == 11) {
								processorLoadMin = (int) row.getCell(cn).getNumericCellValue();
								master.addField("ProcessorLoadMin", processorLoadMin);
							}
							if (cn == 12) {
								WorstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", WorstState);
							}
							if (cn == 13) {
								StatisticalCpuID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalCpuID", StatisticalCpuID);
							}

							if (cn == 14) {
								//String newDate = historyDays(date, count);
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									PollTime = getOneDate() + valueAsInExcel;
									master.addField("PollTime", getDate(PollTime));
								}
							}
							if (cn == 15) {
								DeviceName = row.getCell(cn).getStringCellValue();
								master.addField("DeviceName", DeviceName);
							}

						}

					}
					
				}
				if (!master.isEmpty()) {
					serverMaster.add(master);
				}
				serverMaster.commit();
				 //System.out.println("master : " + master);

				count++;
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}
		System.out.println("BSM CPU data loaded successfully");
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
	
	public static String getOneDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		return dateFormat.format(cal.getTime());
	}
	
	public static String getTwoDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -1);
		return dateFormat.format(cal.getTime());
	}
	
	public static String getThreeDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -2);
		return dateFormat.format(cal.getTime());
	}
	public static String getFourDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -3);
		return dateFormat.format(cal.getTime());
	}
	
	public static String getfiveDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -4);
		return dateFormat.format(cal.getTime());
	}
	
	public static String getsixDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -5);
		return dateFormat.format(cal.getTime());
	}
	public static String getsevenDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -6);
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
