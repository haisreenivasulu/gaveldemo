package com.gavs;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Random;

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
public class MonitoringAlerts {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-monitoring-alerts";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"MonitoringAlerts.xlsx";

		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 1;
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			int EscalationState = 0;
			String Comment = "";
			String NetworkName = "";
			double AspectValue = 0.0;
			String AlertName = "";
			int DeviceID = 0;
			String BestState = "";
			int NetworkInterfaceID = 0;
			String AlertTime = "";
			String AlertDescription = "";
			int Active = 0;
			String NetworkAddress = "";
			int ThresholdState = 0;
			String WorstState = "";
			int AlertCenterStateHistoryID = 0;
			int ProActiveAlertState = 0;
			String DeviceName = "";

			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 0) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								EscalationState = (int) row.getCell(cn).getNumericCellValue();
								master.addField("EscalationState", EscalationState);
							}
							if (cn == 1) {
								Comment = row.getCell(cn).getStringCellValue();
								if (!Comment.isEmpty() && Comment != null) {
									master.addField("Comment", Comment);
								}
							}
							if (cn == 2) {
								NetworkName = row.getCell(cn).getStringCellValue();
								master.addField("NetworkName", NetworkName);
							}
							if (cn == 3) {
								AspectValue = row.getCell(cn).getNumericCellValue();
								if (AspectValue != 0.0) {
									master.addField("AspectValue", AspectValue + "%");
								}
							}
							if (cn == 4) {
								AlertName = row.getCell(cn).getStringCellValue();
								master.addField("AlertName", AlertName);
							}
							if (cn == 5) {
								DeviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", DeviceID);
							}
							if (cn == 6) {
								BestState = row.getCell(cn).getStringCellValue();
								master.addField("BestState", BestState);
							}
							if (cn == 7) {
								NetworkInterfaceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("NetworkInterfaceID", NetworkInterfaceID);
							}
							if (cn == 8) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									AlertTime = getMonthDate(0) + valueAsInExcel;
									master.addField("AlertTime", getDate(AlertTime));
								}
							}
							if (cn == 9) {
								AlertDescription = row.getCell(cn).getStringCellValue();
								master.addField("AlertDescription", AlertDescription);
							}
							if (cn == 10) {
								Active = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Active", Active);
							}
							if (cn == 11) {
								NetworkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", NetworkAddress);
							}
							if (cn == 12) {
								ThresholdState = (int) row.getCell(cn).getNumericCellValue();
								master.addField("ThresholdState", ThresholdState);
							}
							if (cn == 13) {
								WorstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", WorstState);
							}
							if (cn == 14) {
								AlertCenterStateHistoryID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("AlertCenterStateHistoryID", AlertCenterStateHistoryID);
							}
							if (cn == 15) {
								ProActiveAlertState = (int) row.getCell(cn).getNumericCellValue();
								master.addField("ProActiveAlertState", ProActiveAlertState);
							}
							if (cn == 16) {
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
		System.out.println("MonitoringAlerts loaded successfully");
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

	public static String getAlertDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		return dateFormat.format(cal.getTime());
	}

	public static String getMonthDate(int num) {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, num - randInt());
		return dateFormat.format(cal.getTime());
	}

	public static int randInt() {
		int min = 0;
		int max = 30;
		Random rand = new Random();
		int randomNum = rand.nextInt((max - min) + 1) + min;
		return randomNum;
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
