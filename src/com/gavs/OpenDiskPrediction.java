package com.gavs;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

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
public class OpenDiskPrediction {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-prediction-disk";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"PredictionOpenDisk.xlsx";

		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 0;
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			String OnBehalfOf = "";
			String id = "";
			int multiAssignedGroup = 0;
			String pDate = "";
			int ptime = 0;
			String pvalue = "";
			String deviceValue = "";
			double pValue = 0.0;
			int did = 0;
			double pred = 0.0;

			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 0) {
						if (row.getCell(cn) != null) {
							Date date = new Date();
							if (cn == 0) {
								multiAssignedGroup = (int) row.getCell(cn).getNumericCellValue();
								master.addField("PredictionDay", multiAssignedGroup);
							}
							if (cn == 1) {
								id = row.getCell(cn).getStringCellValue();
								master.addField("DriveName", id);

							}
							if (cn == 2) {
								String newDate = forecstDays(date, count);
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									pDate = newDate + valueAsInExcel;
									master.addField("PredictionDate", getDate(pDate));
								}
							}
							if (cn == 3) {
								did = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", did);

							}
							if (cn == 4) {
								pred = (double) row.getCell(cn).getNumericCellValue();
								master.addField("PredictionValue", pred);
							}
							if (cn == 5) {
								master.addField("DeviceID-DriveName-PredictionDay",
										master.getFieldValue("DeviceID") + "||" + master.getFieldValue("DriveName")
												+ "||" + master.getFieldValue("PredictionDay"));
							}

						}

					}
				}
				if (!master.isEmpty()) {
					serverMaster.add(master);
				}
				serverMaster.commit();
				//System.out.println("master: "+master);

				count++;
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}
		System.out.println("DiskPrediction data loaded successfully");
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
	public static String historyDays(Date date, int days) {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		GregorianCalendar cal = new GregorianCalendar();
		cal.setTime(date);
		cal.add(Calendar.DATE, -days);
				
		return dateFormat.format(cal.getTime());
	}
	
	public static String forecstDays(Date date, int days) {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		GregorianCalendar cal = new GregorianCalendar();
		cal.setTime(date);
		cal.add(Calendar.DATE, days);
				
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
