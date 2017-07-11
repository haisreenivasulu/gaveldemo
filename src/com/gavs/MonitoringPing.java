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
public class MonitoringPing {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-monitoring-ping";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"MonitoringPing.xlsx";

		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 1;
		XSSFSheet sheet = workbook.getSheetAt(0);
		String alertdate = getAlertDate();

		for (Row row : sheet) {

			String networkName = "";
			int deviceID = 0;
			int percentAvailable = 0;
			String bestState = "";
			int networkInterfaceID = 0;
			int packetsLost = 0;

			int statisticalPingPacketLossID = 0;
			String networkAddress = "";
			Double timeDelta = 0.0;
			String worstState = "";
			String deviceName = "";
			String pollTime = "";

			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 0) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								int packetsSent = 0;
								packetsSent = (int) row.getCell(cn).getNumericCellValue();
								master.addField("PacketsSent", packetsSent);

							}
							if (cn == 1) {
								networkName = row.getCell(cn).getStringCellValue();
								master.addField("NetworkName", networkName);

							}
							if (cn == 2) {
								deviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", deviceID);

							}
							if (cn == 3) {
								bestState = row.getCell(cn).getStringCellValue();
								master.addField("BestState", bestState);

							}
							if (cn == 4) {
								networkInterfaceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("NetworkInterfaceID", networkInterfaceID);

							}
							if (cn == 5) {
								percentAvailable = (int) row.getCell(cn).getNumericCellValue();
								master.addField("PercentAvailable", percentAvailable);
							}
							if (cn == 6) {
								packetsLost = (int) row.getCell(cn).getNumericCellValue();
								master.addField("PacketsLost", packetsLost);
							}
							if (cn == 7) {
								statisticalPingPacketLossID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalPingPacketLossID", statisticalPingPacketLossID);
							}

							if (cn == 8) {
								networkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", networkAddress);
							}

							if (cn == 9) {
								timeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", timeDelta);
							}

							if (cn == 10) {
								worstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", worstState);
							}
							if (cn == 11) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									pollTime = alertdate + valueAsInExcel;
									master.addField("PollTime", getDate(pollTime));
								}
							}

							if (cn == 12) {
								deviceName = row.getCell(cn).getStringCellValue();
								master.addField("DeviceName", deviceName);
							}

						}
					}

				}
				if (!master.isEmpty()) {
					serverMaster.add(master);
				}
				serverMaster.commit();
				// System.out.println("after: " + master);

				count++;
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}
		System.out.println("Ping data loaded successfully");
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
