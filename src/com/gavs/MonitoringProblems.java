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
public class MonitoringProblems {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-monitoring-problems";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"Monitoring_problems.xlsx";

		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 1;
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {

			String networkName = "";
			int down_time = 0;
			String down_on = getTodayDate();
			int deviceID = 0;
			int percentAvailable = 0;
			String bestState = "";
			int networkInterfaceID = 0;
			int packetsLost = 0;
			String cPUUtilization = "";

			int statisticalPingPacketLossID = 0;
			String networkAddress = "";
			Double timeDelta = 0.0;
			String worstState = "";
			String pollTime = getTodayDate();
			String deviceName = "";
			String diskSpaceUtilization = "";
			String memoryUtilization = "";

			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 0 && row.getRowNum()<=4) {

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
								down_time = (int) row.getCell(cn).getNumericCellValue();
								master.addField("down_time", down_time);

							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									down_on = down_on + valueAsInExcel;
									master.addField("down_on", getDate(down_on));
								}
							}
							if (cn == 4) {
								deviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", deviceID);

							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									down_on = down_on + valueAsInExcel;
									master.addField("up_on", getDate(down_on));
								}
							}
							if (cn == 6) {
								bestState = row.getCell(cn).getStringCellValue();
								master.addField("BestState", bestState);

							}
							if (cn == 7) {
								networkInterfaceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("NetworkInterfaceID", networkInterfaceID);

							}
							if (cn == 8) {
								percentAvailable = (int) row.getCell(cn).getNumericCellValue();
								master.addField("PercentAvailable", percentAvailable);
							}
							if (cn == 9) {
								packetsLost = (int) row.getCell(cn).getNumericCellValue();
								master.addField("PacketsLost", packetsLost);
							}
							if (cn == 10) {
								statisticalPingPacketLossID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalPingPacketLossID", statisticalPingPacketLossID);
							}

							if (cn == 11) {
								cPUUtilization = row.getCell(cn).getStringCellValue();
								master.addField("CPUUtilization", cPUUtilization);
							}
							if (cn == 12) {
								networkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", networkAddress);
							}
							if (cn == 13) {
								diskSpaceUtilization = row.getCell(cn).getStringCellValue();
								master.addField("DiskSpaceUtilization", diskSpaceUtilization);
							}

							if (cn == 14) {
								timeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", timeDelta);
							}

							if (cn == 15) {
								worstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", worstState);
							}
							if (cn == 16) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									down_on = down_on + valueAsInExcel;
									master.addField("PollTime", getDate(down_on));
								}
							}
							if (cn == 17) {
								memoryUtilization = row.getCell(cn).getStringCellValue();
								master.addField("MemoryUtilization", memoryUtilization);
							}
							if (cn == 18) {
								deviceName = row.getCell(cn).getStringCellValue();
								master.addField("DeviceName", deviceName);
							}
							

						}
					}
					if (row.getRowNum() > 4 && row.getRowNum()<=7 ) {

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
								down_time = (int) row.getCell(cn).getNumericCellValue();
								master.addField("down_time", down_time);

							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									down_on = getPreCrDate() + valueAsInExcel;
									master.addField("down_on", getDate(down_on));
								}
							}
							if (cn == 4) {
								deviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", deviceID);

							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									down_on = getPreCrDate() + valueAsInExcel;
									master.addField("up_on", getDate(down_on));
								}
							}
							if (cn == 6) {
								bestState = row.getCell(cn).getStringCellValue();
								master.addField("BestState", bestState);

							}
							if (cn == 7) {
								networkInterfaceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("NetworkInterfaceID", networkInterfaceID);

							}
							if (cn == 8) {
								percentAvailable = (int) row.getCell(cn).getNumericCellValue();
								master.addField("PercentAvailable", percentAvailable);
							}
							if (cn == 9) {
								packetsLost = (int) row.getCell(cn).getNumericCellValue();
								master.addField("PacketsLost", packetsLost);
							}
							if (cn == 10) {
								statisticalPingPacketLossID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalPingPacketLossID", statisticalPingPacketLossID);
							}

							if (cn == 11) {
								cPUUtilization = row.getCell(cn).getStringCellValue();
								master.addField("CPUUtilization", cPUUtilization);
							}
							if (cn == 12) {
								networkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", networkAddress);
							}
							if (cn == 13) {
								diskSpaceUtilization = row.getCell(cn).getStringCellValue();
								master.addField("DiskSpaceUtilization", diskSpaceUtilization);
							}

							if (cn == 14) {
								timeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", timeDelta);
							}

							if (cn == 15) {
								worstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", worstState);
							}
							if (cn == 16) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									down_on = getPreCrDate() + valueAsInExcel;
									master.addField("PollTime", getDate(down_on));
								}
							}
							if (cn == 17) {
								memoryUtilization = row.getCell(cn).getStringCellValue();
								master.addField("MemoryUtilization", memoryUtilization);
							}
							if (cn == 18) {
								deviceName = row.getCell(cn).getStringCellValue();
								master.addField("DeviceName", deviceName);
							}

						}
					}
					if (row.getRowNum() > 7 ) {

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
								down_time = (int) row.getCell(cn).getNumericCellValue();
								master.addField("down_time", down_time);

							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									down_on = getPre2CrDate() + valueAsInExcel;
									master.addField("down_on", getDate(down_on));
								}
							}
							if (cn == 4) {
								deviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", deviceID);

							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									down_on = getPre2CrDate() + valueAsInExcel;
									master.addField("up_on", getDate(down_on));
								}
							}
							if (cn == 6) {
								bestState = row.getCell(cn).getStringCellValue();
								master.addField("BestState", bestState);

							}
							if (cn == 7) {
								networkInterfaceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("NetworkInterfaceID", networkInterfaceID);

							}
							if (cn == 8) {
								percentAvailable = (int) row.getCell(cn).getNumericCellValue();
								master.addField("PercentAvailable", percentAvailable);
							}
							if (cn == 9) {
								packetsLost = (int) row.getCell(cn).getNumericCellValue();
								master.addField("PacketsLost", packetsLost);
							}
							if (cn == 10) {
								statisticalPingPacketLossID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalPingPacketLossID", statisticalPingPacketLossID);
							}

							if (cn == 11) {
								cPUUtilization = row.getCell(cn).getStringCellValue();
								master.addField("CPUUtilization", cPUUtilization);
							}
							if (cn == 12) {
								networkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", networkAddress);
							}
							if (cn == 13) {
								diskSpaceUtilization = row.getCell(cn).getStringCellValue();
								master.addField("DiskSpaceUtilization", diskSpaceUtilization);
							}

							if (cn == 14) {
								timeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", timeDelta);
							}

							if (cn == 15) {
								worstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", worstState);
							}
							if (cn == 16) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									down_on = getPre2CrDate() + valueAsInExcel;
									master.addField("PollTime", getDate(down_on));
								}
							}
							if (cn == 17) {
								memoryUtilization = row.getCell(cn).getStringCellValue();
								master.addField("MemoryUtilization", memoryUtilization);
							}
							if (cn == 18) {
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
		System.out.println("MonitoringProblems data loaded successfully");
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

	public static String getMonthDate(int num) {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, num - randInt());
		return dateFormat.format(cal.getTime());
	}

	public static String getTodayDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		return dateFormat.format(cal.getTime());
	}

	public static String getYesterDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -1);
		return dateFormat.format(cal.getTime());
	}

	public static String getPreCrDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -4);
		return dateFormat.format(cal.getTime());
	}
	
	public static String getPreClDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -2);
		return dateFormat.format(cal.getTime());
	}
	public static String getPre2CrDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -5);
		return dateFormat.format(cal.getTime());
	}
	
	public static String getPre2ClDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -3);
		return dateFormat.format(cal.getTime());
	}
	
	public static int randInt() {
		int min = 1;
		int max = 30;
		Random rand = new Random();
		int randomNum = rand.nextInt((max - min) + 1) + min;
		return randomNum;
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
