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

public class MonitoringMemory {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-monitoring-memory";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"MonitoringMemory.xlsx";

		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 0;
		XSSFSheet sheet = workbook.getSheetAt(0);
		//String alertdate = getAlertDate();

		for (Row row : sheet) {
			String NetworkName = "";
			String Description = "";
			int DeviceID = 0;
			int Size = 0;
			int Index = 0;
			int UsedMax = 0;
			String BestState = "";
			int NetworkInterfaceID = 0;
			String Type = "";
			String NetworkAddress = "";
			double TimeDelta = 0.0;
			String WorstState = "";
			int StatisticalDiskID = 0;
			int UsedAvg = 0;
			int UsedMin = 0;
			String PollTime = "";
			int StatisticalDiskIdentificationID = 0;
			String DeviceName = "";

			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 0 && row.getRowNum()<=8) {

						if (row.getCell(cn) != null) {
							if (cn == 0) {
								NetworkName = row.getCell(cn).getStringCellValue();
								master.addField("NetworkName", NetworkName);
							}
							if (cn == 1) {
								Description = row.getCell(cn).getStringCellValue();
								master.addField("Description", Description);
							}
							if (cn == 2) {
								DeviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", DeviceID);
							}
							if (cn == 3) {
								Size = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Size", Size);
							}
							if (cn == 4) {
								Index = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Index", Index);
							}
							if (cn == 5) {
								UsedMax = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMax", UsedMax);
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
								Type = row.getCell(cn).getStringCellValue();
								master.addField("Type", Type);
							}
							if (cn == 9) {
								NetworkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", NetworkAddress);
							}
							if (cn == 10) {
								StatisticalDiskIdentificationID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryIdentificationID", StatisticalDiskIdentificationID);
							}
							if (cn == 11) {
								TimeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", TimeDelta);
							}
							if (cn == 12) {
								StatisticalDiskID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryID", StatisticalDiskID);
							}
							if (cn == 13) {
								WorstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", WorstState);
							}
							if (cn == 14) {
								UsedAvg = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedAvg", UsedAvg);
							}
							if (cn == 15) {
								UsedMin = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMin", UsedMin);
							}
							if (cn == 16) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									PollTime = getOneDate() + valueAsInExcel;
									master.addField("PollTime", getDate(PollTime));
								}
							}
							if (cn == 17) {
								DeviceName = row.getCell(cn).getStringCellValue();
								master.addField("DeviceName", DeviceName);
							}

						}

					}
					
					if (row.getRowNum() > 8 && row.getRowNum()<=32) {

						if (row.getCell(cn) != null) {
							Date date = new Date();
							if (cn == 0) {
								NetworkName = row.getCell(cn).getStringCellValue();
								master.addField("NetworkName", NetworkName);
							}
							if (cn == 1) {
								Description = row.getCell(cn).getStringCellValue();
								master.addField("Description", Description);
							}
							if (cn == 2) {
								DeviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", DeviceID);
							}
							if (cn == 3) {
								Size = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Size", Size);
							}
							if (cn == 4) {
								Index = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Index", Index);
							}
							if (cn == 5) {
								UsedMax = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMax", UsedMax);
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
								Type = row.getCell(cn).getStringCellValue();
								master.addField("Type", Type);
							}
							if (cn == 9) {
								NetworkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", NetworkAddress);
							}
							if (cn == 10) {
								StatisticalDiskIdentificationID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryIdentificationID", StatisticalDiskIdentificationID);
							}
							if (cn == 11) {
								TimeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", TimeDelta);
							}
							if (cn == 12) {
								StatisticalDiskID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryID", StatisticalDiskID);
							}
							if (cn == 13) {
								WorstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", WorstState);
							}
							if (cn == 14) {
								UsedAvg = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedAvg", UsedAvg);
							}
							if (cn == 15) {
								UsedMin = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMin", UsedMin);
							}
							if (cn == 16) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									PollTime = getTwoDate() + valueAsInExcel;
									master.addField("PollTime", getDate(PollTime));
								}
							}
							if (cn == 17) {
								DeviceName = row.getCell(cn).getStringCellValue();
								master.addField("DeviceName", DeviceName);
							}

						}

					}
					
					if (row.getRowNum() > 32 && row.getRowNum() <=56) {

						if (row.getCell(cn) != null) {
							Date date = new Date();
							if (cn == 0) {
								NetworkName = row.getCell(cn).getStringCellValue();
								master.addField("NetworkName", NetworkName);
							}
							if (cn == 1) {
								Description = row.getCell(cn).getStringCellValue();
								master.addField("Description", Description);
							}
							if (cn == 2) {
								DeviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", DeviceID);
							}
							if (cn == 3) {
								Size = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Size", Size);
							}
							if (cn == 4) {
								Index = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Index", Index);
							}
							if (cn == 5) {
								UsedMax = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMax", UsedMax);
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
								Type = row.getCell(cn).getStringCellValue();
								master.addField("Type", Type);
							}
							if (cn == 9) {
								NetworkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", NetworkAddress);
							}
							if (cn == 10) {
								StatisticalDiskIdentificationID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryIdentificationID", StatisticalDiskIdentificationID);
							}
							if (cn == 11) {
								TimeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", TimeDelta);
							}
							if (cn == 12) {
								StatisticalDiskID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryID", StatisticalDiskID);
							}
							if (cn == 13) {
								WorstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", WorstState);
							}
							if (cn == 14) {
								UsedAvg = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedAvg", UsedAvg);
							}
							if (cn == 15) {
								UsedMin = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMin", UsedMin);
							}
							if (cn == 16) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									PollTime = getThreeDate() + valueAsInExcel;
									master.addField("PollTime", getDate(PollTime));
								}
							}
							if (cn == 17) {
								DeviceName = row.getCell(cn).getStringCellValue();
								master.addField("DeviceName", DeviceName);
							}

						}

					}
					if (row.getRowNum() > 56 && row.getRowNum() <=80) {

						if (row.getCell(cn) != null) {
							Date date = new Date();
							if (cn == 0) {
								NetworkName = row.getCell(cn).getStringCellValue();
								master.addField("NetworkName", NetworkName);
							}
							if (cn == 1) {
								Description = row.getCell(cn).getStringCellValue();
								master.addField("Description", Description);
							}
							if (cn == 2) {
								DeviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", DeviceID);
							}
							if (cn == 3) {
								Size = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Size", Size);
							}
							if (cn == 4) {
								Index = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Index", Index);
							}
							if (cn == 5) {
								UsedMax = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMax", UsedMax);
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
								Type = row.getCell(cn).getStringCellValue();
								master.addField("Type", Type);
							}
							if (cn == 9) {
								NetworkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", NetworkAddress);
							}
							if (cn == 10) {
								StatisticalDiskIdentificationID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryIdentificationID", StatisticalDiskIdentificationID);
							}
							if (cn == 11) {
								TimeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", TimeDelta);
							}
							if (cn == 12) {
								StatisticalDiskID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryID", StatisticalDiskID);
							}
							if (cn == 13) {
								WorstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", WorstState);
							}
							if (cn == 14) {
								UsedAvg = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedAvg", UsedAvg);
							}
							if (cn == 15) {
								UsedMin = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMin", UsedMin);
							}
							if (cn == 16) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									PollTime = getFoureDate() + valueAsInExcel;
									master.addField("PollTime", getDate(PollTime));
								}
							}
							if (cn == 17) {
								DeviceName = row.getCell(cn).getStringCellValue();
								master.addField("DeviceName", DeviceName);
							}

						}

					}
					if (row.getRowNum() > 80 && row.getRowNum() <=104) {

						if (row.getCell(cn) != null) {
							Date date = new Date();
							if (cn == 0) {
								NetworkName = row.getCell(cn).getStringCellValue();
								master.addField("NetworkName", NetworkName);
							}
							if (cn == 1) {
								Description = row.getCell(cn).getStringCellValue();
								master.addField("Description", Description);
							}
							if (cn == 2) {
								DeviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", DeviceID);
							}
							if (cn == 3) {
								Size = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Size", Size);
							}
							if (cn == 4) {
								Index = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Index", Index);
							}
							if (cn == 5) {
								UsedMax = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMax", UsedMax);
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
								Type = row.getCell(cn).getStringCellValue();
								master.addField("Type", Type);
							}
							if (cn == 9) {
								NetworkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", NetworkAddress);
							}
							if (cn == 10) {
								StatisticalDiskIdentificationID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryIdentificationID", StatisticalDiskIdentificationID);
							}
							if (cn == 11) {
								TimeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", TimeDelta);
							}
							if (cn == 12) {
								StatisticalDiskID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryID", StatisticalDiskID);
							}
							if (cn == 13) {
								WorstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", WorstState);
							}
							if (cn == 14) {
								UsedAvg = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedAvg", UsedAvg);
							}
							if (cn == 15) {
								UsedMin = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMin", UsedMin);
							}
							if (cn == 16) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									PollTime = getFiveDate() + valueAsInExcel;
									master.addField("PollTime", getDate(PollTime));
								}
							}
							if (cn == 17) {
								DeviceName = row.getCell(cn).getStringCellValue();
								master.addField("DeviceName", DeviceName);
							}

						}

					}
					if (row.getRowNum() > 104 && row.getRowNum() <=128) {

						if (row.getCell(cn) != null) {
							Date date = new Date();
							if (cn == 0) {
								NetworkName = row.getCell(cn).getStringCellValue();
								master.addField("NetworkName", NetworkName);
							}
							if (cn == 1) {
								Description = row.getCell(cn).getStringCellValue();
								master.addField("Description", Description);
							}
							if (cn == 2) {
								DeviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", DeviceID);
							}
							if (cn == 3) {
								Size = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Size", Size);
							}
							if (cn == 4) {
								Index = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Index", Index);
							}
							if (cn == 5) {
								UsedMax = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMax", UsedMax);
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
								Type = row.getCell(cn).getStringCellValue();
								master.addField("Type", Type);
							}
							if (cn == 9) {
								NetworkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", NetworkAddress);
							}
							if (cn == 10) {
								StatisticalDiskIdentificationID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryIdentificationID", StatisticalDiskIdentificationID);
							}
							if (cn == 11) {
								TimeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", TimeDelta);
							}
							if (cn == 12) {
								StatisticalDiskID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryID", StatisticalDiskID);
							}
							if (cn == 13) {
								WorstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", WorstState);
							}
							if (cn == 14) {
								UsedAvg = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedAvg", UsedAvg);
							}
							if (cn == 15) {
								UsedMin = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMin", UsedMin);
							}
							if (cn == 16) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									PollTime = getSixDate() + valueAsInExcel;
									master.addField("PollTime", getDate(PollTime));
								}
							}
							if (cn == 17) {
								DeviceName = row.getCell(cn).getStringCellValue();
								master.addField("DeviceName", DeviceName);
							}

						}

					}
					if (row.getRowNum() > 128) {

						if (row.getCell(cn) != null) {
							Date date = new Date();
							if (cn == 0) {
								NetworkName = row.getCell(cn).getStringCellValue();
								master.addField("NetworkName", NetworkName);
							}
							if (cn == 1) {
								Description = row.getCell(cn).getStringCellValue();
								master.addField("Description", Description);
							}
							if (cn == 2) {
								DeviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", DeviceID);
							}
							if (cn == 3) {
								Size = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Size", Size);
							}
							if (cn == 4) {
								Index = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Index", Index);
							}
							if (cn == 5) {
								UsedMax = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMax", UsedMax);
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
								Type = row.getCell(cn).getStringCellValue();
								master.addField("Type", Type);
							}
							if (cn == 9) {
								NetworkAddress = row.getCell(cn).getStringCellValue();
								master.addField("NetworkAddress", NetworkAddress);
							}
							if (cn == 10) {
								StatisticalDiskIdentificationID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryIdentificationID", StatisticalDiskIdentificationID);
							}
							if (cn == 11) {
								TimeDelta = row.getCell(cn).getNumericCellValue();
								master.addField("TimeDelta", TimeDelta);
							}
							if (cn == 12) {
								StatisticalDiskID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("StatisticalMemoryID", StatisticalDiskID);
							}
							if (cn == 13) {
								WorstState = row.getCell(cn).getStringCellValue();
								master.addField("WorstState", WorstState);
							}
							if (cn == 14) {
								UsedAvg = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedAvg", UsedAvg);
							}
							if (cn == 15) {
								UsedMin = (int) row.getCell(cn).getNumericCellValue();
								master.addField("UsedMin", UsedMin);
							}
							if (cn == 16) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									PollTime = getSevenDate() + valueAsInExcel;
									master.addField("PollTime", getDate(PollTime));
								}
							}
							if (cn == 17) {
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
		System.out.println("MonitoringMemory data loaded successfully");
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
	
	public static String getFoureDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -3);
		return dateFormat.format(cal.getTime());
	}
	public static String getFiveDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -4);
		return dateFormat.format(cal.getTime());
	}
	public static String getSixDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -5);
		return dateFormat.format(cal.getTime());
	}
	public static String getSevenDate() {
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
