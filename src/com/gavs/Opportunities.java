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
public class Opportunities {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-prediction-opportunities";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"Opportunities.xlsx";

		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 1;
		String createddate = getcreatedDate();
		String expecteddate = getexpectedDate();

		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			String ticketId = "";
			int tvmPredictMinValue = 0;
			int tvmPredictMaxValue = 0;
			String category = "";
			String ticketTitle = "";
			String keywordVar = "";
			String deviceDNSName = "";
			String opportunityDescription = "";
			String currentState = "";
			int deviceID = 0;
			String issueType = "";
			String deviceIPAddress = "";
			int opportunityID = 0;
			int urgency = 0;
			String deviceName = "";
			String deviceType = "";
			String lName = "";
			String coments = "";
			String fName = "";

			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {

					if (row.getRowNum() > 0 && row.getRowNum() <= 5) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								issueType = row.getCell(cn).getStringCellValue();
								master.addField("IssueType", issueType);
							}
							if (cn == 1) {
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")
										|| issueType.equalsIgnoreCase("DISK")) {
									String valueAsInExcel = row.getCell(cn).getStringCellValue();
									if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
										master.addField("DeviceDescription", valueAsInExcel);
									}
								}
							}
							if (cn == 2) {
								category = row.getCell(cn).getStringCellValue();
								if (!category.isEmpty() && category != null) {
									master.addField("Category", category);
								}
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = createddate + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("CreatedTime", getDate(cdate));
								}
							}
							if (cn == 4) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = createddate + valueAsInExcel;
								String openexdate = openExDate() + valueAsInExcel;
								String nextDate = getNextDate() + valueAsInExcel;
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")) {
									if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
										master.addField("ExpectedOccurrenceDate", getDate(cdate));
									}
								} else if (issueType.equalsIgnoreCase("DISK")) {
									master.addField("ExpectedOccurrenceDate", getDate(openexdate));
								} else if (issueType.equalsIgnoreCase("TicketVolume")) {
									master.addField("ExpectedOccurrenceDate", getDate(nextDate));
								}
							}

							if (cn == 5) {
								deviceDNSName = row.getCell(cn).getStringCellValue();
								if (!deviceDNSName.isEmpty() && deviceDNSName != null) {
									master.addField("DeviceDNSName", deviceDNSName);
								}
							}
							if (cn == 6) {
								currentState = row.getCell(cn).getStringCellValue();
								master.addField("CurrentState", currentState);
							}
							if (cn == 7) {
								opportunityDescription = row.getCell(cn).getStringCellValue();
								master.addField("OpportunityDescription", opportunityDescription);
							}

							if (cn == 8) {
								deviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", deviceID);
							}

							if (cn == 9) {
								deviceIPAddress = row.getCell(cn).getStringCellValue();
								master.addField("DeviceIPAddress", deviceIPAddress);
							}
							if (cn == 10) {
								deviceType = row.getCell(cn).getStringCellValue();
								if (!deviceType.isEmpty() && deviceType != null) {
									master.addField("DeviceType", deviceType);
								}
							}
							if (cn == 11) {
								opportunityID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("OpportunityID", opportunityID);
							}

							if (cn == 12) {
								urgency = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Urgency", urgency);
							}
							if (cn == 13) {
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")
										|| issueType.equalsIgnoreCase("DISK")) {
									deviceName = row.getCell(cn).getStringCellValue();
									master.addField("DeviceName", deviceName);
								}
							}
							if (cn == 14) {
								tvmPredictMinValue = (int) row.getCell(cn).getNumericCellValue();
								if (tvmPredictMinValue == 1 || tvmPredictMinValue == 2) {
									master.addField("TvmPredictMinValue", tvmPredictMinValue);
								}
							}
							if (cn == 15) {
								if (issueType.equalsIgnoreCase("TicketVolume")) {
									ticketId = row.getCell(cn).getStringCellValue();
									if (!ticketId.isEmpty() && ticketId != null) {
										master.addField("TicketId", ticketId);
									}
								}
							}

							if (cn == 16) {
								tvmPredictMaxValue = (int) row.getCell(cn).getNumericCellValue();
								if (issueType.equalsIgnoreCase("TicketVolume")) {
									if (tvmPredictMaxValue == 3 || tvmPredictMaxValue == 6) {
										master.addField("TvmPredictMaxValue", tvmPredictMaxValue);
									}
								}
							}
							if (cn == 17) {
								ticketTitle = row.getCell(cn).getStringCellValue();
								if (!ticketTitle.isEmpty() && ticketTitle != null) {
									master.addField("TicketTitle", ticketTitle);
								}
							}
							if (cn == 18) {
								keywordVar = row.getCell(cn).getStringCellValue();
								if (!keywordVar.isEmpty() && keywordVar != null) {
									master.addField("KeywordVar", keywordVar);
								}
							}
							if (cn == 19) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = getresolvedDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("ResolvedTime", getDate(cdate));
								}
							}
							if (cn == 20) {
								coments = row.getCell(cn).getStringCellValue();
								if (!coments.isEmpty() && coments != null) {
									master.addField("ResolvedComments", coments);
								}
							}
							if (cn == 21) {
								fName = row.getCell(cn).getStringCellValue();
								if (!fName.isEmpty() && fName != null) {
									master.addField("ResolvedByFirstName", fName);
								}
							}
							if (cn == 22) {
								lName = row.getCell(cn).getStringCellValue();
								if (!lName.isEmpty() && lName != null) {
									master.addField("ResolvedByLastName", lName);
								}
							}
							if (cn == 23) {
								String resolved = row.getCell(cn).getStringCellValue();
								if (!resolved.isEmpty() && resolved != null) {
									master.addField("ResolvedByEmail", resolved);
								}
							}
							if (cn == 24) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("IgnoredComments", ignoredComments);
								}
							}
							if (cn == 25) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getcreatedDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("IgnoredTime", getDate(idate));
								}

							}
							if (cn == 26) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("DeclinedByEmail", ignoredComments);
								}
							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getResCrDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("DeclinedTime", getDate(idate));
								}

							}

						}

					}

					if (row.getRowNum() > 5 && row.getRowNum() <= 6) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								issueType = row.getCell(cn).getStringCellValue();
								master.addField("IssueType", issueType);
							}
							if (cn == 1) {
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")
										|| issueType.equalsIgnoreCase("DISK")) {
									String valueAsInExcel = row.getCell(cn).getStringCellValue();
									if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
										master.addField("DeviceDescription", valueAsInExcel);
									}
								}
							}
							if (cn == 2) {
								category = row.getCell(cn).getStringCellValue();
								if (!category.isEmpty() && category != null) {
									master.addField("Category", category);
								}
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = getResCrDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("CreatedTime", getDate(cdate));
								}
							}
							if (cn == 4) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								 String cdate = getResCrDate() + valueAsInExcel;
								String openExDate = closedExDate() + valueAsInExcel;
								String nextDate = getNextDate() + valueAsInExcel;
								 if (issueType.equalsIgnoreCase("MEMORY") ||
								  issueType.equalsIgnoreCase("CPU")) { if
								 (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
								  master.addField("ExpectedOccurrenceDate",
								  getDate(cdate)); } } else
								  if (issueType.equalsIgnoreCase("DISK")) {
									master.addField("ExpectedOccurrenceDate", getDate(openExDate));
								} else if (issueType.equalsIgnoreCase("TicketVolume")) {
									master.addField("ExpectedOccurrenceDate", getDate(nextDate));
								}
							}

							if (cn == 5) {
								deviceDNSName = row.getCell(cn).getStringCellValue();
								if (!deviceDNSName.isEmpty() && deviceDNSName != null) {
									master.addField("DeviceDNSName", deviceDNSName);
								}
							}
							if (cn == 6) {
								currentState = row.getCell(cn).getStringCellValue();
								master.addField("CurrentState", currentState);
							}
							if (cn == 7) {
								opportunityDescription = row.getCell(cn).getStringCellValue();
								master.addField("OpportunityDescription", opportunityDescription);
							}

							if (cn == 8) {
								deviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", deviceID);
							}

							if (cn == 9) {
								deviceIPAddress = row.getCell(cn).getStringCellValue();
								master.addField("DeviceIPAddress", deviceIPAddress);
							}
							if (cn == 10) {
								deviceType = row.getCell(cn).getStringCellValue();
								if (!deviceType.isEmpty() && deviceType != null) {
									master.addField("DeviceType", deviceType);
								}
							}
							if (cn == 11) {
								opportunityID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("OpportunityID", opportunityID);
							}

							if (cn == 12) {
								urgency = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Urgency", urgency);
							}
							if (cn == 13) {
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")
										|| issueType.equalsIgnoreCase("DISK")) {
									deviceName = row.getCell(cn).getStringCellValue();
									master.addField("DeviceName", deviceName);
								}
							}
							if (cn == 14) {
								tvmPredictMinValue = (int) row.getCell(cn).getNumericCellValue();
								if (tvmPredictMinValue == 1 || tvmPredictMinValue == 2) {
									master.addField("TvmPredictMinValue", tvmPredictMinValue);
								}
							}
							if (cn == 15) {
								if (issueType.equalsIgnoreCase("TicketVolume")) {
									ticketId = row.getCell(cn).getStringCellValue();
									if (!ticketId.isEmpty() && ticketId != null) {
										master.addField("TicketId", ticketId);
									}
								}
							}

							if (cn == 16) {
								tvmPredictMaxValue = (int) row.getCell(cn).getNumericCellValue();
								if (issueType.equalsIgnoreCase("TicketVolume")) {
									if (tvmPredictMaxValue == 3 || tvmPredictMaxValue == 6) {
										master.addField("TvmPredictMaxValue", tvmPredictMaxValue);
									}
								}
							}
							if (cn == 17) {
								ticketTitle = row.getCell(cn).getStringCellValue();
								if (!ticketTitle.isEmpty() && ticketTitle != null) {
									master.addField("TicketTitle", ticketTitle);
								}
							}
							if (cn == 18) {
								keywordVar = row.getCell(cn).getStringCellValue();
								if (!keywordVar.isEmpty() && keywordVar != null) {
									master.addField("KeywordVar", keywordVar);
								}
							}
							if (cn == 19) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = getresDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("ResolvedTime", getDate(cdate));
								}
							}
							if (cn == 20) {
								coments = row.getCell(cn).getStringCellValue();
								if (!coments.isEmpty() && coments != null) {
									master.addField("ResolvedComments", coments);
								}
							}
							if (cn == 21) {
								fName = row.getCell(cn).getStringCellValue();
								if (!fName.isEmpty() && fName != null) {
									master.addField("ResolvedByFirstName", fName);
								}
							}
							if (cn == 22) {
								lName = row.getCell(cn).getStringCellValue();
								if (!lName.isEmpty() && lName != null) {
									master.addField("ResolvedByLastName", lName);
								}
							}
							if (cn == 23) {
								String resolved = row.getCell(cn).getStringCellValue();
								if (!resolved.isEmpty() && resolved != null) {
									master.addField("ResolvedByEmail", resolved);
								}
							}
							if (cn == 24) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("IgnoredComments", ignoredComments);
								}
							}
							if (cn == 25) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getcreatedDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("IgnoredTime", getDate(idate));
								}

							}
							if (cn == 26) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("DeclinedByEmail", ignoredComments);
								}
							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getResCrDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("DeclinedTime", getDate(idate));
								}

							}

						}

					}
					if (row.getRowNum() > 6 && row.getRowNum() <= 8) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								issueType = row.getCell(cn).getStringCellValue();
								master.addField("IssueType", issueType);
							}
							if (cn == 1) {
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")
										|| issueType.equalsIgnoreCase("DISK")) {
									String valueAsInExcel = row.getCell(cn).getStringCellValue();
									if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
										master.addField("DeviceDescription", valueAsInExcel);
									}
								}
							}
							if (cn == 2) {
								category = row.getCell(cn).getStringCellValue();
								if (!category.isEmpty() && category != null) {
									master.addField("Category", category);
								}
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = getPreCrDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("CreatedTime", getDate(cdate));
								}
							}
							if (cn == 4) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = closedExDate() + valueAsInExcel;
								String openExDate = closedExDate() + valueAsInExcel;
								// String nextDate = getNextDate() +
								// valueAsInExcel;
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")) {
									if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
										master.addField("ExpectedOccurrenceDate", getDate(cdate));
									}
								} else if (issueType.equalsIgnoreCase("DISK")) {
									master.addField("ExpectedOccurrenceDate", getDate(openExDate));
								} /*
									 * else if (issueType.equalsIgnoreCase(
									 * "TicketVolume")) {
									 * master.addField("ExpectedOccurrenceDate",
									 * getDate(nextDate)); }
									 */
							}

							if (cn == 5) {
								deviceDNSName = row.getCell(cn).getStringCellValue();
								if (!deviceDNSName.isEmpty() && deviceDNSName != null) {
									master.addField("DeviceDNSName", deviceDNSName);
								}
							}
							if (cn == 6) {
								currentState = row.getCell(cn).getStringCellValue();
								master.addField("CurrentState", currentState);
							}
							if (cn == 7) {
								opportunityDescription = row.getCell(cn).getStringCellValue();
								master.addField("OpportunityDescription", opportunityDescription);
							}

							if (cn == 8) {
								deviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", deviceID);
							}

							if (cn == 9) {
								deviceIPAddress = row.getCell(cn).getStringCellValue();
								master.addField("DeviceIPAddress", deviceIPAddress);
							}
							if (cn == 10) {
								deviceType = row.getCell(cn).getStringCellValue();
								if (!deviceType.isEmpty() && deviceType != null) {
									master.addField("DeviceType", deviceType);
								}
							}
							if (cn == 11) {
								opportunityID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("OpportunityID", opportunityID);
							}

							if (cn == 12) {
								urgency = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Urgency", urgency);
							}
							if (cn == 13) {
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")
										|| issueType.equalsIgnoreCase("DISK")) {
									deviceName = row.getCell(cn).getStringCellValue();
									master.addField("DeviceName", deviceName);
								}
							}
							if (cn == 14) {
								tvmPredictMinValue = (int) row.getCell(cn).getNumericCellValue();
								if (tvmPredictMinValue == 1 || tvmPredictMinValue == 2) {
									master.addField("TvmPredictMinValue", tvmPredictMinValue);
								}
							}
							if (cn == 15) {
								if (issueType.equalsIgnoreCase("TicketVolume")) {
									ticketId = row.getCell(cn).getStringCellValue();
									if (!ticketId.isEmpty() && ticketId != null) {
										master.addField("TicketId", ticketId);
									}
								}
							}

							if (cn == 16) {
								tvmPredictMaxValue = (int) row.getCell(cn).getNumericCellValue();
								if (issueType.equalsIgnoreCase("TicketVolume")) {
									if (tvmPredictMaxValue == 3 || tvmPredictMaxValue == 6) {
										master.addField("TvmPredictMaxValue", tvmPredictMaxValue);
									}
								}
							}
							if (cn == 17) {
								ticketTitle = row.getCell(cn).getStringCellValue();
								if (!ticketTitle.isEmpty() && ticketTitle != null) {
									master.addField("TicketTitle", ticketTitle);
								}
							}
							if (cn == 18) {
								keywordVar = row.getCell(cn).getStringCellValue();
								if (!keywordVar.isEmpty() && keywordVar != null) {
									master.addField("KeywordVar", keywordVar);
								}
							}
							if (cn == 19) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = getPreClDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("ResolvedTime", getDate(cdate));
								}
							}
							if (cn == 20) {
								coments = row.getCell(cn).getStringCellValue();
								if (!coments.isEmpty() && coments != null) {
									master.addField("ResolvedComments", coments);
								}
							}
							if (cn == 21) {
								fName = row.getCell(cn).getStringCellValue();
								if (!fName.isEmpty() && fName != null) {
									master.addField("ResolvedByFirstName", fName);
								}
							}
							if (cn == 22) {
								lName = row.getCell(cn).getStringCellValue();
								if (!lName.isEmpty() && lName != null) {
									master.addField("ResolvedByLastName", lName);
								}
							}
							if (cn == 23) {
								String resolved = row.getCell(cn).getStringCellValue();
								if (!resolved.isEmpty() && resolved != null) {
									master.addField("ResolvedByEmail", resolved);
								}
							}
							if (cn == 24) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("IgnoredComments", ignoredComments);
								}
							}
							if (cn == 25) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getcreatedDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("IgnoredTime", getDate(idate));
								}

							}
							if (cn == 26) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("DeclinedByEmail", ignoredComments);
								}
							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getResCrDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("DeclinedTime", getDate(idate));
								}

							}

						}

					}

					if (row.getRowNum() > 8 && row.getRowNum() <= 11) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								issueType = row.getCell(cn).getStringCellValue();
								master.addField("IssueType", issueType);
							}
							if (cn == 1) {
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")
										|| issueType.equalsIgnoreCase("DISK")) {
									String valueAsInExcel = row.getCell(cn).getStringCellValue();
									if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
										master.addField("DeviceDescription", valueAsInExcel);
									}
								}
							}
							if (cn == 2) {
								category = row.getCell(cn).getStringCellValue();
								if (!category.isEmpty() && category != null) {
									master.addField("Category", category);
								}
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = getPre2CrDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("CreatedTime", getDate(cdate));
								}
							}
							if (cn == 4) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = closedExDate() + valueAsInExcel;
								String openExDate = closedExDate() + valueAsInExcel;
								// String nextDate = getNextDate() +
								// valueAsInExcel;
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")) {
									if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
										master.addField("ExpectedOccurrenceDate", getDate(cdate));
									}
								} else if (issueType.equalsIgnoreCase("DISK")) {
									master.addField("ExpectedOccurrenceDate", getDate(openExDate));
								} /*
									 * else if (issueType.equalsIgnoreCase(
									 * "TicketVolume")) {
									 * master.addField("ExpectedOccurrenceDate",
									 * getDate(nextDate)); }
									 */
							}

							if (cn == 5) {
								deviceDNSName = row.getCell(cn).getStringCellValue();
								if (!deviceDNSName.isEmpty() && deviceDNSName != null) {
									master.addField("DeviceDNSName", deviceDNSName);
								}
							}
							if (cn == 6) {
								currentState = row.getCell(cn).getStringCellValue();
								master.addField("CurrentState", currentState);
							}
							if (cn == 7) {
								opportunityDescription = row.getCell(cn).getStringCellValue();
								master.addField("OpportunityDescription", opportunityDescription);
							}

							if (cn == 8) {
								deviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", deviceID);
							}

							if (cn == 9) {
								deviceIPAddress = row.getCell(cn).getStringCellValue();
								master.addField("DeviceIPAddress", deviceIPAddress);
							}
							if (cn == 10) {
								deviceType = row.getCell(cn).getStringCellValue();
								if (!deviceType.isEmpty() && deviceType != null) {
									master.addField("DeviceType", deviceType);
								}
							}
							if (cn == 11) {
								opportunityID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("OpportunityID", opportunityID);
							}

							if (cn == 12) {
								urgency = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Urgency", urgency);
							}
							if (cn == 13) {
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")
										|| issueType.equalsIgnoreCase("DISK")) {
									deviceName = row.getCell(cn).getStringCellValue();
									master.addField("DeviceName", deviceName);
								}
							}
							if (cn == 14) {
								tvmPredictMinValue = (int) row.getCell(cn).getNumericCellValue();
								if (tvmPredictMinValue == 1 || tvmPredictMinValue == 2) {
									master.addField("TvmPredictMinValue", tvmPredictMinValue);
								}
							}
							if (cn == 15) {
								if (issueType.equalsIgnoreCase("TicketVolume")) {
									ticketId = row.getCell(cn).getStringCellValue();
									if (!ticketId.isEmpty() && ticketId != null) {
										master.addField("TicketId", ticketId);
									}
								}
							}

							if (cn == 16) {
								tvmPredictMaxValue = (int) row.getCell(cn).getNumericCellValue();
								if (issueType.equalsIgnoreCase("TicketVolume")) {
									if (tvmPredictMaxValue == 3 || tvmPredictMaxValue == 4) {
										master.addField("TvmPredictMaxValue", tvmPredictMaxValue);
									}
								}
							}
							if (cn == 17) {
								ticketTitle = row.getCell(cn).getStringCellValue();
								if (!ticketTitle.isEmpty() && ticketTitle != null) {
									master.addField("TicketTitle", ticketTitle);
								}
							}
							if (cn == 18) {
								keywordVar = row.getCell(cn).getStringCellValue();
								if (!keywordVar.isEmpty() && keywordVar != null) {
									master.addField("KeywordVar", keywordVar);
								}
							}
							if (cn == 19) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = getPre2ClDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("ResolvedTime", getDate(cdate));
								}
							}
							if (cn == 20) {
								coments = row.getCell(cn).getStringCellValue();
								if (!coments.isEmpty() && coments != null) {
									master.addField("ResolvedComments", coments);
								}
							}
							if (cn == 21) {
								fName = row.getCell(cn).getStringCellValue();
								if (!fName.isEmpty() && fName != null) {
									master.addField("ResolvedByFirstName", fName);
								}
							}
							if (cn == 22) {
								lName = row.getCell(cn).getStringCellValue();
								if (!lName.isEmpty() && lName != null) {
									master.addField("ResolvedByLastName", lName);
								}
							}
							if (cn == 23) {
								String resolved = row.getCell(cn).getStringCellValue();
								if (!resolved.isEmpty() && resolved != null) {
									master.addField("ResolvedByEmail", resolved);
								}
							}
							if (cn == 24) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("IgnoredComments", ignoredComments);
								}
							}
							if (cn == 25) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getPreCrDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("IgnoredTime", getDate(idate));
								}

							}
							if (cn == 26) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("DeclinedByEmail", ignoredComments);
								}
							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getResCrDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("DeclinedTime", getDate(idate));
								}

							}

						}

					}
					if (row.getRowNum() > 11) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								issueType = row.getCell(cn).getStringCellValue();
								master.addField("IssueType", issueType);
							}
							if (cn == 1) {
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")
										|| issueType.equalsIgnoreCase("DISK")) {
									String valueAsInExcel = row.getCell(cn).getStringCellValue();
									if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
										master.addField("DeviceDescription", valueAsInExcel);
									}
								}
							}
							if (cn == 2) {
								category = row.getCell(cn).getStringCellValue();
								if (!category.isEmpty() && category != null) {
									master.addField("Category", category);
								}
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = getcreatedDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("CreatedTime", getDate(cdate));
								}
							}
							if (cn == 4) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								// String cdate = getResCrDate() +
								// valueAsInExcel;
								String openExDate = closedExDate() + valueAsInExcel;
								String nextDate = getcreatedDate() + valueAsInExcel;
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")) {
									if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
										master.addField("ExpectedOccurrenceDate", getDate(nextDate));
									}
								} else if (issueType.equalsIgnoreCase("DISK")) {
									master.addField("ExpectedOccurrenceDate", getDate(openExDate));
								} else if (issueType.equalsIgnoreCase("TicketVolume")) {
									master.addField("ExpectedOccurrenceDate", getDate(nextDate));
								}
							}

							if (cn == 5) {
								deviceDNSName = row.getCell(cn).getStringCellValue();
								if (!deviceDNSName.isEmpty() && deviceDNSName != null) {
									master.addField("DeviceDNSName", deviceDNSName);
								}
							}
							if (cn == 6) {
								currentState = row.getCell(cn).getStringCellValue();
								master.addField("CurrentState", currentState);
							}
							if (cn == 7) {
								opportunityDescription = row.getCell(cn).getStringCellValue();
								master.addField("OpportunityDescription", opportunityDescription);
							}

							if (cn == 8) {
								deviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", deviceID);
							}

							if (cn == 9) {
								deviceIPAddress = row.getCell(cn).getStringCellValue();
								master.addField("DeviceIPAddress", deviceIPAddress);
							}
							if (cn == 10) {
								deviceType = row.getCell(cn).getStringCellValue();
								if (!deviceType.isEmpty() && deviceType != null) {
									master.addField("DeviceType", deviceType);
								}
							}
							if (cn == 11) {
								opportunityID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("OpportunityID", opportunityID);
							}

							if (cn == 12) {
								urgency = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Urgency", urgency);
							}
							if (cn == 13) {
								if (issueType.equalsIgnoreCase("MEMORY") || issueType.equalsIgnoreCase("CPU")
										|| issueType.equalsIgnoreCase("DISK")) {
									deviceName = row.getCell(cn).getStringCellValue();
									master.addField("DeviceName", deviceName);
								}
							}
							if (cn == 14) {
								tvmPredictMinValue = (int) row.getCell(cn).getNumericCellValue();
								if (tvmPredictMinValue == 1 || tvmPredictMinValue == 2) {
									master.addField("TvmPredictMinValue", tvmPredictMinValue);
								}
							}
							if (cn == 15) {
								if (issueType.equalsIgnoreCase("TicketVolume")) {
									ticketId = row.getCell(cn).getStringCellValue();
									if (!ticketId.isEmpty() && ticketId != null) {
										master.addField("TicketId", ticketId);
									}
								}
							}

							if (cn == 16) {
								tvmPredictMaxValue = (int) row.getCell(cn).getNumericCellValue();
								if (issueType.equalsIgnoreCase("TicketVolume")) {
									if (tvmPredictMaxValue == 3 || tvmPredictMaxValue == 6) {
										master.addField("TvmPredictMaxValue", tvmPredictMaxValue);
									}
								}
							}
							if (cn == 17) {
								ticketTitle = row.getCell(cn).getStringCellValue();
								if (!ticketTitle.isEmpty() && ticketTitle != null) {
									master.addField("TicketTitle", ticketTitle);
								}
							}
							if (cn == 18) {
								keywordVar = row.getCell(cn).getStringCellValue();
								if (!keywordVar.isEmpty() && keywordVar != null) {
									master.addField("KeywordVar", keywordVar);
								}
							}
							if (cn == 19) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String cdate = getcreatedDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("ResolvedTime", getDate(cdate));
								}
							}
							if (cn == 20) {
								coments = row.getCell(cn).getStringCellValue();
								if (!coments.isEmpty() && coments != null) {
									master.addField("ResolvedComments", coments);
								}
							}
							if (cn == 21) {
								fName = row.getCell(cn).getStringCellValue();
								if (!fName.isEmpty() && fName != null) {
									master.addField("ResolvedByFirstName", fName);
								}
							}
							if (cn == 22) {
								lName = row.getCell(cn).getStringCellValue();
								if (!lName.isEmpty() && lName != null) {
									master.addField("ResolvedByLastName", lName);
								}
							}
							if (cn == 23) {
								String resolved = row.getCell(cn).getStringCellValue();
								if (!resolved.isEmpty() && resolved != null) {
									master.addField("ResolvedByEmail", resolved);
								}
							}
							if (cn == 24) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("IgnoredComments", ignoredComments);
								}
							}
							if (cn == 25) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getcreatedDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("IgnoredTime", getDate(idate));
								}
							}
							if (cn == 26) {
								String inprogressComments = row.getCell(cn).getStringCellValue();
								if (!inprogressComments.isEmpty() && inprogressComments != null) {
									master.addField("InprogressByFirstName", inprogressComments);
								}
							}
							if (cn == 27) {
								String iEmail = row.getCell(cn).getStringCellValue();
								if (!iEmail.isEmpty() && iEmail != null) {
									master.addField("InprogressByLastName", iEmail);
								}
							}
							if (cn == 28) {
								String inprogressComments = row.getCell(cn).getStringCellValue();
								if (!inprogressComments.isEmpty() && inprogressComments != null) {
									master.addField("InprogressComments", inprogressComments);
								}
							}
							if (cn == 29) {
								String iEmail = row.getCell(cn).getStringCellValue();
								if (!iEmail.isEmpty() && iEmail != null) {
									master.addField("InprogressByEmail", iEmail);
								}
							}
							if (cn == 30) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getcreatedDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("InprogressTime", getDate(idate));
								}

							}
							if (cn == 31) {
								String inprogressComments = row.getCell(cn).getStringCellValue();
								if (!inprogressComments.isEmpty() && inprogressComments != null) {
									master.addField("DeclinedByFirstName", inprogressComments);
								}
							}
							if (cn == 32) {
								String iEmail = row.getCell(cn).getStringCellValue();
								if (!iEmail.isEmpty() && iEmail != null) {
									master.addField("DeclinedByLastName", iEmail);
								}
							}
							if (cn == 33) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("DeclinedComments", ignoredComments);
								}
							}
							if (cn == 34) {
								String ignoredComments = row.getCell(cn).getStringCellValue();
								if (!ignoredComments.isEmpty() && ignoredComments != null) {
									master.addField("DeclinedByEmail", ignoredComments);
								}
							}
							if (cn == 35) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String idate = getcreatedDate() + valueAsInExcel;
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									master.addField("DeclinedTime", getDate(idate));
								}
							}
							if (cn == 36) {
								int ignoredComments = (int) row.getCell(cn).getNumericCellValue();
								if (ignoredComments >=1) {
									master.addField("DeclinedWaitingHrs", ignoredComments);
								}
							}

						}

					}

				}
				if (!master.isEmpty()) {
					serverMaster.add(master);
				}
				serverMaster.commit();
				// System.out.println("master : "+master);

				count++;
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}
		System.out.println("Opportunities data loaded successfully");
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

	public static String getcreatedDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		return dateFormat.format(cal.getTime());
	}
	public static String getResCrDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -1);
		return dateFormat.format(cal.getTime());
	}
	public static String getResExDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		return dateFormat.format(cal.getTime());
	}
	public static String getresolvedDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 3);
		return dateFormat.format(cal.getTime());
	}
	public static String getexpectedDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 2);
		return dateFormat.format(cal.getTime());
	}
	public static String openExDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 37);
		return dateFormat.format(cal.getTime());
	}
	public static String closedExDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 8);
		return dateFormat.format(cal.getTime());
	}
	public static String getNextDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 1);
		return dateFormat.format(cal.getTime());
	}
	public static String getresDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
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
