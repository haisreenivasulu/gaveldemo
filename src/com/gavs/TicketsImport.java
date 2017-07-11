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
import org.jsoup.Jsoup;

import com.gavs.constant.GavelConstant;

/**
 * @author sreenivasulu.s
 *
 */
public class TicketsImport {
	
	private static final String CI_APPLICATIONS = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-tickets-manageengine";

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-tickets-master";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer server = new HttpSolrServer(CI_APPLICATIONS);
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"TicketsMasterHM.xlsx";

		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 1;
		XSSFSheet sheet = workbook.getSheetAt(0);
		String resolvedDate = getResolvedDate();
		String updatedDate = getUpdatedDate();
		String createdDate = getCreatedDate();

		for (Row row : sheet) {


			int number = 0;
			String createdOn = "";
			String Incidentstate = "";
			String Priority = "";
			String Resolved = "";
			String Assignedto = "";
			String Assignmentgroup = "";
			String ShortDescription = "";
			String Closed = "";
			String Category = "";
			String OnBehalfOf = "";
			String Updated = "";
			String updatedBy = "";
			String desc = "";
			String component = "";
			Float ETC_PREDICT_VAL = (float) 0.0;
			String notes = "";
			String closenotescomments = "";
			String productCategory = "";
			int surveyTaken = 0;
			int csatAwaitingActionETC = 0;
			int csatAwaitingActionSentiment = 0;
			int responseSlaBreach = 0;
			int resolutionSlaBreach = 0;
			String overallSentiment = "";
			String suggestedVirtualSupervisior = "";
			String emailId = "";
			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {

					if (row.getRowNum() > 0 && row.getRowNum() <= 25) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								Assignedto = row.getCell(cn).getStringCellValue();
								if (!Assignedto.isEmpty() && Assignedto != null) {
									master.addField("Assignee", Assignedto);

								}
							}
							if (cn == 1) {
								Assignmentgroup = row.getCell(cn).getStringCellValue();
								master.addField("AssignedGroup", Assignmentgroup);
							}
							if (cn == 2) {
								Category = row.getCell(cn).getStringCellValue();
								master.addField("Department", Category);
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Closed = resolvedDate + valueAsInExcel;
									master.addField("CompletedDate", getDate(Closed));
								}
							}
							if (cn == 4) {
								component = row.getCell(cn).getStringCellValue();
								master.addField("ReportedMode", component);
							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								createdOn = resolvedDate + valueAsInExcel;
								master.addField("CreatedDate", getDate(createdOn));
							}
							if (cn == 6) {

								ShortDescription = row.getCell(cn).getStringCellValue();
								String shortDescription = Jsoup.parse(ShortDescription.toString()).text();
								shortDescription = shortDescription.toString().replaceAll("(\r\n|\n)", " ");
								shortDescription = CleanInvalidXmlChars(shortDescription.toString(), "");
								if (!shortDescription.isEmpty() && shortDescription != null) {
								master.addField("Description", shortDescription);
								}
							}
							if (cn == 7) {
								Incidentstate = row.getCell(cn).getStringCellValue();
								master.addField("Status", Incidentstate);
							}
							if (cn == 8) {
								productCategory = row.getCell(cn).getStringCellValue();
								master.addField("ProductCategory", productCategory);
							}
							if (cn == 9) {
								number = (int) row.getCell(cn).getNumericCellValue();
								master.addField("IncidentID", number);
							}
							if (cn == 10) {
								OnBehalfOf = row.getCell(cn).getStringCellValue();
								if (!OnBehalfOf.isEmpty() && OnBehalfOf != null) {
									master.addField("FirstName", OnBehalfOf);

								}
							}
							if (cn == 11) {
								emailId = row.getCell(cn).getStringCellValue();
								master.addField("EmailId", emailId);
							}
							if (cn == 12) {
								Priority = row.getCell(cn).getStringCellValue();
								master.addField("Priority", Priority);
							}
							if (cn == 13) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Resolved = resolvedDate + valueAsInExcel;
									master.addField("ResolvedDate", getDate(Resolved));
								}
							}

							if (cn == 14) {
								desc = row.getCell(cn).getStringCellValue();
								String descr = Jsoup.parse(desc.toString()).text();
								descr = descr.toString().replaceAll("(\r\n|\n)", " ");
								descr = CleanInvalidXmlChars(descr.toString(), "");
								if (!descr.isEmpty() && descr != null) {
								master.addField("Title", descr);
								}
							}

							if (cn == 15) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = resolvedDate + valueAsInExcel;
									master.addField("LastModifiedDate", getDate(Updated));
								}
							}
							if (cn == 16) {

								notes = row.getCell(cn).getStringCellValue();
								if (!notes.isEmpty() && notes != null) {
																
								if (String.valueOf(notes).contains("Additional comments")) {

									String comment = String.valueOf(notes)
											.replace("(Additional comments (Customer Visible))", "~");
									String[] comments = String.valueOf(comment).split("~");
									String notesCreatedByDetails = comments[0];

									String[] NotesCreated = String.valueOf(notesCreatedByDetails).split(" -");
									String notesCreatedDate = NotesCreated[0];
									String notesCreatedBy = NotesCreated[1];

									String note = comments[1];

									note = note.toString().replaceAll("(\r\n|\n)", " ");
									note = CleanInvalidXmlChars(note.toString(), "");

									String actualNotesCreatedBy = notesCreatedByDetails.substring(22);

									if (!notes.isEmpty() && notes != null) {
										master.addField("Note", note);
									}
									if (!notesCreatedBy.isEmpty() && notesCreatedBy != null) {
										master.addField("NotesCreatedBy", actualNotesCreatedBy);
									}
									if (!notesCreatedDate.isEmpty() && notesCreatedDate != null) {
										master.addField("NotesCreatedDate", master.getFieldValue("LastModifiedDate"));
									}

								}
								}
							}
							if (cn == 17) {
								closenotescomments = row.getCell(cn).getStringCellValue();
								String Closenotescomments = Jsoup.parse(closenotescomments.toString()).text();
								Closenotescomments = Closenotescomments.toString().replaceAll("(\r\n|\n)", " ");
								Closenotescomments = CleanInvalidXmlChars(Closenotescomments.toString(), "");
								master.addField("ClosureComments", Closenotescomments);

							}
							if (cn == 18) {
								updatedBy = row.getCell(cn).getStringCellValue();
								if (!updatedBy.isEmpty() && updatedBy != null) {
									master.addField("UpdatedBy", updatedBy);
								}
							}
							if (cn == 19) {
								ETC_PREDICT_VAL = (float) row.getCell(cn).getNumericCellValue();
								master.addField("ETC_PREDICT_VALUE", ETC_PREDICT_VAL);

							}
							if (cn == 20) {
								overallSentiment = row.getCell(cn).getStringCellValue();
								if (!overallSentiment.isEmpty() && overallSentiment != null) {
									master.addField("OverallSentiment", overallSentiment);
								}

							}
							if (cn == 21) {
								surveyTaken = (int) row.getCell(cn).getNumericCellValue();
								if (surveyTaken == 0 || surveyTaken == 1) {
									master.addField("SurveyTaken", surveyTaken);
								}

							}
							if (cn == 22) {
								csatAwaitingActionSentiment = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionSentiment == 0 || csatAwaitingActionSentiment == 1) {
									master.addField("CSATAwaitingActionSentiment", csatAwaitingActionSentiment);
								}

							}
							if (cn == 23) {
								csatAwaitingActionETC = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionETC == 0 || csatAwaitingActionETC == 1) {
									master.addField("CSATAwaitingActionETC", csatAwaitingActionETC);
								}

							}
							if (cn == 24) {
								responseSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (responseSlaBreach == 0 || responseSlaBreach == 1 ||responseSlaBreach == 2) {
									master.addField("ResponseSlaBreach", responseSlaBreach);
								}

							}
							if (cn == 25) {
								resolutionSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (resolutionSlaBreach == 0 || resolutionSlaBreach == 1 ||responseSlaBreach == 2) {
									master.addField("ResolutionSlaBreach", resolutionSlaBreach);
								}

							}
							if (cn == 26) {
								suggestedVirtualSupervisior = row.getCell(cn).getStringCellValue();
								if (!suggestedVirtualSupervisior.isEmpty() && suggestedVirtualSupervisior != null) {
									master.addField("SuggestedVirtualSupervisior", suggestedVirtualSupervisior);
								}

							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = resolvedDate + valueAsInExcel;
									master.addField("ReportedTime", getDate(Updated));
								}
							}
							if (cn == 28) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("Level1", level1);
								}

							}
							if (cn == 29) {
								String level2 = row.getCell(cn).getStringCellValue();
								if (!level2.isEmpty() && level2 != null) {
									master.addField("Level2", level2);
								}

							}
							if (cn == 30) {
								String level3 = row.getCell(cn).getStringCellValue();
								if (!level3.isEmpty() && level3 != null) {
									master.addField("Level3", level3);
								}

							}
							if (cn == 31) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("TvmKeywordVar", level1);
								}
							}
							if (cn == 32) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMinValue", tvm);
								}
							}
							if (cn == 33) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMaxValue", tvm);
								}
							}

						}

					}

					if (row.getRowNum() > 25 && row.getRowNum() <= 45) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								Assignedto = row.getCell(cn).getStringCellValue();
								if (!Assignedto.isEmpty() && Assignedto != null) {
									master.addField("Assignee", Assignedto);
								}
							}
							if (cn == 1) {
								Assignmentgroup = row.getCell(cn).getStringCellValue();
								master.addField("AssignedGroup", Assignmentgroup);
							}
							if (cn == 2) {
								Category = row.getCell(cn).getStringCellValue();
								master.addField("Department", Category);
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Closed = resolvedDate + valueAsInExcel;
									master.addField("CompletedDate", getDate(Closed));
								}
							}
							if (cn == 4) {
								component = row.getCell(cn).getStringCellValue();
								master.addField("ReportedMode", component);
							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								String created = getUpdatedDate() + valueAsInExcel;
								master.addField("CreatedDate", getDate(created));
							}
							if (cn == 6) {

								ShortDescription = row.getCell(cn).getStringCellValue();
								String shortDescription = Jsoup.parse(ShortDescription.toString()).text();
								shortDescription = shortDescription.toString().replaceAll("(\r\n|\n)", " ");
								shortDescription = CleanInvalidXmlChars(shortDescription.toString(), "");

								master.addField("Description", shortDescription);
							}
							if (cn == 7) {
								Incidentstate = row.getCell(cn).getStringCellValue();
								master.addField("Status", Incidentstate);
							}
							if (cn == 8) {
								productCategory = row.getCell(cn).getStringCellValue();
								master.addField("ProductCategory", productCategory);
							}
							if (cn == 9) {
								number = (int) row.getCell(cn).getNumericCellValue();
								master.addField("IncidentID", number);
							}
							if (cn == 10) {
								OnBehalfOf = row.getCell(cn).getStringCellValue();
								if (!OnBehalfOf.isEmpty() && OnBehalfOf != null) {
									
									master.addField("FirstName", OnBehalfOf);
									
								}
							}
							if (cn == 11) {
								emailId = row.getCell(cn).getStringCellValue();
								master.addField("EmailId", emailId);
							}
							if (cn == 12) {
								Priority = row.getCell(cn).getStringCellValue();
								master.addField("Priority", Priority);
							}
							if (cn == 13) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Resolved = resolvedDate + valueAsInExcel;
									master.addField("ResolvedDate", getDate(Resolved));
								}
							}

							if (cn == 14) {
								desc = row.getCell(cn).getStringCellValue();
								String descr = Jsoup.parse(desc.toString()).text();
								descr = descr.toString().replaceAll("(\r\n|\n)", " ");
								descr = CleanInvalidXmlChars(descr.toString(), "");

								master.addField("Title", descr);
							}

							if (cn == 15) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									String uDate = getUpdatedDate() + valueAsInExcel;
									master.addField("LastModifiedDate", getDate(uDate));
								}
							}
							if (cn == 16) {

								notes = row.getCell(cn).getStringCellValue();
								if (!notes.isEmpty() && notes != null) {
																
								if (String.valueOf(notes).contains("Additional comments")) {

									String comment = String.valueOf(notes)
											.replace("(Additional comments (Customer Visible))", "~");
									String[] comments = String.valueOf(comment).split("~");
									String notesCreatedByDetails = comments[0];

									String[] NotesCreated = String.valueOf(notesCreatedByDetails).split(" -");
									String notesCreatedDate = NotesCreated[0];
									String notesCreatedBy = NotesCreated[1];

									String note = comments[1];

									note = note.toString().replaceAll("(\r\n|\n)", " ");
									note = CleanInvalidXmlChars(note.toString(), "");

									String actualNotesCreatedBy = notesCreatedByDetails.substring(22);

									if (!notes.isEmpty() && notes != null) {
										master.addField("Note", note);
									}
									if (!notesCreatedBy.isEmpty() && notesCreatedBy != null) {
										master.addField("NotesCreatedBy", actualNotesCreatedBy);
									}
									if (!notesCreatedDate.isEmpty() && notesCreatedDate != null) {
										master.addField("NotesCreatedDate", master.getFieldValue("LastModifiedDate"));
									}

								}
								}
							}
							if (cn == 17) {
								closenotescomments = row.getCell(cn).getStringCellValue();
								String Closenotescomments = Jsoup.parse(closenotescomments.toString()).text();
								Closenotescomments = Closenotescomments.toString().replaceAll("(\r\n|\n)", " ");
								Closenotescomments = CleanInvalidXmlChars(Closenotescomments.toString(), "");
								master.addField("ClosureComments", Closenotescomments);

							}
							if (cn == 18) {
								updatedBy = row.getCell(cn).getStringCellValue();
								if (!updatedBy.isEmpty() && updatedBy != null) {
									master.addField("UpdatedBy", updatedBy);
								}
							}
							if (cn == 19) {
								ETC_PREDICT_VAL = (float) row.getCell(cn).getNumericCellValue();
								master.addField("ETC_PREDICT_VALUE", ETC_PREDICT_VAL);

							}
							if (cn == 20) {
								overallSentiment = row.getCell(cn).getStringCellValue();
								if (!overallSentiment.isEmpty() && overallSentiment != null) {
									master.addField("OverallSentiment", overallSentiment);
								}

							}
							if (cn == 21) {
								surveyTaken = (int) row.getCell(cn).getNumericCellValue();
								if (surveyTaken == 0 || surveyTaken == 1) {
									master.addField("SurveyTaken", surveyTaken);
								}

							}
							if (cn == 22) {
								csatAwaitingActionSentiment = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionSentiment == 0 || csatAwaitingActionSentiment == 1) {
									master.addField("CSATAwaitingActionSentiment", csatAwaitingActionSentiment);
								}

							}
							if (cn == 23) {
								csatAwaitingActionETC = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionETC == 0 || csatAwaitingActionETC == 1) {
									master.addField("CSATAwaitingActionETC", csatAwaitingActionETC);
								}

							}
							if (cn == 24) {
								responseSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (responseSlaBreach == 0 || responseSlaBreach == 1 ||responseSlaBreach == 2) {
									master.addField("ResponseSlaBreach", responseSlaBreach);
								}

							}
							if (cn == 25) {
								resolutionSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (resolutionSlaBreach == 0 || resolutionSlaBreach == 1 ||responseSlaBreach == 2) {
									master.addField("ResolutionSlaBreach", resolutionSlaBreach);
								}

							}
							if (cn == 26) {
								suggestedVirtualSupervisior = row.getCell(cn).getStringCellValue();
								if (!suggestedVirtualSupervisior.isEmpty() && suggestedVirtualSupervisior != null) {
									master.addField("SuggestedVirtualSupervisior", suggestedVirtualSupervisior);
								}

							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									String uDate = getUpdatedDate() + valueAsInExcel;
									master.addField("ReportedTime", getDate(uDate));
								}
							}
							if (cn == 28) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("Level1", level1);
								}

							}
							if (cn == 29) {
								String level2 = row.getCell(cn).getStringCellValue();
								if (!level2.isEmpty() && level2 != null) {
									master.addField("Level2", level2);
								}

							}
							if (cn == 30) {
								String level3 = row.getCell(cn).getStringCellValue();
								if (!level3.isEmpty() && level3 != null) {
									master.addField("Level3", level3);
								}

							}
							if (cn == 31) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("TvmKeywordVar", level1);
								}
							}
							if (cn == 32) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMinValue", tvm);
								}
							}
							if (cn == 33) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMaxValue", tvm);
								}
							}

						}

					}
					
					if (row.getRowNum() > 45 && row.getRowNum() <= 97) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								Assignedto = row.getCell(cn).getStringCellValue();
								if (!Assignedto.isEmpty() && Assignedto != null) {
									master.addField("Assignee", Assignedto);
								}
							}
							if (cn == 1) {
								Assignmentgroup = row.getCell(cn).getStringCellValue();
								master.addField("AssignedGroup", Assignmentgroup);
							}
							if (cn == 2) {
								Category = row.getCell(cn).getStringCellValue();
								master.addField("Department", Category);
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Closed = resolvedDate + valueAsInExcel;
									master.addField("CompletedDate", getDate(Closed));
								}
							}
							if (cn == 4) {
								component = row.getCell(cn).getStringCellValue();
								master.addField("ItemName", component);
							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								createdOn = createdDate + valueAsInExcel;
								master.addField("CreatedDate", getDate(createdOn));
							}
							if (cn == 6) {

								ShortDescription = row.getCell(cn).getStringCellValue();
								String shortDescription = Jsoup.parse(ShortDescription.toString()).text();
								shortDescription = shortDescription.toString().replaceAll("(\r\n|\n)", " ");
								shortDescription = CleanInvalidXmlChars(shortDescription.toString(), "");

								master.addField("Description", shortDescription);
							}
							if (cn == 7) {
								Incidentstate = row.getCell(cn).getStringCellValue();
								master.addField("Status", Incidentstate);
							}
							if (cn == 8) {
								productCategory = row.getCell(cn).getStringCellValue();
								master.addField("ProductCategory", productCategory);
							}
							if (cn == 9) {
								number = (int) row.getCell(cn).getNumericCellValue();
								master.addField("IncidentID", number);
							}
							if (cn == 10) {
								OnBehalfOf = row.getCell(cn).getStringCellValue();
								if (!OnBehalfOf.isEmpty() && OnBehalfOf != null) {
																		
									master.addField("FirstName", OnBehalfOf);
								}
							}
							if (cn == 11) {
								emailId = row.getCell(cn).getStringCellValue();
								master.addField("EmailId", emailId);
							}
							if (cn == 12) {
								Priority = row.getCell(cn).getStringCellValue();
								master.addField("Priority", Priority);
							}
							if (cn == 13) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Resolved = resolvedDate + valueAsInExcel;
									master.addField("ResolvedDate", getDate(Resolved));
								}
							}

							if (cn == 14) {
								desc = row.getCell(cn).getStringCellValue();
								String descr = Jsoup.parse(desc.toString()).text();
								descr = descr.toString().replaceAll("(\r\n|\n)", " ");
								descr = CleanInvalidXmlChars(descr.toString(), "");

								master.addField("Title", descr);
							}

							if (cn == 15) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = updatedDate + valueAsInExcel;
									master.addField("LastModifiedDate", getDate(Updated));
								}
							}
							if (cn == 16) {

								notes = row.getCell(cn).getStringCellValue();
								if (!notes.isEmpty() && notes != null) {
								if (String.valueOf(notes).contains("Additional comments")) {

									String comment = String.valueOf(notes)
											.replace("(Additional comments (Customer Visible))", "~");
									String[] comments = String.valueOf(comment).split("~");
									String notesCreatedByDetails = comments[0];

									String[] NotesCreated = String.valueOf(notesCreatedByDetails).split(" -");
									String notesCreatedDate = NotesCreated[0];
									String notesCreatedBy = NotesCreated[1];

									String note = comments[1];

									note = note.toString().replaceAll("(\r\n|\n)", " ");
									note = CleanInvalidXmlChars(note.toString(), "");

									String actualNotesCreatedBy = notesCreatedByDetails.substring(22);

									if (!notes.isEmpty() && notes != null) {
										master.addField("Note", note);
									}
									if (!notesCreatedBy.isEmpty() && notesCreatedBy != null) {
										master.addField("NotesCreatedBy", actualNotesCreatedBy);
									}
									if (!notesCreatedDate.isEmpty() && notesCreatedDate != null) {
										master.addField("NotesCreatedDate", master.getFieldValue("LastModifiedDate"));
									}

								}
								}
							}
							if (cn == 17) {
								closenotescomments = row.getCell(cn).getStringCellValue();
								String Closenotescomments = Jsoup.parse(closenotescomments.toString()).text();
								Closenotescomments = Closenotescomments.toString().replaceAll("(\r\n|\n)", " ");
								Closenotescomments = CleanInvalidXmlChars(Closenotescomments.toString(), "");
								master.addField("ClosureComments", Closenotescomments);

							}
							if (cn == 18) {
								updatedBy = row.getCell(cn).getStringCellValue();
								if (!updatedBy.isEmpty() && updatedBy != null) {
									master.addField("UpdatedBy", updatedBy);
								}
							}
							if (cn == 19) {
								ETC_PREDICT_VAL = (float) row.getCell(cn).getNumericCellValue();
								master.addField("ETC_PREDICT_VALUE", ETC_PREDICT_VAL);

							}
							if (cn == 20) {
								overallSentiment = row.getCell(cn).getStringCellValue();
								if (!overallSentiment.isEmpty() && overallSentiment != null) {
									master.addField("OverallSentiment", overallSentiment);
								}

							}
							if (cn == 21) {
								surveyTaken = (int) row.getCell(cn).getNumericCellValue();
								if (surveyTaken == 0 || surveyTaken == 1) {
									master.addField("SurveyTaken", surveyTaken);
								}

							}
							if (cn == 22) {
								csatAwaitingActionSentiment = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionSentiment == 0 || csatAwaitingActionSentiment == 1) {
									master.addField("CSATAwaitingActionSentiment", csatAwaitingActionSentiment);
								}

							}
							if (cn == 23) {
								csatAwaitingActionETC = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionETC == 0 || csatAwaitingActionETC == 1) {
									master.addField("CSATAwaitingActionETC", csatAwaitingActionETC);
								}

							}
							if (cn == 24) {
								responseSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (responseSlaBreach == 0 || responseSlaBreach == 1 || responseSlaBreach == 2) {
									master.addField("ResponseSlaBreach", responseSlaBreach);
								}

							}
							if (cn == 25) {
								resolutionSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (resolutionSlaBreach == 0 || resolutionSlaBreach == 1 || responseSlaBreach == 2) {
									master.addField("ResolutionSlaBreach", resolutionSlaBreach);
								}

							}
							if (cn == 26) {
								suggestedVirtualSupervisior = row.getCell(cn).getStringCellValue();
								if (!suggestedVirtualSupervisior.isEmpty() && suggestedVirtualSupervisior != null) {
									master.addField("SuggestedVirtualSupervisior", suggestedVirtualSupervisior);
								}

							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = updatedDate + valueAsInExcel;
									master.addField("ReportedTime", getDate(Updated));
								}
							}
							if (cn == 28) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("Level1", level1);
								}

							}
							if (cn == 29) {
								String level2 = row.getCell(cn).getStringCellValue();
								if (!level2.isEmpty() && level2 != null) {
									master.addField("Level2", level2);
								}

							}
							if (cn == 30) {
								String level3 = row.getCell(cn).getStringCellValue();
								if (!level3.isEmpty() && level3 != null) {
									master.addField("Level3", level3);
								}

							}
							if (cn == 31) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("TvmKeywordVar", level1);
								}
							}
							if (cn == 32) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMinValue", tvm);
								}
							}
							if (cn == 33) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMaxValue", tvm);
								}
							}

						}

					}
					if (row.getRowNum() > 97 && row.getRowNum() <= 102) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								Assignedto = row.getCell(cn).getStringCellValue();
								if (!Assignedto.isEmpty() && Assignedto != null) {
									master.addField("Assignee", Assignedto);

								}
							}
							if (cn == 1) {
								Assignmentgroup = row.getCell(cn).getStringCellValue();
								master.addField("AssignedGroup", Assignmentgroup);
							}
							if (cn == 2) {
								Category = row.getCell(cn).getStringCellValue();
								master.addField("Department", Category);
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Closed = getMonthDate() + valueAsInExcel;
									master.addField("CompletedDate", getDate(Closed));
								}
							}
							if (cn == 4) {
								component = row.getCell(cn).getStringCellValue();
								master.addField("ItemName", component);
							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								createdOn = getMonthDate() + valueAsInExcel;
								master.addField("CreatedDate", getDate(createdOn));
							}
							if (cn == 6) {

								ShortDescription = row.getCell(cn).getStringCellValue();
								String shortDescription = Jsoup.parse(ShortDescription.toString()).text();
								shortDescription = shortDescription.toString().replaceAll("(\r\n|\n)", " ");
								shortDescription = CleanInvalidXmlChars(shortDescription.toString(), "");

								master.addField("Description", shortDescription);
							}
							if (cn == 7) {
								Incidentstate = row.getCell(cn).getStringCellValue();
								master.addField("Status", Incidentstate);
							}
							if (cn == 8) {
								productCategory = row.getCell(cn).getStringCellValue();
								master.addField("ProductCategory", productCategory);
							}
							if (cn == 9) {
								number = (int) row.getCell(cn).getNumericCellValue();
								master.addField("IncidentID", number);
							}
							if (cn == 10) {
								OnBehalfOf = row.getCell(cn).getStringCellValue();
								if (!OnBehalfOf.isEmpty() && OnBehalfOf != null) {
																		
									master.addField("FirstName", OnBehalfOf);
									
								}
							}
							if (cn == 11) {
								emailId = row.getCell(cn).getStringCellValue();
								master.addField("EmailId", emailId);
							}
							if (cn == 12) {
								Priority = row.getCell(cn).getStringCellValue();
								master.addField("Priority", Priority);
							}
							if (cn == 13) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Resolved = getMonthDate() + valueAsInExcel;
									master.addField("ResolvedDate", getDate(Resolved));
								}
							}

							if (cn == 14) {
								desc = row.getCell(cn).getStringCellValue();
								String descr = Jsoup.parse(desc.toString()).text();
								descr = descr.toString().replaceAll("(\r\n|\n)", " ");
								descr = CleanInvalidXmlChars(descr.toString(), "");

								master.addField("Title", descr);
							}

							if (cn == 15) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = getMonthDate() + valueAsInExcel;
									master.addField("LastModifiedDate", getDate(Updated));
								}
							}
							if (cn == 16) {

								notes = row.getCell(cn).getStringCellValue();
								if (!notes.isEmpty() && notes != null) {
								if (String.valueOf(notes).contains("Additional comments")) {

									String comment = String.valueOf(notes)
											.replace("(Additional comments (Customer Visible))", "~");
									String[] comments = String.valueOf(comment).split("~");
									String notesCreatedByDetails = comments[0];

									String[] NotesCreated = String.valueOf(notesCreatedByDetails).split(" -");
									String notesCreatedDate = NotesCreated[0];
									String notesCreatedBy = NotesCreated[1];

									String note = comments[1];

									note = note.toString().replaceAll("(\r\n|\n)", " ");
									note = CleanInvalidXmlChars(note.toString(), "");

									String actualNotesCreatedBy = notesCreatedByDetails.substring(22);

									if (!notes.isEmpty() && notes != null) {
										master.addField("Note", note);
									}
									if (!notesCreatedBy.isEmpty() && notesCreatedBy != null) {
										master.addField("NotesCreatedBy", actualNotesCreatedBy);
									}
									if (!notesCreatedDate.isEmpty() && notesCreatedDate != null) {
										master.addField("NotesCreatedDate", master.getFieldValue("LastModifiedDate"));
									}

								}
								}
							}
							if (cn == 17) {
								closenotescomments = row.getCell(cn).getStringCellValue();
								String Closenotescomments = Jsoup.parse(closenotescomments.toString()).text();
								Closenotescomments = Closenotescomments.toString().replaceAll("(\r\n|\n)", " ");
								Closenotescomments = CleanInvalidXmlChars(Closenotescomments.toString(), "");
								master.addField("ClosureComments", Closenotescomments);

							}
							if (cn == 18) {
								updatedBy = row.getCell(cn).getStringCellValue();
								if (!updatedBy.isEmpty() && updatedBy != null) {
									master.addField("UpdatedBy", updatedBy);
								}
							}
							if (cn == 19) {
								ETC_PREDICT_VAL = (float) row.getCell(cn).getNumericCellValue();
								master.addField("ETC_PREDICT_VALUE", ETC_PREDICT_VAL);

							}
							if (cn == 20) {
								overallSentiment = row.getCell(cn).getStringCellValue();
								if (!overallSentiment.isEmpty() && overallSentiment != null) {
									master.addField("OverallSentiment", overallSentiment);
								}

							}
							if (cn == 21) {
								surveyTaken = (int) row.getCell(cn).getNumericCellValue();
								if (surveyTaken == 0 || surveyTaken == 1) {
									master.addField("SurveyTaken", surveyTaken);
								}

							}
							if (cn == 22) {
								csatAwaitingActionSentiment = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionSentiment == 0 || csatAwaitingActionSentiment == 1) {
									master.addField("CSATAwaitingActionSentiment", csatAwaitingActionSentiment);
								}

							}
							if (cn == 23) {
								csatAwaitingActionETC = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionETC == 0 || csatAwaitingActionETC == 1) {
									master.addField("CSATAwaitingActionETC", csatAwaitingActionETC);
								}

							}
							if (cn == 24) {
								responseSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (responseSlaBreach == 0 || responseSlaBreach == 1 || responseSlaBreach == 2) {
									master.addField("ResponseSlaBreach", responseSlaBreach);
								}

							}
							if (cn == 25) {
								resolutionSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (resolutionSlaBreach == 0 || resolutionSlaBreach == 1 || responseSlaBreach == 2) {
									master.addField("ResolutionSlaBreach", resolutionSlaBreach);
								}

							}
							if (cn == 26) {
								suggestedVirtualSupervisior = row.getCell(cn).getStringCellValue();
								if (!suggestedVirtualSupervisior.isEmpty() && suggestedVirtualSupervisior != null) {
									master.addField("SuggestedVirtualSupervisior", suggestedVirtualSupervisior);
								}

							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = updatedDate + valueAsInExcel;
									master.addField("ReportedTime", getDate(Updated));
								}
							}
							if (cn == 28) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("Level1", level1);
								}

							}
							if (cn == 29) {
								String level2 = row.getCell(cn).getStringCellValue();
								if (!level2.isEmpty() && level2 != null) {
									master.addField("Level2", level2);
								}

							}
							if (cn == 30) {
								String level3 = row.getCell(cn).getStringCellValue();
								if (!level3.isEmpty() && level3 != null) {
									master.addField("Level3", level3);
								}

							}
							if (cn == 31) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("TvmKeywordVar", level1);
								}
							}
							if (cn == 32) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMinValue", tvm);
								}
							}
							if (cn == 33) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMaxValue", tvm);
								}
							}

						}

					}
					if (row.getRowNum() > 102 && row.getRowNum() <= 111) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								Assignedto = row.getCell(cn).getStringCellValue();
								if (!Assignedto.isEmpty() && Assignedto != null) {
									master.addField("Assignee", Assignedto);
								}
							}
							if (cn == 1) {
								Assignmentgroup = row.getCell(cn).getStringCellValue();
								master.addField("AssignedGroup", Assignmentgroup);
							}
							if (cn == 2) {
								Category = row.getCell(cn).getStringCellValue();
								master.addField("Department", Category);
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Closed = get30To90Dates() + valueAsInExcel;
									master.addField("CompletedDate", getDate(Closed));
								}
							}
							if (cn == 4) {
								component = row.getCell(cn).getStringCellValue();
								master.addField("ItemName", component);
							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								createdOn = get30To90Dates() + valueAsInExcel;
								master.addField("CreatedDate", getDate(createdOn));
							}
							if (cn == 6) {

								ShortDescription = row.getCell(cn).getStringCellValue();
								String shortDescription = Jsoup.parse(ShortDescription.toString()).text();
								shortDescription = shortDescription.toString().replaceAll("(\r\n|\n)", " ");
								shortDescription = CleanInvalidXmlChars(shortDescription.toString(), "");

								master.addField("Description", shortDescription);
							}
							if (cn == 7) {
								Incidentstate = row.getCell(cn).getStringCellValue();
								master.addField("Status", Incidentstate);
							}
							if (cn == 8) {
								productCategory = row.getCell(cn).getStringCellValue();
								master.addField("ProductCategory", productCategory);
							}
							if (cn == 9) {
								number = (int) row.getCell(cn).getNumericCellValue();
								master.addField("IncidentID", number);
							}
							if (cn == 10) {
								OnBehalfOf = row.getCell(cn).getStringCellValue();
								if (!OnBehalfOf.isEmpty() && OnBehalfOf != null) {
									
									master.addField("FirstName", OnBehalfOf);
								}
							}
							if (cn == 11) {
								emailId = row.getCell(cn).getStringCellValue();
								master.addField("EmailId", emailId);
							}
							if (cn == 12) {
								Priority = row.getCell(cn).getStringCellValue();
								master.addField("Priority", Priority);
							}
							if (cn == 13) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Resolved = get30To90Dates() + valueAsInExcel;
									master.addField("ResolvedDate", getDate(Resolved));
								}
							}

							if (cn == 14) {
								desc = row.getCell(cn).getStringCellValue();
								String descr = Jsoup.parse(desc.toString()).text();
								descr = descr.toString().replaceAll("(\r\n|\n)", " ");
								descr = CleanInvalidXmlChars(descr.toString(), "");

								master.addField("Title", descr);
							}

							if (cn == 15) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = get30To90Dates() + valueAsInExcel;
									master.addField("LastModifiedDate", getDate(Updated));
								}
							}
							if (cn == 16) {

								notes = row.getCell(cn).getStringCellValue();
								if (!notes.isEmpty() && notes != null) {
								if (String.valueOf(notes).contains("Additional comments")) {

									String comment = String.valueOf(notes)
											.replace("(Additional comments (Customer Visible))", "~");
									String[] comments = String.valueOf(comment).split("~");
									String notesCreatedByDetails = comments[0];

									String[] NotesCreated = String.valueOf(notesCreatedByDetails).split(" -");
									String notesCreatedDate = NotesCreated[0];
									String notesCreatedBy = NotesCreated[1];

									String note = comments[1];

									note = note.toString().replaceAll("(\r\n|\n)", " ");
									note = CleanInvalidXmlChars(note.toString(), "");

									String actualNotesCreatedBy = notesCreatedByDetails.substring(22);

									if (!notes.isEmpty() && notes != null) {
										master.addField("Note", note);
									}
									if (!notesCreatedBy.isEmpty() && notesCreatedBy != null) {
										master.addField("NotesCreatedBy", actualNotesCreatedBy);
									}
									if (!notesCreatedDate.isEmpty() && notesCreatedDate != null) {
										master.addField("NotesCreatedDate", master.getFieldValue("LastModifiedDate"));
									}

								}
								}
							}
							if (cn == 17) {
								closenotescomments = row.getCell(cn).getStringCellValue();
								String Closenotescomments = Jsoup.parse(closenotescomments.toString()).text();
								Closenotescomments = Closenotescomments.toString().replaceAll("(\r\n|\n)", " ");
								Closenotescomments = CleanInvalidXmlChars(Closenotescomments.toString(), "");
								master.addField("ClosureComments", Closenotescomments);

							}
							if (cn == 18) {
								updatedBy = row.getCell(cn).getStringCellValue();
								if (!updatedBy.isEmpty() && updatedBy != null) {
									master.addField("UpdatedBy", updatedBy);
								}
							}
							if (cn == 19) {
								ETC_PREDICT_VAL = (float) row.getCell(cn).getNumericCellValue();
								master.addField("ETC_PREDICT_VALUE", ETC_PREDICT_VAL);

							}
							if (cn == 20) {
								overallSentiment = row.getCell(cn).getStringCellValue();
								if (!overallSentiment.isEmpty() && overallSentiment != null) {
									master.addField("OverallSentiment", overallSentiment);
								}

							}
							if (cn == 21) {
								surveyTaken = (int) row.getCell(cn).getNumericCellValue();
								if (surveyTaken == 0 || surveyTaken == 1) {
									master.addField("SurveyTaken", surveyTaken);
								}

							}
							if (cn == 22) {
								csatAwaitingActionSentiment = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionSentiment == 0 || csatAwaitingActionSentiment == 1) {
									master.addField("CSATAwaitingActionSentiment", csatAwaitingActionSentiment);
								}

							}
							if (cn == 23) {
								csatAwaitingActionETC = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionETC == 0 || csatAwaitingActionETC == 1) {
									master.addField("CSATAwaitingActionETC", csatAwaitingActionETC);
								}

							}
							if (cn == 24) {
								responseSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (responseSlaBreach == 0 || responseSlaBreach == 1 ||responseSlaBreach == 2) {
									master.addField("ResponseSlaBreach", responseSlaBreach);
								}

							}
							if (cn == 25) {
								resolutionSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (resolutionSlaBreach == 0 || resolutionSlaBreach == 1 ||responseSlaBreach == 2) {
									master.addField("ResolutionSlaBreach", resolutionSlaBreach);
								}

							}
							if (cn == 26) {
								suggestedVirtualSupervisior = row.getCell(cn).getStringCellValue();
								if (!suggestedVirtualSupervisior.isEmpty() && suggestedVirtualSupervisior != null) {
									master.addField("SuggestedVirtualSupervisior", suggestedVirtualSupervisior);
								}

							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = updatedDate + valueAsInExcel;
									master.addField("ReportedTime", getDate(Updated));
								}
							}
							if (cn == 28) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("Level1", level1);
								}

							}
							if (cn == 29) {
								String level2 = row.getCell(cn).getStringCellValue();
								if (!level2.isEmpty() && level2 != null) {
									master.addField("Level2", level2);
								}

							}
							if (cn == 30) {
								String level3 = row.getCell(cn).getStringCellValue();
								if (!level3.isEmpty() && level3 != null) {
									master.addField("Level3", level3);
								}

							}
							if (cn == 31) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("TvmKeywordVar", level1);
								}
							}
							if (cn == 32) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMinValue", tvm);
								}
							}
							if (cn == 33) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMaxValue", tvm);
								}
							}

						}

					}
					if (row.getRowNum() > 111 && row.getRowNum() <= 117) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								Assignedto = row.getCell(cn).getStringCellValue();
								if (!Assignedto.isEmpty() && Assignedto != null) {
									master.addField("Assignee", Assignedto);
								}
							}
							if (cn == 1) {
								Assignmentgroup = row.getCell(cn).getStringCellValue();
								master.addField("AssignedGroup", Assignmentgroup);
							}
							if (cn == 2) {
								Category = row.getCell(cn).getStringCellValue();
								master.addField("SubCategory", Category);
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Closed = get90To180Dates() + valueAsInExcel;
									master.addField("CompletedDate", getDate(Closed));
								}
							}
							if (cn == 4) {
								component = row.getCell(cn).getStringCellValue();
								master.addField("ItemName", component);
							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								createdOn = get90To180Dates() + valueAsInExcel;
								master.addField("CreatedDate", getDate(createdOn));
							}
							if (cn == 6) {

								ShortDescription = row.getCell(cn).getStringCellValue();
								String shortDescription = Jsoup.parse(ShortDescription.toString()).text();
								shortDescription = shortDescription.toString().replaceAll("(\r\n|\n)", " ");
								shortDescription = CleanInvalidXmlChars(shortDescription.toString(), "");

								master.addField("Description", shortDescription);
							}
							if (cn == 7) {
								Incidentstate = row.getCell(cn).getStringCellValue();
								master.addField("Status", Incidentstate);
							}
							if (cn == 8) {
								productCategory = row.getCell(cn).getStringCellValue();
								master.addField("ProductCategory", productCategory);
							}
							if (cn == 9) {
								number = (int) row.getCell(cn).getNumericCellValue();
								master.addField("IncidentID", number);
							}
							if (cn == 10) {
								OnBehalfOf = row.getCell(cn).getStringCellValue();
								if (!OnBehalfOf.isEmpty() && OnBehalfOf != null) {
									master.addField("FirstName", OnBehalfOf);
								}
							}
							if (cn == 11) {
								emailId = row.getCell(cn).getStringCellValue();
								master.addField("EmailId", emailId);
							}
							if (cn == 12) {
								Priority = row.getCell(cn).getStringCellValue();
								master.addField("Priority", Priority);
							}
							if (cn == 13) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Resolved = get90To180Dates() + valueAsInExcel;
									master.addField("ResolvedDate", getDate(Resolved));
								}
							}

							if (cn == 14) {
								desc = row.getCell(cn).getStringCellValue();
								String descr = Jsoup.parse(desc.toString()).text();
								descr = descr.toString().replaceAll("(\r\n|\n)", " ");
								descr = CleanInvalidXmlChars(descr.toString(), "");

								master.addField("Title", descr);
							}

							if (cn == 15) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = get90To180Dates() + valueAsInExcel;
									master.addField("LastModifiedDate", getDate(Updated));
								}
							}
							if (cn == 16) {

								notes = row.getCell(cn).getStringCellValue();
								if (!notes.isEmpty() && notes != null) {
								if (String.valueOf(notes).contains("Additional comments")) {

									String comment = String.valueOf(notes)
											.replace("(Additional comments (Customer Visible))", "~");
									String[] comments = String.valueOf(comment).split("~");
									String notesCreatedByDetails = comments[0];

									String[] NotesCreated = String.valueOf(notesCreatedByDetails).split(" -");
									String notesCreatedDate = NotesCreated[0];
									String notesCreatedBy = NotesCreated[1];

									String note = comments[1];

									note = note.toString().replaceAll("(\r\n|\n)", " ");
									note = CleanInvalidXmlChars(note.toString(), "");

									String actualNotesCreatedBy = notesCreatedByDetails.substring(22);

									if (!notes.isEmpty() && notes != null) {
										master.addField("Note", note);
									}
									if (!notesCreatedBy.isEmpty() && notesCreatedBy != null) {
										master.addField("NotesCreatedBy", actualNotesCreatedBy);
									}
									if (!notesCreatedDate.isEmpty() && notesCreatedDate != null) {
										master.addField("NotesCreatedDate", master.getFieldValue("LastModifiedDate"));
									}

								}
								}
							}
							if (cn == 17) {
								closenotescomments = row.getCell(cn).getStringCellValue();
								String Closenotescomments = Jsoup.parse(closenotescomments.toString()).text();
								Closenotescomments = Closenotescomments.toString().replaceAll("(\r\n|\n)", " ");
								Closenotescomments = CleanInvalidXmlChars(Closenotescomments.toString(), "");
								master.addField("ClosureComments", Closenotescomments);

							}
							if (cn == 18) {
								updatedBy = row.getCell(cn).getStringCellValue();
								if (!updatedBy.isEmpty() && updatedBy != null) {
									master.addField("UpdatedBy", updatedBy);
								}
							}
							if (cn == 19) {
								ETC_PREDICT_VAL = (float) row.getCell(cn).getNumericCellValue();
								master.addField("ETC_PREDICT_VALUE", ETC_PREDICT_VAL);

							}
							if (cn == 20) {
								overallSentiment = row.getCell(cn).getStringCellValue();
								if (!overallSentiment.isEmpty() && overallSentiment != null) {
									master.addField("OverallSentiment", overallSentiment);
								}

							}
							if (cn == 21) {
								surveyTaken = (int) row.getCell(cn).getNumericCellValue();
								if (surveyTaken == 0 || surveyTaken == 1) {
									master.addField("SurveyTaken", surveyTaken);
								}

							}
							if (cn == 22) {
								csatAwaitingActionSentiment = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionSentiment == 0 || csatAwaitingActionSentiment == 1) {
									master.addField("CSATAwaitingActionSentiment", csatAwaitingActionSentiment);
								}

							}
							if (cn == 23) {
								csatAwaitingActionETC = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionETC == 0 || csatAwaitingActionETC == 1) {
									master.addField("CSATAwaitingActionETC", csatAwaitingActionETC);
								}

							}
							if (cn == 24) {
								responseSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (responseSlaBreach == 0 || responseSlaBreach == 1 || responseSlaBreach == 2) {
									master.addField("ResponseSlaBreach", responseSlaBreach);
								}

							}
							if (cn == 25) {
								resolutionSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (resolutionSlaBreach == 0 || resolutionSlaBreach == 1 || resolutionSlaBreach == 2) {
									master.addField("ResolutionSlaBreach", resolutionSlaBreach);
								}

							}
							if (cn == 26) {
								suggestedVirtualSupervisior = row.getCell(cn).getStringCellValue();
								if (!suggestedVirtualSupervisior.isEmpty() && suggestedVirtualSupervisior != null) {
									master.addField("SuggestedVirtualSupervisior", suggestedVirtualSupervisior);
								}

							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = updatedDate + valueAsInExcel;
									master.addField("ReportedTime", getDate(Updated));
								}
							}
							if (cn == 28) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("Level1", level1);
								}

							}
							if (cn == 29) {
								String level2 = row.getCell(cn).getStringCellValue();
								if (!level2.isEmpty() && level2 != null) {
									master.addField("Level2", level2);
								}

							}
							if (cn == 30) {
								String level3 = row.getCell(cn).getStringCellValue();
								if (!level3.isEmpty() && level3 != null) {
									master.addField("Level3", level3);
								}

							}
							if (cn == 31) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("TvmKeywordVar", level1);
								}
							}
							if (cn == 32) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMinValue", tvm);
								}
							}
							if (cn == 33) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMaxValue", tvm);
								}
							}

						}

					}
					if (row.getRowNum() > 117 && row.getRowNum() <= 121) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								Assignedto = row.getCell(cn).getStringCellValue();
								if (!Assignedto.isEmpty() && Assignedto != null) {
									master.addField("Assignee", Assignedto);

								}
							}
							if (cn == 1) {
								Assignmentgroup = row.getCell(cn).getStringCellValue();
								master.addField("AssignedGroup", Assignmentgroup);
							}
							if (cn == 2) {
								Category = row.getCell(cn).getStringCellValue();
								master.addField("SubCategory", Category);
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Closed = get180To360Dates() + valueAsInExcel;
									master.addField("CompletedDate", getDate(Closed));
								}
							}
							if (cn == 4) {
								component = row.getCell(cn).getStringCellValue();
								master.addField("ItemName", component);
							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								createdOn = get180To360Dates() + valueAsInExcel;
								master.addField("CreatedDate", getDate(createdOn));
							}
							if (cn == 6) {

								ShortDescription = row.getCell(cn).getStringCellValue();
								String shortDescription = Jsoup.parse(ShortDescription.toString()).text();
								shortDescription = shortDescription.toString().replaceAll("(\r\n|\n)", " ");
								shortDescription = CleanInvalidXmlChars(shortDescription.toString(), "");

								master.addField("Description", shortDescription);
							}
							if (cn == 7) {
								Incidentstate = row.getCell(cn).getStringCellValue();
								master.addField("Status", Incidentstate);
							}
							if (cn == 8) {
								productCategory = row.getCell(cn).getStringCellValue();
								master.addField("ProductCategory", productCategory);
							}
							if (cn == 9) {
								number = (int) row.getCell(cn).getNumericCellValue();
								master.addField("IncidentID", number);
							}
							if (cn == 10) {
								OnBehalfOf = row.getCell(cn).getStringCellValue();
								if (!OnBehalfOf.isEmpty() && OnBehalfOf != null) {
									
									master.addField("FirstName", OnBehalfOf);
								}
							}
							if (cn == 11) {
								emailId = row.getCell(cn).getStringCellValue();
								master.addField("EmailId", emailId);
							}
							if (cn == 12) {
								Priority = row.getCell(cn).getStringCellValue();
								master.addField("Priority", Priority);
							}
							if (cn == 13) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Resolved = get180To360Dates() + valueAsInExcel;
									master.addField("ResolvedDate", getDate(Resolved));
								}
							}

							if (cn == 14) {
								desc = row.getCell(cn).getStringCellValue();
								String descr = Jsoup.parse(desc.toString()).text();
								descr = descr.toString().replaceAll("(\r\n|\n)", " ");
								descr = CleanInvalidXmlChars(descr.toString(), "");

								master.addField("Title", descr);
							}

							if (cn == 15) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = get180To360Dates() + valueAsInExcel;
									master.addField("LastModifiedDate", getDate(Updated));
								}
							}
							if (cn == 16) {

								notes = row.getCell(cn).getStringCellValue();
								if (!notes.isEmpty() && notes != null) {
								if (String.valueOf(notes).contains("Additional comments")) {

									String comment = String.valueOf(notes)
											.replace("(Additional comments (Customer Visible))", "~");
									String[] comments = String.valueOf(comment).split("~");
									String notesCreatedByDetails = comments[0];

									String[] NotesCreated = String.valueOf(notesCreatedByDetails).split(" -");
									String notesCreatedDate = NotesCreated[0];
									String notesCreatedBy = NotesCreated[1];

									String note = comments[1];

									note = note.toString().replaceAll("(\r\n|\n)", " ");
									note = CleanInvalidXmlChars(note.toString(), "");

									String actualNotesCreatedBy = notesCreatedByDetails.substring(22);

									if (!notes.isEmpty() && notes != null) {
										master.addField("Note", note);
									}
									if (!notesCreatedBy.isEmpty() && notesCreatedBy != null) {
										master.addField("NotesCreatedBy", actualNotesCreatedBy);
									}
									if (!notesCreatedDate.isEmpty() && notesCreatedDate != null) {
										master.addField("NotesCreatedDate", master.getFieldValue("LastModifiedDate"));
									}

								}
								}
							}
							if (cn == 17) {
								closenotescomments = row.getCell(cn).getStringCellValue();
								String Closenotescomments = Jsoup.parse(closenotescomments.toString()).text();
								Closenotescomments = Closenotescomments.toString().replaceAll("(\r\n|\n)", " ");
								Closenotescomments = CleanInvalidXmlChars(Closenotescomments.toString(), "");
								master.addField("ClosureComments", Closenotescomments);

							}
							if (cn == 18) {
								updatedBy = row.getCell(cn).getStringCellValue();
								if (!updatedBy.isEmpty() && updatedBy != null) {
									master.addField("UpdatedBy", updatedBy);
								}
							}
							if (cn == 19) {
								ETC_PREDICT_VAL = (float) row.getCell(cn).getNumericCellValue();
								master.addField("ETC_PREDICT_VALUE", ETC_PREDICT_VAL);

							}
							if (cn == 20) {
								overallSentiment = row.getCell(cn).getStringCellValue();
								if (!overallSentiment.isEmpty() && overallSentiment != null) {
									master.addField("OverallSentiment", overallSentiment);
								}

							}
							if (cn == 21) {
								surveyTaken = (int) row.getCell(cn).getNumericCellValue();
								if (surveyTaken == 0 || surveyTaken == 1) {
									master.addField("SurveyTaken", surveyTaken);
								}

							}
							if (cn == 22) {
								csatAwaitingActionSentiment = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionSentiment == 0 || csatAwaitingActionSentiment == 1) {
									master.addField("CSATAwaitingActionSentiment", csatAwaitingActionSentiment);
								}

							}
							if (cn == 23) {
								csatAwaitingActionETC = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionETC == 0 || csatAwaitingActionETC == 1) {
									master.addField("CSATAwaitingActionETC", csatAwaitingActionETC);
								}

							}
							if (cn == 24) {
								responseSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (responseSlaBreach == 0 || responseSlaBreach == 1 || responseSlaBreach == 2) {
									master.addField("ResponseSlaBreach", responseSlaBreach);
								}

							}
							if (cn == 25) {
								resolutionSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (resolutionSlaBreach == 0 || resolutionSlaBreach == 1 || resolutionSlaBreach == 2) {
									master.addField("ResolutionSlaBreach", resolutionSlaBreach);
								}

							}
							if (cn == 26) {
								suggestedVirtualSupervisior = row.getCell(cn).getStringCellValue();
								if (!suggestedVirtualSupervisior.isEmpty() && suggestedVirtualSupervisior != null) {
									master.addField("SuggestedVirtualSupervisior", suggestedVirtualSupervisior);
								}

							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									String Update = get180To360Dates() + valueAsInExcel;
									master.addField("ReportedTime", getDate(Update));
								}
							}
							if (cn == 28) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("Level1", level1);
								}

							}
							if (cn == 29) {
								String level2 = row.getCell(cn).getStringCellValue();
								if (!level2.isEmpty() && level2 != null) {
									master.addField("Level2", level2);
								}

							}
							if (cn == 30) {
								String level3 = row.getCell(cn).getStringCellValue();
								if (!level3.isEmpty() && level3 != null) {
									master.addField("Level3", level3);
								}

							}
							if (cn == 31) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("TvmKeywordVar", level1);
								}
							}
							if (cn == 32) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMinValue", tvm);
								}
							}
							if (cn == 33) {
								int tvm = (int) row.getCell(cn).getNumericCellValue();
								if (tvm == 1 || tvm == 2 || tvm == 3) {
									master.addField("TvmPredictMaxValue", tvm);
								}
							}

						}

					}
					if (row.getRowNum() > 121 && row.getRowNum()<=150) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								Assignedto = row.getCell(cn).getStringCellValue();
								if (!Assignedto.isEmpty() && Assignedto != null) {
									master.addField("Assignee", Assignedto);

								}
							}
							if (cn == 1) {
								Assignmentgroup = row.getCell(cn).getStringCellValue();
								master.addField("AssignedGroup", Assignmentgroup);
							}
							if (cn == 2) {
								Category = row.getCell(cn).getStringCellValue();
								master.addField("SubCategory", Category);
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Closed = getweekDate() + valueAsInExcel;
									master.addField("CompletedDate", getDate(Closed));
								}
							}
							if (cn == 4) {
								component = row.getCell(cn).getStringCellValue();
								master.addField("ItemName", component);
							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								createdOn = getweekDate() + valueAsInExcel;
								master.addField("CreatedDate", getDate(createdOn));
							}
							if (cn == 6) {

								ShortDescription = row.getCell(cn).getStringCellValue();
								String shortDescription = Jsoup.parse(ShortDescription.toString()).text();
								shortDescription = shortDescription.toString().replaceAll("(\r\n|\n)", " ");
								shortDescription = CleanInvalidXmlChars(shortDescription.toString(), "");

								master.addField("Description", shortDescription);
							}
							if (cn == 7) {
								Incidentstate = row.getCell(cn).getStringCellValue();
								master.addField("Status", Incidentstate);
							}
							if (cn == 8) {
								productCategory = row.getCell(cn).getStringCellValue();
								master.addField("ProductCategory", productCategory);
							}
							if (cn == 9) {
								number = (int) row.getCell(cn).getNumericCellValue();
								master.addField("IncidentID", number);
							}
							if (cn == 10) {
								OnBehalfOf = row.getCell(cn).getStringCellValue();
								if (!OnBehalfOf.isEmpty() && OnBehalfOf != null) {
									
									master.addField("FirstName", OnBehalfOf);
								}
							}
							if (cn == 11) {
								emailId = row.getCell(cn).getStringCellValue();
								master.addField("EmailId", emailId);
							}
							if (cn == 12) {
								Priority = row.getCell(cn).getStringCellValue();
								master.addField("Priority", Priority);
							}
							if (cn == 13) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Resolved = getweekDate() + valueAsInExcel;
									master.addField("ResolvedDate", getDate(Resolved));
								}
							}

							if (cn == 14) {
								desc = row.getCell(cn).getStringCellValue();
								String descr = Jsoup.parse(desc.toString()).text();
								descr = descr.toString().replaceAll("(\r\n|\n)", " ");
								descr = CleanInvalidXmlChars(descr.toString(), "");

								master.addField("Title", descr);
							}

							if (cn == 15) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = getweekDate() + valueAsInExcel;
									master.addField("LastModifiedDate", getDate(Updated));
								}
							}
							if (cn == 16) {

								notes = row.getCell(cn).getStringCellValue();
								if (!notes.isEmpty() && notes != null) {
								if (String.valueOf(notes).contains("Additional comments")) {

									String comment = String.valueOf(notes)
											.replace("(Additional comments (Customer Visible))", "~");
									String[] comments = String.valueOf(comment).split("~");
									String notesCreatedByDetails = comments[0];

									String[] NotesCreated = String.valueOf(notesCreatedByDetails).split(" -");
									String notesCreatedDate = NotesCreated[0];
									String notesCreatedBy = NotesCreated[1];

									String note = comments[1];

									note = note.toString().replaceAll("(\r\n|\n)", " ");
									note = CleanInvalidXmlChars(note.toString(), "");

									String actualNotesCreatedBy = notesCreatedByDetails.substring(22);

									if (!notes.isEmpty() && notes != null) {
										master.addField("Note", note);
									}
									if (!notesCreatedBy.isEmpty() && notesCreatedBy != null) {
										master.addField("NotesCreatedBy", actualNotesCreatedBy);
									}
									if (!notesCreatedDate.isEmpty() && notesCreatedDate != null) {
										master.addField("NotesCreatedDate", master.getFieldValue("LastModifiedDate"));
									}

								}
								}
							}
							if (cn == 17) {
								closenotescomments = row.getCell(cn).getStringCellValue();
								String Closenotescomments = Jsoup.parse(closenotescomments.toString()).text();
								Closenotescomments = Closenotescomments.toString().replaceAll("(\r\n|\n)", " ");
								Closenotescomments = CleanInvalidXmlChars(Closenotescomments.toString(), "");
								master.addField("ClosureComments", Closenotescomments);

							}
							if (cn == 18) {
								updatedBy = row.getCell(cn).getStringCellValue();
								if (!updatedBy.isEmpty() && updatedBy != null) {
									master.addField("UpdatedBy", updatedBy);
								}
							}
							if (cn == 19) {
								ETC_PREDICT_VAL = (float) row.getCell(cn).getNumericCellValue();
								master.addField("ETC_PREDICT_VALUE", ETC_PREDICT_VAL);

							}
							if (cn == 20) {
								overallSentiment = row.getCell(cn).getStringCellValue();
								if (!overallSentiment.isEmpty() && overallSentiment != null) {
									master.addField("OverallSentiment", overallSentiment);
								}

							}
							if (cn == 21) {
								surveyTaken = (int) row.getCell(cn).getNumericCellValue();
								if (surveyTaken == 0 || surveyTaken == 1) {
									master.addField("SurveyTaken", surveyTaken);
								}

							}
							if (cn == 22) {
								csatAwaitingActionSentiment = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionSentiment == 0 || csatAwaitingActionSentiment == 1) {
									master.addField("CSATAwaitingActionSentiment", csatAwaitingActionSentiment);
								}

							}
							if (cn == 23) {
								csatAwaitingActionETC = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionETC == 0 || csatAwaitingActionETC == 1) {
									master.addField("CSATAwaitingActionETC", csatAwaitingActionETC);
								}

							}
							if (cn == 24) {
								responseSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (responseSlaBreach == 0 || responseSlaBreach == 1 || responseSlaBreach == 2) {
									master.addField("ResponseSlaBreach", responseSlaBreach);
								}

							}
							if (cn == 25) {
								resolutionSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (resolutionSlaBreach == 0 || resolutionSlaBreach == 1 || resolutionSlaBreach == 2) {
									master.addField("ResolutionSlaBreach", resolutionSlaBreach);
								}

							}
							if (cn == 26) {
								suggestedVirtualSupervisior = row.getCell(cn).getStringCellValue();
								if (!suggestedVirtualSupervisior.isEmpty() && suggestedVirtualSupervisior != null) {
									master.addField("SuggestedVirtualSupervisior", suggestedVirtualSupervisior);
								}

							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = updatedDate + valueAsInExcel;
									master.addField("ReportedTime", getDate(Updated));
								}
							}
							if (cn == 28) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("Level1", level1);
								}

							}
							if (cn == 29) {
								String level2 = row.getCell(cn).getStringCellValue();
								if (!level2.isEmpty() && level2 != null) {
									master.addField("Level2", level2);
								}

							}
							if (cn == 30) {
								String level3 = row.getCell(cn).getStringCellValue();
								if (!level3.isEmpty() && level3 != null) {
									master.addField("Level3", level3);
								}

							}

						if (cn == 31) {
							String kwar = row.getCell(cn).getStringCellValue();
							if (!kwar.isEmpty() && kwar != null) {
								master.addField("TvmKeywordVar", kwar);
							}
						}
						if (cn == 32) {
							int tvm = (int) row.getCell(cn).getNumericCellValue();
							if (tvm == 1 || tvm == 2 || tvm == 3) {
								master.addField("TvmPredictMinValue", tvm);
							}
						}
						if (cn == 33) {
							int tvm = (int) row.getCell(cn).getNumericCellValue();
							if (tvm == 1 || tvm == 2 || tvm == 3) {
								master.addField("TvmPredictMaxValue", tvm);
							}
						}

					}
				}
					if (row.getRowNum() > 150) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								Assignedto = row.getCell(cn).getStringCellValue();
								if (!Assignedto.isEmpty() && Assignedto != null) {
									master.addField("Assignee", Assignedto);

								}
							}
							if (cn == 1) {
								Assignmentgroup = row.getCell(cn).getStringCellValue();
								master.addField("AssignedGroup", Assignmentgroup);
							}
							if (cn == 2) {
								Category = row.getCell(cn).getStringCellValue();
								master.addField("SubCategory", Category);
							}
							if (cn == 3) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Closed = getweekDate() + valueAsInExcel;
									master.addField("CompletedDate", getDate(Closed));
								}
							}
							if (cn == 4) {
								component = row.getCell(cn).getStringCellValue();
								master.addField("ItemName", component);
							}
							if (cn == 5) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								createdOn = getResolvedDate() + valueAsInExcel;
								master.addField("CreatedDate", getDate(createdOn));
							}
							if (cn == 6) {

								ShortDescription = row.getCell(cn).getStringCellValue();
								String shortDescription = Jsoup.parse(ShortDescription.toString()).text();
								shortDescription = shortDescription.toString().replaceAll("(\r\n|\n)", " ");
								shortDescription = CleanInvalidXmlChars(shortDescription.toString(), "");

								master.addField("Description", shortDescription);
							}
							if (cn == 7) {
								Incidentstate = row.getCell(cn).getStringCellValue();
								master.addField("Status", Incidentstate);
							}
							if (cn == 8) {
								productCategory = row.getCell(cn).getStringCellValue();
								master.addField("ProductCategory", productCategory);
							}
							if (cn == 9) {
								number = (int) row.getCell(cn).getNumericCellValue();
								master.addField("IncidentID", number);
							}
							if (cn == 10) {
								OnBehalfOf = row.getCell(cn).getStringCellValue();
								if (!OnBehalfOf.isEmpty() && OnBehalfOf != null) {
							
									master.addField("FirstName", OnBehalfOf);
								}
							}
							if (cn == 11) {
								emailId = row.getCell(cn).getStringCellValue();
								master.addField("EmailId", emailId);
							}
							if (cn == 12) {
								Priority = row.getCell(cn).getStringCellValue();
								master.addField("Priority", Priority);
							}
							if (cn == 13) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Resolved = getResolvedDate() + valueAsInExcel;
									master.addField("ResolvedDate", getDate(Resolved));
								}
							}

							if (cn == 14) {
								desc = row.getCell(cn).getStringCellValue();
								String descr = Jsoup.parse(desc.toString()).text();
								descr = descr.toString().replaceAll("(\r\n|\n)", " ");
								descr = CleanInvalidXmlChars(descr.toString(), "");

								master.addField("Title", descr);
							}

							if (cn == 15) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = getResolvedDate() + valueAsInExcel;
									master.addField("LastModifiedDate", getDate(Updated));
								}
							}
							if (cn == 16) {

								notes = row.getCell(cn).getStringCellValue();
								if (!notes.isEmpty() && notes != null) {
								if (String.valueOf(notes).contains("Additional comments")) {

									String comment = String.valueOf(notes)
											.replace("(Additional comments (Customer Visible))", "~");
									String[] comments = String.valueOf(comment).split("~");
									String notesCreatedByDetails = comments[0];

									String[] NotesCreated = String.valueOf(notesCreatedByDetails).split(" -");
									String notesCreatedDate = NotesCreated[0];
									String notesCreatedBy = NotesCreated[1];

									String note = comments[1];

									note = note.toString().replaceAll("(\r\n|\n)", " ");
									note = CleanInvalidXmlChars(note.toString(), "");

									String actualNotesCreatedBy = notesCreatedByDetails.substring(22);

									if (!notes.isEmpty() && notes != null) {
										master.addField("Note", note);
									}
									if (!notesCreatedBy.isEmpty() && notesCreatedBy != null) {
										master.addField("NotesCreatedBy", actualNotesCreatedBy);
									}
									if (!notesCreatedDate.isEmpty() && notesCreatedDate != null) {
										master.addField("NotesCreatedDate", master.getFieldValue("LastModifiedDate"));
									}

								}
								}
							}
							if (cn == 17) {
								closenotescomments = row.getCell(cn).getStringCellValue();
								String Closenotescomments = Jsoup.parse(closenotescomments.toString()).text();
								Closenotescomments = Closenotescomments.toString().replaceAll("(\r\n|\n)", " ");
								Closenotescomments = CleanInvalidXmlChars(Closenotescomments.toString(), "");
								master.addField("ClosureComments", Closenotescomments);

							}
							if (cn == 18) {
								updatedBy = row.getCell(cn).getStringCellValue();
								if (!updatedBy.isEmpty() && updatedBy != null) {
									master.addField("UpdatedBy", updatedBy);
								}
							}
							if (cn == 19) {
								ETC_PREDICT_VAL = (float) row.getCell(cn).getNumericCellValue();
								master.addField("ETC_PREDICT_VALUE", ETC_PREDICT_VAL);

							}
							if (cn == 20) {
								overallSentiment = row.getCell(cn).getStringCellValue();
								if (!overallSentiment.isEmpty() && overallSentiment != null) {
									master.addField("OverallSentiment", overallSentiment);
								}

							}
							if (cn == 21) {
								surveyTaken = (int) row.getCell(cn).getNumericCellValue();
								if (surveyTaken == 0 || surveyTaken == 1) {
									master.addField("SurveyTaken", surveyTaken);
								}

							}
							if (cn == 22) {
								csatAwaitingActionSentiment = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionSentiment == 0 || csatAwaitingActionSentiment == 1) {
									master.addField("CSATAwaitingActionSentiment", csatAwaitingActionSentiment);
								}

							}
							if (cn == 23) {
								csatAwaitingActionETC = (int) row.getCell(cn).getNumericCellValue();
								if (csatAwaitingActionETC == 0 || csatAwaitingActionETC == 1) {
									master.addField("CSATAwaitingActionETC", csatAwaitingActionETC);
								}

							}
							if (cn == 24) {
								responseSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (responseSlaBreach == 0 || responseSlaBreach == 1 || responseSlaBreach == 2) {
									master.addField("ResponseSlaBreach", responseSlaBreach);
								}

							}
							if (cn == 25) {
								resolutionSlaBreach = (int) row.getCell(cn).getNumericCellValue();
								if (resolutionSlaBreach == 0 || resolutionSlaBreach == 1 || resolutionSlaBreach == 2) {
									master.addField("ResolutionSlaBreach", resolutionSlaBreach);
								}

							}
							if (cn == 26) {
								suggestedVirtualSupervisior = row.getCell(cn).getStringCellValue();
								if (!suggestedVirtualSupervisior.isEmpty() && suggestedVirtualSupervisior != null) {
									master.addField("SuggestedVirtualSupervisior", suggestedVirtualSupervisior);
								}

							}
							if (cn == 27) {
								DataFormatter fmt = new DataFormatter();
								String valueAsInExcel = fmt.formatCellValue(row.getCell(cn));
								if (!valueAsInExcel.isEmpty() && valueAsInExcel != null) {
									Updated = getResolvedDate() + valueAsInExcel;
									master.addField("ReportedTime", getDate(Updated));
								}
							}
							if (cn == 28) {
								String level1 = row.getCell(cn).getStringCellValue();
								if (!level1.isEmpty() && level1 != null) {
									master.addField("Level1", level1);
								}

							}
							if (cn == 29) {
								String level2 = row.getCell(cn).getStringCellValue();
								if (!level2.isEmpty() && level2 != null) {
									master.addField("Level2", level2);
								}

							}
							if (cn == 30) {
								String level3 = row.getCell(cn).getStringCellValue();
								if (!level3.isEmpty() && level3 != null) {
									master.addField("Level3", level3);
								}

							}

						
						if (cn == 31) {
							String level1 = row.getCell(cn).getStringCellValue();
							if (!level1.isEmpty() && level1 != null) {
								master.addField("TvmKeywordVar", level1);
							}
						}
						if (cn == 32) {
							int tvm = (int) row.getCell(cn).getNumericCellValue();
							if (tvm == 1 || tvm == 2 || tvm == 3) {
								master.addField("TvmPredictMinValue", tvm);
							}
						}
						if (cn == 33) {
							int tvm = (int) row.getCell(cn).getNumericCellValue();
							if (tvm == 1 || tvm == 2 || tvm == 3) {
								master.addField("TvmPredictMaxValue", tvm);
							}
						}
						}
					}
				
				}
				
				// }
				//System.out.println(master);
				if (!master.isEmpty()) {

					if (master.getFieldValue("CreatedDate") != null
							&& !master.getFieldValue("CreatedDate").toString().isEmpty()) {
						key = String.valueOf(master.getFieldValue("IncidentID"));
						// String cDate =
						// getDate(""+master.getFieldValue("CreatedDate"));

						// master.removeField("LastModifiedDate");
						// master.removeField("CreatedDate");

						// master.addField("CreatedDate",cDate);

						// master.addField("LastModifiedDate",cDate);

						// String lDate= getDate(""+
						// master.getFieldValue("LastModifiedDate"));
						SolrInputDocument submited = new SolrInputDocument();
						submited = master.deepCopy();
						submited.removeField("SuggestedVirtualSupervisior");
						submited.removeField("ResponseSlaBreach");
						submited.removeField("ResolutionSlaBreach");
						submited.removeField("CSATAwaitingActionSentiment");
						submited.removeField("CSATAwaitingActionETC");
						submited.removeField("OverallSentiment");
						submited.removeField("SurveyTaken");
						submited.removeField("Status");
						submited.removeField("Level1");
						submited.removeField("Level2");
						submited.removeField("Level3");
						submited.removeField("TvmKeywordVar");
						submited.removeField("TvmPredictMinValue");
						submited.removeField("TvmPredictMaxValue");
						submited.removeField("ClosureComments");
						submited.addField("Status", "Submitted");
						submited.removeField("CompletedDate");
						submited.removeField("LastModifiedDate");
						submited.removeField("ResolvedDate");
						submited.addField("LastModifiedDate", master.getFieldValue("CreatedDate"));

						String priority = (String) submited.getFieldValue("Priority");
						submited.removeField("Priority");
						submited.addField("Priority", Priority);
						// System.out.println(master.getFieldValue("LastModifiedDate"));
						submited.addField("IncidentID-LastModifiedDate",
								master.getFieldValue("IncidentID") + "-" + master.getFieldValue("CreatedDate"));
						server.add(submited);
						submited.clear();
					}

					String checkStatus = master.getFieldValue("Status").toString();

					if (checkStatus.equalsIgnoreCase("Work In Progress")) {
						if (master.getFieldValue("CreatedDate") != null
								&& !master.getFieldValue("CreatedDate").toString().isEmpty()) {

							/*
							 * //String cDate =
							 * getDate(""+master.getFieldValue("CreatedDate"));
							 * master.removeField("LastModifiedDate");
							 * //master.removeField("CreatedDate"); String
							 * lmodifiedDate =""; // lmodifiedDate =
							 * getDate(json.getString("sys_updated_on"));
							 * 
							 * lmodifiedDate = getDate(Updated);
							 * master.addField("LastModifiedDate",lmodifiedDate)
							 * ;
							 */

							SolrInputDocument submited = new SolrInputDocument();
							submited = master.deepCopy();
							submited.removeField("SuggestedVirtualSupervisior");
							submited.removeField("ResponseSlaBreach");
							submited.removeField("ResolutionSlaBreach");
							submited.removeField("CSATAwaitingActionSentiment");
							submited.removeField("CSATAwaitingActionETC");
							submited.removeField("OverallSentiment");
							submited.removeField("SurveyTaken");
							submited.removeField("Status");
							submited.removeField("Level1");
							submited.removeField("Level2");
							submited.removeField("Level3");
							submited.removeField("TvmKeywordVar");
							submited.removeField("TvmPredictMinValue");
							submited.removeField("TvmPredictMaxValue");
							submited.removeField("ClosureComments");
							submited.addField("Status", "InProgress");
							submited.removeField("CompletedDate");
							submited.removeField("ResolvedDate");
							String priority = (String) submited.getFieldValue("Priority");
							submited.removeField("Priority");
							submited.addField("Priority", Priority);
							// lmodifiedDate=getDate(Updated);
							submited.addField("IncidentID-LastModifiedDate", master.getFieldValue("IncidentID") + "-"
									+ master.getFieldValue("LastModifiedDate"));
							server.add(submited);
							submited.clear();
						}

					}

					// pending
					if (checkStatus.equalsIgnoreCase("Awaiting User Info")
							|| checkStatus.equalsIgnoreCase("Awaiting Problem")
							|| checkStatus.equalsIgnoreCase("Awaiting Problem Resolution")) {

						if (master.getFieldValue("CreatedDate") != null
								&& !master.getFieldValue("CreatedDate").toString().isEmpty()) {

							// master.removeField("LastModifiedDate");
							// master.removeField("CreatedDate");

							// master.addField("CreatedDate",master.getFieldValue("CreatedDate"));
							/*
							 * String lmodifiedDate =""; lmodifiedDate =
							 * getDate(Updated);
							 */
							// master.addField("LastModifiedDate",master.getFieldValue(name));

							SolrInputDocument submited = new SolrInputDocument();
							submited = master.deepCopy();
							submited.removeField("SuggestedVirtualSupervisior");
							submited.removeField("ResponseSlaBreach");
							submited.removeField("ResolutionSlaBreach");
							submited.removeField("CSATAwaitingActionSentiment");
							submited.removeField("CSATAwaitingActionETC");
							submited.removeField("OverallSentiment");
							submited.removeField("SurveyTaken");
							submited.removeField("Status");
							submited.removeField("Level1");
							submited.removeField("Level2");
							submited.removeField("Level3");
							submited.removeField("TvmKeywordVar");
							submited.removeField("TvmPredictMinValue");
							submited.removeField("TvmPredictMaxValue");
							submited.removeField("ClosureComments");
							submited.addField("Status", "Pending");
							submited.removeField("CompletedDate");
							submited.removeField("ResolvedDate");
							String priority = (String) submited.getFieldValue("Priority");
							submited.removeField("Priority");
							submited.addField("Priority", Priority);

							submited.addField("IncidentID-LastModifiedDate", master.getFieldValue("IncidentID") + "-"
									+ master.getFieldValue("LastModifiedDate"));
							server.add(submited);
							submited.clear();
						}
					}

					// Resolved
					if (checkStatus.equalsIgnoreCase("Resolved") || checkStatus.equalsIgnoreCase("Closed")
							|| checkStatus.equalsIgnoreCase("Closed Complete")) {
						// master.removeField("LastModifiedDate");

						/*
						 * String lmodifiedDate =""; lmodifiedDate =
						 * getDate(Updated);
						 * master.addField("LastModifiedDate",lmodifiedDate);
						 */

						SolrInputDocument submited = new SolrInputDocument();
						submited = master.deepCopy();
						submited.removeField("SuggestedVirtualSupervisior");
						submited.removeField("ResponseSlaBreach");
						submited.removeField("ResolutionSlaBreach");
						submited.removeField("CSATAwaitingActionSentiment");
						submited.removeField("CSATAwaitingActionETC");
						submited.removeField("OverallSentiment");
						submited.removeField("SurveyTaken");
						submited.removeField("Status");
						submited.removeField("Level1");
						submited.removeField("Level2");
						submited.removeField("Level3");
						submited.removeField("TvmKeywordVar");
						submited.removeField("TvmPredictMinValue");
						submited.removeField("TvmPredictMaxValue");
						submited.removeField("CompletedDate");
						submited.removeField("LastModifiedDate");
						submited.addField("LastModifiedDate", master.getFieldValue("ResolvedDate"));
						// submited.removeField("ResolvedDate");
						submited.addField("Status", "Closed");

						/*
						 * if(master.getFieldValue("ResolvedDate")!=null &&
						 * !master.getFieldValue("ResolvedDate").toString().
						 * isEmpty()) { submited.addField("ResolvedDate",
						 * getDate(master.getFieldValue("ResolvedDate").toString
						 * ())); }
						 * if(master.getFieldValue("CompletedDate")!=null &&
						 * !master.getFieldValue("CompletedDate").toString().
						 * isEmpty()) {
						 * master.addField("CompletedDate",getDate(master.
						 * getFieldValue("CompletedDate").toString())); }
						 */

						String priority = (String) submited.getFieldValue("Priority");
						submited.removeField("Priority");
						submited.addField("Priority", Priority);
						submited.removeField("IncidentID-LastModifiedDate");
						submited.addField("IncidentID-LastModifiedDate",
								master.getFieldValue("IncidentID") + "-" + master.getFieldValue("ResolvedDate"));
						server.add(submited);
						submited.clear();
					} else {
						master.removeField("ResolvedDate");
					}

					// master.removeField("IncidentID-LastModifiedDate");

					String masterSstatus = master.getFieldValue("Status").toString();
					if (masterSstatus.equalsIgnoreCase("Closed") || checkStatus.equalsIgnoreCase("Closed Complete")) {

						// String rDate =
						// getDate(""+master.getFieldValue("ResolvedDate"));
						// String cDate =
						// getDate(""+master.getFieldValue("CompletedDate"));
						// master.removeField("ResolvedDate");
						// master.removeField("CompletedDate");
						// master.addField("CompletedDate", cDate);
						// master.addField("ResolvedDate", rDate);
					} /*
						 * else if (masterSstatus.equalsIgnoreCase("Resolved"))
						 * { //String rDate =
						 * getDate(""+master.getFieldValue("ResolvedDate"));
						 * //master.removeField("ResolvedDate");
						 * master.removeField("CompletedDate");
						 * //master.addField("ResolvedDate", rDate); }
						 */ else if (checkStatus.equalsIgnoreCase("Active")) {
						master.removeField("ResolvedDate");
						master.removeField("CompletedDate");
						master.removeField("Status");
						master.addField("Status", "Submitted");
					} else if (checkStatus.equalsIgnoreCase("Awaiting User Info")
							|| checkStatus.equalsIgnoreCase("Awaiting Problem")
							|| checkStatus.equalsIgnoreCase("Awaiting Problem Resolution")) {
						master.removeField("ResolvedDate");
						master.removeField("CompletedDate");
						master.removeField("Status");
						master.addField("Status", "Pending");

						// System.out.println(master.toString());

					} else if (checkStatus.equalsIgnoreCase("Work In Progress")) {
						master.removeField("ClosureComments");
						master.removeField("ResolvedDate");
						master.removeField("CompletedDate");
						master.removeField("Status");
						master.addField("Status", "InProgress");

					} else if (checkStatus.equalsIgnoreCase("Cancelled")) {
						// System.out.println("Cancelled: "+checkStatus);
						continue;
					}

					String priority = (String) master.getFieldValue("Priority");
					master.removeField("Priority");
					master.addField("Priority", Priority);

					if (!master.isEmpty()) {
						serverMaster.add(master);
					}
					serverMaster.commit();
					server.commit();
				}
				// System.out.println("master: "+master);
				count++;
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}
		System.out.println("Tickets data loaded sucessfully");
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

	public static String getResolvedDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, 0);
		return dateFormat.format(cal.getTime());
	}

	public static String getUpdatedDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -1);
		return dateFormat.format(cal.getTime());
	}

	public static String getCreatedDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -2);
		return dateFormat.format(cal.getTime());
	}

	public static String getMonthDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -35);
		return dateFormat.format(cal.getTime());
	}

	public static String getweekDate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -5);
		return dateFormat.format(cal.getTime());
	}

	public static String get90To180Dates() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -99);
		return dateFormat.format(cal.getTime());
	}

	public static String get180To360Dates() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -195);
		return dateFormat.format(cal.getTime());
	}

	public static String get30To90Dates() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd ");
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, -40);
		return dateFormat.format(cal.getTime());
	}

	public static String CleanInvalidXmlChars(String text, String replacement) {
		String re = "[^\\x09\\x0A\\x0D\\x20-\\xD7FF\\xE000-\\xFFFD\\x10000-x10FFFF]";
		return text.replaceAll(re, replacement);
	}

	public static void deleteSolrData() {
		HttpSolrServer solr = new HttpSolrServer(master);
		HttpSolrServer solr2 = new HttpSolrServer(CI_APPLICATIONS);
		try {
			solr.deleteByQuery("*:*");
			solr2.deleteByQuery("*:*");
		} catch (SolrServerException e) {
			throw new RuntimeException("Failed to delete data in Solr. " + e.getMessage(), e);
		} catch (IOException e) {
			throw new RuntimeException("Failed to delete data in Solr. " + e.getMessage(), e);
		}
	}
}
