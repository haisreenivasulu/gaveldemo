package com.gavs;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

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
public class MemoryPrediction {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-prediction-memory";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"Memory_Prediction.xlsx";

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
								ptime = (int) row.getCell(cn).getNumericCellValue();
								master.addField("Index", ptime);

							}
							if (cn == 3) {
								multiAssignedGroup = row.getCell(cn).getStringCellValue();
								master.addField("AtRisk", multiAssignedGroup);

							}
							if (cn == 4) {
								// deviceValue=row.getCell(cn).getStringCellValue();
								master.addField("DeviceID-Index",
										master.getFieldValue("DeviceID") + "||" + master.getFieldValue("Index"));
							}
							if (cn == 5) {
								pvalue = row.getCell(cn).getStringCellValue();
								master.addField("Alert", pvalue);
							}
							if (cn == 6) {
								deviceValue = row.getCell(cn).getStringCellValue();
								master.addField("NumberOfPredictions", deviceValue);
							}

							if (cn == 7) {
								ldate = (int) row.getCell(cn).getNumericCellValue();
								master.addField("LatestDateTime", ldate);
							}
						}

					}

				}
				if (!master.isEmpty()) {
					serverMaster.add(master);
				}
				serverMaster.commit();
				// System.out.println("after: "+master);

				count++;
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}
		System.out.println("MemoryPrediction count");
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
// }
