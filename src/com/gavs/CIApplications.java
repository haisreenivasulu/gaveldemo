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

public class CIApplications {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-ci-applications";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"Applications.xlsx";
		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 1;
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			String Department = "";
			String ApplicationName = "";
			String Description = "";
			int DependsOnDeviceID = 0;
			int ApplicationID = 0;
			String Location = "";
			String BusinessImpact = "";
			String DependsOnHost = "";
			String DeviceName = "";
			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 0) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								Department = row.getCell(cn).getStringCellValue();
								master.addField("Department", Department);
							}
							if (cn == 1) {
								ApplicationName = row.getCell(cn).getStringCellValue();
								master.addField("ApplicationName", ApplicationName);
							}
							if (cn == 2) {
								Description = row.getCell(cn).getStringCellValue();
								master.addField("Description", Description);
							}
							if (cn == 3) {
								BusinessImpact = row.getCell(cn).getStringCellValue();
								master.addField("BusinessImpact", BusinessImpact);
							}
							if (cn == 4) {
								DependsOnDeviceID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DependsOnDeviceID", DependsOnDeviceID);
							}
							if (cn == 5) {
								DependsOnHost = row.getCell(cn).getStringCellValue();
								master.addField("DependsOnHost", DependsOnHost);
							}
							if (cn == 6) {
								ApplicationID = (int) row.getCell(cn).getNumericCellValue();
								master.addField("ApplicationID", ApplicationID);
							}
							if (cn == 7) {
								Location = row.getCell(cn).getStringCellValue();
								master.addField("Location", Location);
							}
						}

					}

				}
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
		System.out.println("Applications data loaded successfully");
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
