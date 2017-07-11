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

public class CIAssetsImport {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-ci-assets";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"Assests.xlsx";
		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 1;
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			String deviceType = "";
			String description = "";
			String businessImpact = "";
			int id = 0;
			String IPAddress = "";
			String DNSName = "";
			String DeviceName = "";
			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 0) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								deviceType = row.getCell(cn).getStringCellValue();
								master.addField("DeviceType", deviceType);
							}
							if (cn == 1) {
								description = row.getCell(cn).getStringCellValue();
								master.addField("Description", description);
							}
							if (cn == 2) {
								businessImpact = row.getCell(cn).getStringCellValue();
								master.addField("BusinessImpact", businessImpact);
							}
							if (cn == 3) {
								id = (int) row.getCell(cn).getNumericCellValue();
								master.addField("DeviceID", id);
							}
							if (cn == 4) {
								IPAddress = row.getCell(cn).getStringCellValue();
								master.addField("IPAddress", IPAddress);
							}
							if (cn == 5) {
								DNSName = row.getCell(cn).getStringCellValue();
								master.addField("DNSName", DNSName);
							}
							if (cn == 6) {
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
				// System.out.println("after: " + master);
				count++;
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}
		System.out.println("Assets data loaded successfully ");
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
