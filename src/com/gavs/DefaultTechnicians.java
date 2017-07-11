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

public class DefaultTechnicians {

	private static final String master = GavelConstant.SOLR_URL +""+GavelConstant.GAVEL_COMPANY +"-technicians-manageengine";
	private static String key = "";

	public static void uploadSolrData() throws IOException {
		SolrServer serverMaster = new HttpSolrServer(master);
		deleteSolrData();
		String fileName = GavelConstant.FILE_PATH+"Default_Technicians.xlsx";

		File file = new File(fileName);
		FileInputStream fIP = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fIP);
		int count = 1;
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (Row row : sheet) {
			String OnBehalfOf = "";
			int id = 0;
			String multiAssignedGroup = "";
			String assignedGroup = "";

			SolrInputDocument master = new SolrInputDocument();
			try {
				for (int cn = 0; cn <= row.getLastCellNum(); cn++) {
					if (row.getRowNum() > 0) {

						if (row.getCell(cn) != null) {

							if (cn == 0) {
								OnBehalfOf = row.getCell(cn).getStringCellValue();
								if (!OnBehalfOf.isEmpty() && OnBehalfOf != null) {
									String custarray[] = OnBehalfOf.split(" ");
									master.addField("TechnicianName", OnBehalfOf);

									String cutEmail = OnBehalfOf.replaceAll("\\s+", "");

									cutEmail = cutEmail + "@gavstech.com";

									master.addField("TechnicianEmailID", cutEmail + "");
								}
							}
							if (cn == 1) {
								id = (int) row.getCell(cn).getNumericCellValue();
								master.addField("TechnicianID", id);

							}
							if (cn == 2) {
								multiAssignedGroup = row.getCell(cn).getStringCellValue();
								master.addField("MultiAssignedGroup", multiAssignedGroup);

							}
							if (cn == 3) {
								assignedGroup = row.getCell(cn).getStringCellValue();
								master.addField("AssignedGroup", assignedGroup);
							}

						}

					}

				}
				if (!master.isEmpty()) {
					serverMaster.add(master);
				}
				serverMaster.commit();

				count++;
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}
		System.out.println("Technicians data loaded successfully");
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
