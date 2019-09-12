package bcbl.officiel.convocation;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ConvocationsOfficiel {
	public static void main(String[] args) {

		SimpleDateFormat dateFormat = new SimpleDateFormat();
		dateFormat.applyPattern("dd MMM yyyy");

		SimpleDateFormat heureFormat = new SimpleDateFormat();
		heureFormat.applyPattern("HH:mm");

		String extract = null;
		String output = null;
		String prolog = null;
		boolean overwrite = false;

		for (int i = 0; i < args.length; i++) {
			if ("-extract".equals(args[i])) {
				extract = args[++i];
			} else if ("-output".equals(args[i])) {
				output = args[++i];
			} else if ("-prolog".equals(args[i])) {
				prolog = args[++i];
			} else if ("-overwrite".equals(args[i])) {
				overwrite = true;
			}
		}

		if (extract == null) {
			System.err.println("Fichier d'extraction non spécifié");
			System.exit(1);
		}
		if (output == null) {
			System.err.println("Fichier de sortie non spécifié");
			System.exit(1);
		}

		try {
			HSSFWorkbook extractFbiWb = new HSSFWorkbook(new FileInputStream(extract));
			HSSFSheet extractFbiSheet = extractFbiWb.getSheetAt(0);

			BufferedWriter writer;
			File fOutput = new File(output);
			if (!overwrite && fOutput.exists()) {
				extractFbiWb.close();
				System.err.println("Fichier de sortie non spécifié");
				System.exit(1);
				return;
			} else {
				writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(fOutput)));
			}

			writer.write("<html><meta charset=\"UTF-8\"><body>");
			writer.newLine();

			if (prolog != null) {
				BufferedReader br = new BufferedReader(new FileReader(prolog));
				String line = br.readLine();
				while (line != null) {
					writer.write(line);
					writer.newLine();
					line = br.readLine();
				}
			}

			int rows = extractFbiSheet.getPhysicalNumberOfRows();
			
			System.out.println("Number of rows:" + rows);

			writer.write("<table>");
			writer.newLine();
			writer.write(
					"<thead><th>Arbitre</th><th>Date</th><th>Heure</th><th>Catégorie</th><th>Match (en gras - équipe recevant)</th></thead>");
			writer.newLine();
			writer.write("<tbody>");
			writer.newLine();

			String nom = "";
			String prenom = "";
			String fonction = "";
			String categorie = "";
			for (int rowIndex = 8; rowIndex < rows; rowIndex++) {
				HSSFRow row = extractFbiSheet.getRow(rowIndex);

				if (row != null && row.getPhysicalNumberOfCells() > 0) {
					System.out.println("Row " + rowIndex + " - number of cells : " + row.getPhysicalNumberOfCells());
					String s = null;
					try {
						s = row.getCell(2).getStringCellValue();
					} catch (Exception e) {

					}
					if (s != null && !s.isEmpty()) {
						fonction = s;
					}
					// On s'assure qu'il s'agit bien d'un arbitre officiel

					if ("ARBITRE".equalsIgnoreCase(fonction)) {

						s = row.getCell(0).getStringCellValue();
						if (s != null && !s.isEmpty()) {
							nom = s;
						}
						s = row.getCell(1).getStringCellValue();
						if (s != null && !s.isEmpty()) {
							prenom = s;
						}

						String equipe1 = normalizeNomEquipe(row.getCell(7).getStringCellValue());
						String equipe2 = normalizeNomEquipe(row.getCell(8).getStringCellValue());
						if (!row.getCell(5).getStringCellValue().trim().isEmpty()) { 
							categorie = row.getCell(5).getStringCellValue();
						}

						Date date = row.getCell(13).getDateCellValue();
						double heure = row.getCell(14).getNumericCellValue();
						date.setHours((int) heure / 100);
						date.setMinutes((int) heure % 100);

						writer.write("<tr>");
						String[] values = new String[] { nom + " " + prenom, dateFormat.format(date),
								heureFormat.format(date), categorie, "<b>" + equipe1 + "</b> vs " + equipe2 };
						for (String value : values) {
							writer.write("<td>" + value + "</td>");
						}
						writer.write("</tr>");
						writer.newLine();
					}
				}
			}

			writer.write("</tbody>");
			writer.newLine();
			writer.write("</table>");
			writer.newLine();
			writer.write("</body></html>");

			writer.close();

			extractFbiWb.close();

		} catch (IOException ioe) {
			ioe.printStackTrace();
		}

	}
	
	public static String normalizeNomEquipe(String equipe) {
		// Supprimer les - 1 ou - 2 en fin de nom d'équipe
		int idx = equipe.lastIndexOf('-');
		if (idx > 0) {
			return equipe.substring(0, idx);
		}
		return equipe;
	}

}
