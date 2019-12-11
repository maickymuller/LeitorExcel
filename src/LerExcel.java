import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//@SuppressWarnings("unused")
public class LerExcel {

	public static void main(String[] args) throws IOException {
		JFileChooser escolhedor = new JFileChooser();

		escolhedor.setFileFilter(new FileNameExtensionFilter("Microsoft Excel Files", "xls", "xlsx"));
		escolhedor.setAcceptAllFileFilterUsed(false);

		int opcaoEscolhida = escolhedor.showOpenDialog(null); // Janela para abrir um arquivo

		if (opcaoEscolhida == JFileChooser.APPROVE_OPTION) {
			// Código pra caso o usuário tenha escolhido abrir um arquivo

//        	escolhedor.getSelectedFile().getName()

			// Mosta o nome do arquivo Excel lido
			System.out.println(escolhedor.getSelectedFile().getName());

//          File excelFile = new File("contacts.xlsx");
			File excelFile = new File(escolhedor.getSelectedFile().getName());

			FileInputStream fis = new FileInputStream(excelFile);

			// we create an XSSF Workbook object for our XLSX Excel File
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			// we get first sheet
			XSSFSheet sheet = workbook.getSheetAt(0);

			// we iterate on rows
			Iterator<Row> rowIt = sheet.iterator();

			while (rowIt.hasNext()) {
				Row row = rowIt.next();

				// iterate on cells for the current row
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					System.out.print(cell.toString() + ";");
				}

				System.out.println();
			}

			workbook.close();
			fis.close();
		}

	}
}

//public class ReadExcel {
//
//  public static void main(String[] args) throws IOException {
//    File excelFile = new File("contacts.xlsx");
//    FileInputStream fis = new FileInputStream(excelFile);
//
//    // we create an XSSF Workbook object for our XLSX Excel File
//    XSSFWorkbook workbook = new XSSFWorkbook(fis);
//    // we get first sheet
//    XSSFSheet sheet = workbook.getSheetAt(0);
//
//    // we iterate on rows
//    Iterator<Row> rowIt = sheet.iterator();
//
//    while(rowIt.hasNext()) {
//      Row row = rowIt.next();
//
//      // iterate on cells for the current row
//      Iterator<Cell> cellIterator = row.cellIterator();
//
//      while (cellIterator.hasNext()) {
//        Cell cell = cellIterator.next();
//        System.out.print(cell.toString() + ";");
//      }
//
//      System.out.println();
//    }
//
//    workbook.close();
//    fis.close();
//  }
//	
//}
