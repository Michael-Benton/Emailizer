package emailizer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import javax.swing.JFileChooser;
import javax.swing.JPanel;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class read_in_excel
{
  static void read_excel()
    throws IOException
  {
    // Creates a file explorer window to allow the user to select a file with a specific extension
    JFileChooser fileChooser = new JFileChooser();
    
    FileNameExtensionFilter filter = new FileNameExtensionFilter("xls", new String[] { "xlsx" });
    
    fileChooser.setFileFilter(filter);
    
    fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
    
    int returnVal = fileChooser.showOpenDialog(new JPanel());

    if (returnVal == 0)
    {
      File OGFile = fileChooser.getSelectedFile();
      
      InputStream inputStreamFile = new FileInputStream(OGFile);
      XSSFWorkbook wb = new XSSFWorkbook(inputStreamFile);
      
      XSSFSheet sheet = wb.getSheetAt(0);
      ArrayList<String> data = new ArrayList();
      Iterator<Row> rowIterator = sheet.iterator();
      
      int columnIndexEmail = -5;
      int columnIndexHasEmail = -6;

      for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++)
      {
        Row row = sheet.getRow(0);
        Cell cellEmail = row.getCell(i);
        if (cellEmail != null) {
          if ("email".equals(cellEmail.getStringCellValue().toLowerCase())) {
            columnIndexEmail = i;
          } else if ("hasemail".equals(cellEmail.getStringCellValue().toLowerCase())) {
            columnIndexHasEmail = i;
          }
        }
      }

      if (columnIndexEmail == -5) {
        System.exit(0);
      }

      for (int rowIndex = 2; rowIndex <= sheet.getPhysicalNumberOfRows(); rowIndex++)
      {
        Row row = CellUtil.getRow(rowIndex, sheet);
        Cell cellEmail = CellUtil.getCell(row, columnIndexEmail);
        Cell cellHasEmail = CellUtil.getCell(row, columnIndexHasEmail);
        if ((cellHasEmail.getNumericCellValue() != 0.0D) || 
        
          (data.contains(cellEmail.getStringCellValue()))) {
          sheet.removeRow(row);
        } else {
          data.add(cellEmail.getStringCellValue());
        }
      }

      inputStreamFile.close();
      
      FileOutputStream outFile = new FileOutputStream(new File("UpdatedFile.xlsx"));
      wb.write(outFile);
      outFile.close();
    }
  }
}