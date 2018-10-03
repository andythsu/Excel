package excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

    /**
     * called by other classes
     * @param workbook
     * @param sheet
     * @param data
     * @param overwrite
     * @throws IOException
     */
    public static void write(String workbook, String sheet, Map<Integer, Object[]> data, boolean overwrite) throws IOException {
        System.out.println("Writing to excel...");
        if(exist(workbook)) {
            if(overwrite) {
                System.out.println("File with same name already exists. Attempting to overwrite...");
                write(workbook, sheet, data);
            }else {
                System.out.println("File with same name already exists. No permission to overwrite.");
                return;
            }
        }else {
            write(workbook, sheet, data);
        }

    }

    /**
     * core function to write to excel
     * @param workbook
     * @param sheet
     * @param data
     * @throws IOException
     */
    private static void write(String workbook, String sheet, Map<Integer, Object[]> data) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sht = wb.createSheet(sheet);
        XSSFRow row;
        int rowid = 0;
        for (int key : data.keySet()) {
            row = sht.createRow(rowid++);
            Object[] objArr = data.get(key);
            int cellid = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String) obj);
            }
        }

        // Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File(workbook));
        wb.write(out);
        out.close();
        wb.close();
        System.out.println(workbook + " written successfully");
    }

    /**
     * check if file exists
     * @param workbook
     * @return
     */
    private static boolean exist(String workbook) {
        // TODO Auto-generated method stub
        File f = new File(workbook);
        return f.exists();
    }
}