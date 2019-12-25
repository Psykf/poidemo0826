package poi0826;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

public class PoiTest {

    //读操作
    @Test
    public void readExcel() throws Exception{
        //1 输入流获取文件
        InputStream in = new FileInputStream("E:/0826/0826.xls");
        //2 创建workbook
        Workbook workbook = new HSSFWorkbook(in);
        //3 获取sheet
        Sheet sheet = workbook.getSheetAt(0);
        //4 获取row
        Row row = sheet.getRow(0);
        //5 获取cell
        Cell cellone = row.getCell(0);
        Cell celltwo = row.getCell(1);
        //6 从cell获取内容
        String stringCellValue1 = cellone.getStringCellValue();
        // String stringCellValue2 = celltwo.getStringCellValue();
        double numericCellValue = celltwo.getNumericCellValue(); //读不同的数据类型
        System.out.println(stringCellValue1);
        System.out.println(numericCellValue);
    }





    @Test
    public void writeExcel() throws Exception{
        //1 创建workbook
        //HSSFWorkbook 03版本excel   xls
        //XSSFWorkbook 07版本excel  xlsx
        //Workbook workbook = new XSSFWorkbook();
        Workbook workbook = new HSSFWorkbook();
        //2 创建sheet
        Sheet sheet = workbook.createSheet("用户管理");
        //3 创建row
        Row row = sheet.createRow(0);
        //4 创建cell
        Cell cell = row.createCell(0);
        //5 设置内容
        cell.setCellValue("name");

        //6 通过输出流
        OutputStream out = new FileOutputStream("E:/0826/0826.xls");
        workbook.write(out);
        out.close();
    }

}
