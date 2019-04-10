import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExcelUtil {

    //工作簿 1 2
    Workbook wookbook, endworkbook;
    //    public static String excelPath = "E://poi.xls";
    File sourceExcel, endExcel;
    Sheet sheet, endSheet;
    FileInputStream fs = null;  //获取d://test.xls
    FileOutputStream os = null;
    Cell cell = null;
    Row rows = null;
    Row endRows = null;
    int Cols, endCols, rowNum = 0;


    public void createExcel(String[] title, String pathName, String sheetName) {
        try {
            sourceExcel = new File(pathName);
            wookbook = new XSSFWorkbook();
            //创建工作表sheet
            Sheet sheet = wookbook.createSheet(sheetName);
            //创建第一行
            Row row = sheet.createRow(0);
            Cell cell = null;
            //插入第一行数据的表头
            for (int i = 0; i < title.length; i++) {
                cell = row.createCell(i);
                cell.setCellValue(title[i]);
            }
            //写入数据
//        Row nrow = sheet.createRow(1);
//        Cell ncell = null;
//        for (int i = 0; i < data1.length; i++) {
//            ncell = nrow.createCell(i);
//            ncell.setCellValue(data1[i]);
//        }
            //创建excel文件
            FileOutputStream stream = new FileOutputStream(pathName);
            wookbook.write(stream);
            stream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public void addExcel(CallBack callBack, String[] data, String filePath) {
        sourceExcel = new File(filePath);
        FileInputStream fs = null;  //获取d://test.xls
        try {
            fs = new FileInputStream(sourceExcel);
//            POIFSFileSystem ps = new POIFSFileSystem(fs);  //使用POI提供的方法得到excel的信息
            wookbook = new XSSFWorkbook(fs);
            Sheet sheet = wookbook.getSheetAt(0);  //获取到工作表，因为一个excel可能有多个工作表
            Row row = sheet.getRow(0);  //获取第一行（excel中的行默认从0开始，所以这就是为什么，一个excel必须有字段列头），即，字段列头，便于赋值
            FileOutputStream out = new FileOutputStream(sourceExcel);  //向d://test.xls中写数据
            row = sheet.createRow((short) (sheet.getLastRowNum() + 1)); //在现有行号后追加数据
            for (int i = 0; i < data.length; i++) {
                row.createCell(i).setCellValue(data[i]); //设置第一个（从0开始）单元格的数据
            }
            out.flush();
            wookbook.write(out);
            out.close();
            callBack.onSuccess();
        } catch (IOException e) {
            e.printStackTrace();
            callBack.onError();
        }
    }


    /**
     * 读取excel
     */
    public void readExcel(String sourcePath, String endPath, String area, String[] areaTitle) {
        sourceExcel = new File(sourcePath);
        endExcel = new File(endPath);
        try {
            fs = new FileInputStream(sourceExcel);
            wookbook = new XSSFWorkbook(fs);
            endworkbook = new XSSFWorkbook();
            // 获得第一个工作表对象
            sheet = wookbook.getSheetAt(0);
            endSheet = endworkbook.createSheet(area);
            rowNum = sheet.getLastRowNum();
            endCols = areaTitle.length;
            for (int i = 0; i < rowNum; i++) {
                Cols = rows.getPhysicalNumberOfCells();
                rows = sheet.getRow(i);
                endRows = endSheet.createRow(i);
                endRows.createCell(i).setCellValue(areaTitle[i]);
                String str = rows.getCell(3).getStringCellValue();
                if (str.contains(area)) {
                    System.out.println(i + "====" + str);
                    for (int j = 0; j < Cols; j++) {
                        switch (j) {
                            case 0:
                                cell = endRows.createCell(j);
                                cell.setCellValue(rows.getCell(j).getNumericCellValue());
                                break;
                            case 3:
                            case 4:
                                cell = endRows.createCell(j - 2);
                                cell.setCellValue(rows.getCell(j).getStringCellValue());
                                break;
                            case 5:
                                cell = endRows.createCell(j - 2);
                                cell.setCellValue(rows.getCell(j).getNumericCellValue());
                                break;
                        }
                    }
                }

            }
            os = new FileOutputStream(endExcel);
            endworkbook.write(os);
            fs.close();
            os.close();
        } catch (Exception e) {
            System.out.println(e);
        }
    }


}
