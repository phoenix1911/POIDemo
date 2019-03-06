package poi;

import dto.Excel;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

/**
 * Created by Tjl on 2019/3/6 10:22.
 * POI读取.xlsx文件 将文件数据转为对象
 */
public class POIRead {

    public static void main(String[] args) throws IOException {
        exportXLS();
        importXLS();
    }

    public static void exportXLS() throws IOException {


        long startTime = System.currentTimeMillis();

        // 第一步，创建一个XSSFWorkbook，对应一个Excel文件
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 第二步，创建一个sheet,对应sheet
        XSSFSheet sheet = workbook.createSheet("demo1");

        // 获取第一个表单
        Sheet first = workbook.getSheetAt(0);
        for (int i = 0; i < 1000; i++) {
            Row row = first.createRow(i);
            for (int j = 0; j < 11; j++) {
                if(i == 0) {
                    // 首行
                    row.createCell(j).setCellValue("column" + j);
                } else {
                    // 数据
                    if (j == 0) {
                        CellUtil.createCell(row, j, String.valueOf(i));
                    } else
                        CellUtil.createCell(row, j, String.valueOf(Math.random()*10000));
                }
            }
        }
        // 写入文件
        FileOutputStream out = new FileOutputStream("read.xlsx");
        workbook.write(out);
        out.close();

        long endTime = System.currentTimeMillis();
        System.out.println("程序运行时间：" + (endTime - startTime)/1000 + "s");

    }

    public static void importXLS(){

        ArrayList<Excel> list = new ArrayList<>();
        try {
            //1、获取文件输入流
            InputStream inputStream = new FileInputStream("read.xlsx");
            //2、获取Excel工作簿对象
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            //3、得到Excel工作表对象
            XSSFSheet sheetAt = workbook.getSheetAt(0);
            //4、循环读取表格数据
            Excel excel=null;
            for (Row row : sheetAt) {
                //首行（即表头）不读取
                if (row.getRowNum() == 0) {
                    continue;
                }
                //读取当前行中单元格数据，索引从0开始
                String column0 = row.getCell(0).getStringCellValue();
                String column1 = row.getCell(1).getStringCellValue();
                String column2 = row.getCell(2).getStringCellValue();
                String column3 = row.getCell(3).getStringCellValue();
                String column4 = row.getCell(4).getStringCellValue();
                String column5 = row.getCell(5).getStringCellValue();
                String column6 = row.getCell(6).getStringCellValue();
                String column7 = row.getCell(7).getStringCellValue();
                String column8 = row.getCell(8).getStringCellValue();
                String column9 = row.getCell(9).getStringCellValue();

                excel = new Excel();
                excel.setColumn0(column0);
                excel.setColumn1(column1);
                excel.setColumn2(column2);
                excel.setColumn3(column3);
                excel.setColumn4(column4);
                excel.setColumn5(column5);
                excel.setColumn6(column6);
                excel.setColumn7(column7);
                excel.setColumn8(column8);
                excel.setColumn9(column9);
                System.out.println(excel);
                list.add(excel);
            }
            //5、关闭流
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("数据："+ list);
    }


}
