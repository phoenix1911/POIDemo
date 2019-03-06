package poi;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Tjl on 2019/3/6 9:14.
 * 百万级数据导出excel
 * XSSF  百万级数据导出很难 10w 30s
 * SXSSF 100w 60~75s
 */
public class POIWrite {
    public static void main(String[] args) throws IOException {
        XSSF();
//        SXSSF();
    }

    /**
     * 用xssf写数据到.xlsx文件
     * @throws IOException
     */
    public static void XSSF() throws IOException {


        long startTime = System.currentTimeMillis();

        // 第一步，创建一个XSSFWorkbook，对应一个Excel文件
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 第二步，创建一个sheet,对应sheet
        XSSFSheet sheet = workbook.createSheet("demo1");

        // 获取第一个表单
        Sheet first = workbook.getSheetAt(0);
        for (int i = 0; i < 1000000; i++) {
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
        FileOutputStream out = new FileOutputStream("workbook1.xlsx");
        workbook.write(out);
        out.close();

        long endTime = System.currentTimeMillis();
        System.out.println("程序运行时间：" + (endTime - startTime)/1000 + "s");

    }

    /**
     * 用sxssf写数据到.xlsx
     * @throws IOException
     */
    public static void SXSSF() throws IOException {

        long startTime = System.currentTimeMillis();

        XSSFWorkbook workbook1 = new XSSFWorkbook();
        XSSFSheet sheet = workbook1.createSheet("demo2");

        //rowAccessWindowSize 在内存中保存的行数 百万数据100行65s 1000行61s 10000行70s
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(workbook1, 10000);

        Sheet first = sxssfWorkbook.getSheetAt(0);
        for (int i = 0; i < 1000000; i++) {
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
                        CellUtil.createCell(row, j, String.valueOf(Math.random()));
                }
            }
        }
        FileOutputStream out = new FileOutputStream("workbook2.xlsx");
        sxssfWorkbook.write(out);
        out.close();

        long endTime = System.currentTimeMillis();
        System.out.println("程序运行时间：" + (endTime - startTime)/1000 + "s");
    }
}
