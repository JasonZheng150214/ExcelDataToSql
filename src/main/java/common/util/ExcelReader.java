package common.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

public class ExcelReader {
    /**
     * 对外提供读取excel 的方法
     */
    public static List<List<Object>> readExcel(File file) throws IOException {
        String fileName = file.getName();
        String extension = fileName.lastIndexOf(".") == -1 ?
                "" :
                fileName.substring(fileName.lastIndexOf(".") + 1);
        if ("xls".equals(extension)) {
            return read2003Excel(file,null);
        } else if ("xlsx".equals(extension)) {
            return read2007Excel(file,null);
        } else {
            throw new IOException("不支持的文件类型");
        }
    }

    /**
     * 读取 office 2003 excel
     *
     * @throws IOException
     * @throws FileNotFoundException
     */
    private static List<List<Object>> read2003Excel(File file,Integer colspan) throws IOException {
        List<List<Object>> list = new LinkedList<List<Object>>();
        HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = hwb.getSheetAt(0);
        Object value = null;
        HSSFRow row = null;
        HSSFCell cell = null;
        //行数循环
        for (int i = sheet.getFirstRowNum(); i < sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);//第i行对象
            if (row == null) {
                continue;
            }
            List<Object> linked = new LinkedList<Object>();
            if (row.getFirstCellNum() < 0) {
                break;
            }

            if(colspan ==null){
                colspan = Integer.valueOf(row.getLastCellNum());//多少列
            }
            for (int j = row.getFirstCellNum(); j < colspan; j++) {
                cell = row.getCell(j);
                if (cell == null) {
                    value = "";

                } else {
                    DecimalFormat df = new DecimalFormat("0.00000000");// 格式化 number String 字符
                    DecimalFormat nf = new DecimalFormat("0");// 格式化数字
                    switch (cell.getCellType()) {
                        case XSSFCell.CELL_TYPE_STRING:
                            value = cell.getStringCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            //    System.out.println(i+"行"+j+" 列 is Number type ; DateFormt:"+cell.getCellStyle().getDataFormatString());
                            /*if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                                value = df.format(cell.getNumericCellValue());
                            } else if ("General".equals(cell.getCellStyle().getDataFormatString())) {
                                value = nf.format(cell.getNumericCellValue());
                            } else {
                                //    value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
                                value = nf.format(cell.getNumericCellValue());
                            }*/

                            String str = ""+cell.getNumericCellValue();
                            if(null != str && !"".equals(str)){
                                int length = str.length()-str.indexOf(".")-1;
                                if(str.contains(".0") && length==1){
                                    value = nf.format(cell.getNumericCellValue());
                                }else{
                                    value = df.format(cell.getNumericCellValue());
                                }
                                if(value.toString().contains(".00000000")){
                                    value = value.toString().replace(".00000000","");
                                }
                            }
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            value = cell.getBooleanCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_BLANK:
                            value = "";
                            break;
                        default:
                            value = cell.toString();
                    }
                    if (value == null || ((String) value).trim().equals("")) {
                        value = "";
                    }
                }
                linked.add(value);
            }
            list.add(linked);
        }
        return list;
    }

    private static Boolean isBlankRow(XSSFRow row) {
        if (row == null)
            return true;
        try {
            XSSFCell cell = null;
            for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                cell = row.getCell(j);
                if (cell == null)
                    continue;

                switch (cell.getCellType()) {
                    case XSSFCell.CELL_TYPE_STRING:
                        String value = cell.getStringCellValue();
                        if (!StringUtils.isEmpty(value) && !StringUtils.isEmpty(value.trim()))
                            return false;
                        break;
                    case XSSFCell.CELL_TYPE_BOOLEAN:
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        return false;
                    case XSSFCell.CELL_TYPE_BLANK:
                        break;
                }
            }
        } catch (Exception e) {
            return true;
        }
        return true;
    }

    /**
     * 读取Office 2007 excel
     */
    private static List<List<Object>> read2007Excel(File file,Integer colspan) throws IOException {
        List<List<Object>> list = new LinkedList<List<Object>>();
        // 构造 XSSFWorkbook 对象，strPath 传入文件路径
        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(file));
        // 读取第一章表格内容
        XSSFSheet sheet = xwb.getSheetAt(0);
        Object value = null;
        XSSFRow row = null;
        XSSFCell cell = null;
        int size = sheet.getPhysicalNumberOfRows();
        for (int i = sheet.getFirstRowNum(); i < size; i++) {

            row = sheet.getRow(i);
            //   if (row == null || isBlankRow(row)) {
            if (row == null) {
                continue;
            }
            List<Object> linked = new LinkedList<Object>();
            if (row.getFirstCellNum() < 0)
                break;
            if(colspan ==null){
                colspan = Integer.valueOf(row.getLastCellNum());
            }
            for (int j = row.getFirstCellNum(); j < colspan; j++) {
                //            for (int j = row.getFirstCellNum(); j < 40; j++) {
                cell = row.getCell(j);
                if (cell == null) {
                    value = "0";

                    //     continue;
                } else {


                    DecimalFormat df = new DecimalFormat("0.00000000");// 格式化 number String 字符
                    DecimalFormat nf = new DecimalFormat("0");// 格式化数字
                    switch (cell.getCellType()) {
                        case XSSFCell.CELL_TYPE_STRING:
                            //   System.out.println(i+"行"+j+" 列 is String type");
                            value = cell.getStringCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                            //   System.out.println(i+"行"+j+" 列 is Number type ; DateFormt:"+cell.getCellStyle().getDataFormatString());
                            /*if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                                value = df.format(cell.getNumericCellValue());
                            } else if ("General"
                                .equals(cell.getCellStyle().getDataFormatString())) {
                                value = nf.format(cell.getNumericCellValue());
                            } else {
                                //   value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
                                value = nf.format(cell.getNumericCellValue());
                            }*/
                            String str = ""+cell.getNumericCellValue();
                            if(!StringUtils.isEmpty(str)){
                                int length = str.length()-str.indexOf(".")-1;
                                if(str.contains(".0") && length==1){
                                    value = nf.format(cell.getNumericCellValue());
                                }else{
                                    value = df.format(cell.getNumericCellValue());
                                }
                                if(value.toString().contains(".00000000")){
                                    value = value.toString().replace(".00000000","");
                                }
                            }
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN:
                            //   System.out.println(i+"行"+j+" 列 is Boolean type");
                            value = cell.getBooleanCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_BLANK:
                            //      System.out.println(i+"行"+j+" 列 is Blank type");
                            value = "0";
                            break;
                        default:
                            //   System.out.println(i+"行"+j+" 列 is default type");
                            value = cell.toString();
                    }
                    if (value == null || ((String) value).trim().equals("")) {
                        //     continue;
                        value = "0";
                    }

                }
                linked.add(value);
            }
            list.add(linked);
        }
        return list;
    }
}
