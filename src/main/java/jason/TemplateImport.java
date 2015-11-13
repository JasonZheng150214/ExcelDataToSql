package jason;

import common.util.ExcelReader;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.Scanner;

/**
 * Created by jason on 15/8/17.
 */
public class TemplateImport {


    public static void main(String[] args) {

        System.out.println("请输入源文件全路径:例子/users/jason/Documents/SQL/1.xls");
        Scanner sc = new Scanner(System.in);
        String sourceFilePath = sc.next();

        File file = new File(sourceFilePath);
        String fileName = file.getName();
        String extension = fileName.lastIndexOf(".") == -1 ?
                "" :
                fileName.substring(fileName.lastIndexOf(".") + 1);
        if (!"xls".equals(extension) && ! "xlsx".equals(extension)) {
            System.out.println("输入的格式不对,不支持该文件类型!只支持xls和xlsx");
        } else if (!file.exists()) {
            System.out.println("该文件不存在!");
        } else {
            System.out.println("请输入数据库表名:例子legend_shop");
            String tableName = sc.next();

            //读取excel数据
            readExcelData(tableName, file);
        }


    }

    /**
     * 读取excel数据
     *
     */
    public static void readExcelData(String tableName, File file) {

        try {
            //开始Excel数据导入
            if (file.isFile() && file.exists()) {
                List<List<Object>> excelList = ExcelReader.readExcel(file);
                if(!CollectionUtils.isEmpty(excelList)){
                    //组装insert into 头部sql
                    String prefixStr = createPrefixStr(tableName,excelList);
                    Integer fieldSize = excelList.get(0).size();//源文件字段数
                    Integer excelListSize = excelList.size();//源文件总行数
                    if(excelListSize>1) {

                        List totalList = new ArrayList();

                        for (int i = 1; i < excelListSize; i++) {
                            List list = new LinkedList();

                            //组装一行数据
                            for (int j = 0; j < fieldSize; j++) {
                                list.add(excelList.get(i).get(j));
                            }
                            totalList.add(list);
                        }

                        if (!CollectionUtils.isEmpty(totalList)) {

                            //生成批量insert sql
                            batchInsert(prefixStr,totalList,file);
                        } else {
                            System.out.println(">>>>>>>>>>>>源文件没有数据!");

                        }


                    }
                }
            } else {
                System.out.println(">>>>>>>>>>>>文件不存在");

            }
        } catch (Exception e) {
            System.out.println(">>>>>>>>>>>>读取excel异常");
        }
    }



    public static void batchInsert(String prefixStr ,List list,File file) {

        //清空数据
        String sqlFilePath =  writeToFile(file,"",0);

        Integer MAX_SIZE = 1000 ;//默认批量插入1000行

        int totalSize = list.size();//总行数
        int size = totalSize % MAX_SIZE == 0 ? totalSize / MAX_SIZE : totalSize / MAX_SIZE + 1;
        List<List> subList;

        for (int i = 0; i < size; i++) {
            if (i + 1 == size) {
                subList = list.subList(i * MAX_SIZE, totalSize);
            } else {
                subList = list.subList(i * MAX_SIZE, i * MAX_SIZE + MAX_SIZE);
            }

            StringBuffer sb = new StringBuffer(prefixStr);
            for (List listVo : subList) {
                sb.append(" ( now(), now(), ");

                for (Object s : listVo) {

                    sb.append("'").append(s).append("'").append(",");
                }
                sb.deleteCharAt(sb.lastIndexOf(",")).append(")").append(",");

            }
            sb.deleteCharAt(sb.lastIndexOf(",")).append(";");

            //写批量的insert sql到文件中
            writeToFile(file, sb.toString(),1);

        }

        System.out.println(">>>>>>>输出sql到文件成功! 文件路径" + sqlFilePath);

    }


    public static String writeToFile (File file, String data,Integer type) {

        //默认输入的sql的文件路径
        String filePath = file.getParent() + "/sql.txt";

        //BufferedWriter bw = null;
        try {
            if (!file.exists()) {
                file.createNewFile();
            }
            FileWriter fw;
            if (type == 0) {
                fw = new FileWriter(filePath);//清空文件内容再插入
            } else {
                fw = new FileWriter(filePath,true);//append到文件末尾
            }

            fw.write(data);
            fw.close();
            //bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(fileName,true)));
            //bw.write(data);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return filePath;

    }

    //组装insert into 头
    //默认excel 第一行是数据库的对应的字段名称
    public static String createPrefixStr (String tableName, List<List<Object>> excelList) {
        Integer fieldSize = excelList.get(0).size();//源文件字段数
        StringBuffer prefixSb = new StringBuffer("INSERT INTO ");
        prefixSb.append(tableName).append(" ( gmt_create, gmt_modified, ");
        for (int j = 0; j < fieldSize; j++) {

            String fieldName = (String) excelList.get(0).get(j);
            if(!StringUtils.isEmpty(fieldName)){
                prefixSb.append(fieldName).append(",");
            }
        }
        //截掉最后一个逗号
        prefixSb.deleteCharAt(prefixSb.lastIndexOf(",")).append(") VALUES");

        return prefixSb.toString();


    }
}
