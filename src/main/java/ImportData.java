import common.util.ExcelReader;
import entity.ActGoodsCarRel;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

/**
 * Created by jason on 15/8/17.
 */
public class ImportData {


    public static void main(String[] args) {

        System.out.println("请输入源文件全路径:例子/users/jason/Documents/SQL/1.xls");
        Scanner sc = new Scanner(System.in);
        //String sourceFilePath = "/users/jason/Documents/SQL/1.xls";
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

            System.out.println(sourceFilePath);
            importDataToDB(sourceFilePath);
        }


    }

    public static void importDataToDB(String filePath) {

        try {
            //开始Excel数据导入
            File file = new File(filePath);
            if (file.isFile() && file.exists()) {
                List<List<Object>> excelList = ExcelReader.readExcel(file);
                if(!CollectionUtils.isEmpty(excelList)){
                    Integer excelListSize = excelList.size();
                    if(excelListSize>1) {

                        List<ActGoodsCarRel> actGoodsCarRelList = new ArrayList();
                        //数据校验
                        for (int i = 1; i < excelListSize; i++) {

                            ActGoodsCarRel actGoodsCarRel = new ActGoodsCarRel();
                            if(!StringUtils.isEmpty(excelList.get(i).get(0))){
                                actGoodsCarRel.setServiceType(Integer.valueOf(excelList.get(i).get(0).toString()));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(1))){
                                actGoodsCarRel.setAdaptStandard((String) excelList.get(i).get(1));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(2))){
                                actGoodsCarRel.setAdaptService((String) excelList.get(i).get(2));
                            }
                            if(!StringUtils.isEmpty( excelList.get(i).get(3))){
                                actGoodsCarRel.setCarBrandName((String) excelList.get(i).get(3));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(4))){
                                actGoodsCarRel.setCarBrandId(Long.valueOf(excelList.get(i).get(4).toString()));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(5))){
                                actGoodsCarRel.setCarSeriesName((String) excelList.get(i).get(5));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(6))){
                                actGoodsCarRel.setCarSeriesId(Long.valueOf(excelList.get(i).get(6).toString()));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(7))){
                                actGoodsCarRel.setCarTypeName((String) excelList.get(i).get(7));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(8))){
                                actGoodsCarRel.setCarTypeId(Long.valueOf(excelList.get(i).get(8).toString()));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(9))){
                                actGoodsCarRel.setCarYearName((String) excelList.get(i).get(9));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(10))){
                                actGoodsCarRel.setCarYearId(Long.valueOf(excelList.get(i).get(10).toString()));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(11))){
                                actGoodsCarRel.setCarDetailName((String)excelList.get(i).get(11));
                            }
                            if(!StringUtils.isEmpty(excelList.get(i).get(12))){
                                actGoodsCarRel.setCarDetailId(Long.valueOf(excelList.get(i).get(12).toString()));
                            }

                            actGoodsCarRelList.add(actGoodsCarRel);
                        }

                        batchInsert(actGoodsCarRelList,1000);

                    }
                }
            } else {
                System.out.println(">>>>>>>>>>>>文件不存在");

            }
        } catch (Exception e) {
            System.out.println(">>>>>>>>>>>>读取excel异常");
        }
    }



    public static void batchInsert(List<ActGoodsCarRel> list , Integer MAX_SIZE ) {
        String fileName = "/Users/jason/Documents/SQL/sql.txt";

        if(MAX_SIZE == null){
            MAX_SIZE = 1000 ;
        }
        int totalSize = list.size();
        int size = totalSize % MAX_SIZE == 0 ? totalSize / MAX_SIZE : totalSize / MAX_SIZE + 1;
        List<ActGoodsCarRel> subList;
        String prefixStr = "INSERT INTO legend_act_goods_car_rel_zc (service_type,adapt_standard," +
                "adapt_service,car_brand_id,car_brand_name,car_series_id,car_series_name,car_type_id," +
                "car_type_name,car_year_id,car_year_name,car_detail_id,car_detail_name) VALUES ";
        for (int i = 0; i < size; i++) {
            if (i + 1 == size) {
                subList = list.subList(i * MAX_SIZE, totalSize);
            } else {
                subList = list.subList(i * MAX_SIZE, i * MAX_SIZE + MAX_SIZE );
            }

            StringBuffer sb = new StringBuffer(prefixStr);
            for (ActGoodsCarRel a : subList) {
                sb.append("(");
                sb.append(a.getServiceType()).append(",")
                        .append("'").append(a.getAdaptStandard()).append("'").append(",")
                        .append("'").append(a.getAdaptService()).append("'").append(",")
                        .append(a.getCarBrandId()).append(",")
                        .append("'").append(a.getCarBrandName()).append("'").append(",")
                        .append(a.getCarSeriesId()).append(",")
                        .append("'").append(a.getCarSeriesName()).append("'").append(",")
                        .append(a.getCarTypeId()).append(",")
                        .append("'").append(a.getCarTypeName()).append("'").append(",")
                        .append(a.getCarYearId()).append(",")
                        .append("'").append(a.getCarYearName()).append("'").append(",")
                        .append(a.getCarDetailId()).append(",")
                        .append("'").append(a.getCarDetailName()).append("'").append("),");
            }
            sb.deleteCharAt(sb.lastIndexOf(",")).append(";");

            writeToFile(fileName,sb.toString());

        }
    }


    public static void writeToFile (String fileName, String data) {

        //BufferedWriter bw = null;
        try {
            File file = new File(fileName);

            if (!file.exists()) {
                file.createNewFile();
            }
            FileWriter fw = new FileWriter(fileName,true);

            fw.write(data);
            fw.close();
            //bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(fileName,true)));
            //bw.write(data);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
