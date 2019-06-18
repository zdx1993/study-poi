package dx.study;

import com.alibaba.fastjson.JSON;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.*;

/**
 * @description: 用于学习poi的Excel导入功能
 * @author: zhang.da.xin
 * @create: 2019-06-18 16:54
 **/


public class PoiApplication {
    public static void main(String[] args) throws IOException {
        //无论是创建Excel还是导入Excel,顶层接口都是WorkBook!
        Workbook wb = new XSSFWorkbook("C:\\Users\\dx\\GOGOGO\\study-poi\\等级评价展示数据20190618.xlsx");
        //获取sheet,从0开始
        Sheet sheet = wb.getSheetAt(0);
        //获取总行数,这里从0开始进行计算
        int lastRowNum = sheet.getLastRowNum();

        //定义行与列
        Row row;
        Cell cell;
        //定义用来接收的list
        ArrayList<Map<String, String>> list = new ArrayList<Map<String, String>>();

        //循环单元格所有的行,从1开始
//        for (int i = 1; i <= lastRowNum; i++) {
        for (int i = 2; i < 43; i++) {
            row = sheet.getRow(i);
            //获取所有的列,从1开始的
            int lastCellNum = row.getLastCellNum();
//            for (int c = 0 ; c<lastCellNum;c++){
            HashMap<String, String> map = new HashMap<>();
            for (int c = 1; c < lastCellNum; c++) {
                //获取每一个单元格
                cell = row.getCell(c);
                //对单元格的数据进行处理,获取每一个数据项
                String stringFormtValue = getStringFormtValue(cell);
                System.out.println(stringFormtValue);
                int condition = c % 6;
                System.out.println(condition);
                if(condition == 0){
                    map.put("yxqz",stringFormtValue);
                }else if(condition == 1){
                    map.put("no",stringFormtValue);
                }else if(condition == 2){
                    map.put("name",stringFormtValue);
                }else if(condition == 3){
                    map.put("address",stringFormtValue);
                }else if(condition == 4){
                    map.put("level",stringFormtValue);
                }else if(condition == 5){
                    map.put("reportdate",stringFormtValue);
                }
            }
            list.add(map);
        }
        String s = JSON.toJSONString(list);
        System.out.println(s);
    }

    public static String getStringFormtValue(Cell cell) {
        //定义返回值
        String backFormatString = null;
        //获取单元格类型
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case STRING:
                backFormatString = cell.getStringCellValue();
                break;
            case NUMERIC: //数字类型,包含数字和日期
                if (DateUtil.isCellDateFormatted(cell)) { //使用poi的提供的工具类,判断当前数据类型的值,是否为日期
                    backFormatString = formatDate(cell.getDateCellValue());
                } else {
                    backFormatString = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case FORMULA:
                backFormatString = cell.getCellFormula();
                break;
            default:
                break;
        }
        return backFormatString;
    }

    public static String formatDate(Date date) {
        LocalDateTime localDateTime = LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault());
        LocalDate localDate = localDateTime.toLocalDate();
        return localDate.toString();
    }
}
