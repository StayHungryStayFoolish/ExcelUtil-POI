package com.excel.utils;

import com.excel.bean.ImportExcelBean;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by bonismo@hotmail.com
 * 上午10:32 on 17/11/1.
 */
public class ReadExcel {

    public static void main(String[] args) throws IllegalAccessException, InvocationTargetException, InstantiationException, SecurityException, NoSuchMethodException {
        File path = new File("/Users/bonismo/Desktop/1023.xls");
        String[] columnName = new String[]{"orgCode", "orgName", "userCode", "userName"};
        List<ImportExcelBean> list = ReadExcel.parse(ImportExcelBean.class, columnName, path);
        for (ImportExcelBean excelBean : list) {
            excelBean.getUserName();
            System.out.println(excelBean.getUserName());
        }

    }

    /**
     * 根据传入的 Bean 类，域字段，文件路径分析 Excel
     * @param cc Excel 对应的类模型（利用反应，反射机制处理）
     * @param columnName Excel 列 对应的 Bean 域
     * @param file Excel 文件路径
     * @param <T> 泛型
     * @return 返回 Bean 类的 List 集合
     */
    public static <T> List<T> parse(Class<T> cc, String[] columnName, File file) throws IllegalAccessException, InvocationTargetException, InstantiationException, SecurityException, NoSuchMethodException {
        List<T> list = new ArrayList<T>();
        List<String[]> data = parse(file);
        for (int i = 0; i < data.size(); i++) {
            String[] row = data.get(i);
            Map<String, Object> map = new HashMap<>();
            for (int k = 0; k < row.length; k++) {
                map.put(columnName[k], row[k]);
            }
            T bean = BeanHelper.bean(map, cc);
            list.add(bean);
        }
        return list;
    }

    /**
     * 根据传入的文件路径，分析 Excel
     * @param file 文件路径
     * @return String 数组的 List 集合
     */
    public static List<String[]> parse(File file) throws IllegalAccessException, InvocationTargetException, InstantiationException, SecurityException, NoSuchMethodException {
        List<String[]> list = new ArrayList<>();
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
            HSSFSheet sheet = workbook.getSheetAt(0);
            int rows = sheet.getPhysicalNumberOfRows();
            int cols = sheet.getRow(0).getPhysicalNumberOfCells();
            // 遍历行,索引从0 开始，第0行可用作表头，不获取。
            for (int i = 1; i < rows; i++) {
                // 读取左上端单元格
                HSSFRow row = sheet.getRow(i);
                // 行不为空
                if (row != null) {
                    // 获取到Excel文件中的所有的列
                    String value = "";
                    // 防止当使用POI处理excel的时候cell.getNumbericCellValue()
                    // 当长度大一点的时候会变成科学计数法形式。
                    DecimalFormat df = new DecimalFormat("0");
                    // ***下方3为cells的值，更换为固定列数，解决空值问题***
                    for (int j = 0; j < cols; j++) {
                        HSSFCell cell = row.getCell(j);
                        if (cell != null) {
                            row.getCell(j).setCellType(HSSFCell.CELL_TYPE_STRING);
                            value += cell.getStringCellValue() + ",";
                        } else {
                            value += "#" + ",";
                        }
                    }
                    String[] val = value.split(",");
                    String[] arr = new String[cols];
                    for (int k = 0; k < cols; k++) {
                        arr[k] = val[k];
                    }
                    list.add(arr);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return list;
    }
}
