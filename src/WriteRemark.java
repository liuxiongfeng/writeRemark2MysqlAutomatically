/**
 * Created by yilun on 2017/5/23.
 */

import java.io.*;
import java.sql.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteRemark {
    //定义常量
    private final String mysqlDriver = "com.mysql.jdbc.Driver";
    private String mysqlurl = "jdbc:mysql://localhost:3306/manageplatform_fenkucdh?useUnicode=true&characterEncoding=UTF-8";
    private String mysqlName = "root";
    private String mysqlPassword = "root";
    private String excelPath = "e:\\tables.xlsx";
    private Workbook wb = null;
    private Connection mysqlconn=null;
    //检查Excel的格式是否正确
    public Message checkAndSaveRightExcel() {
        Message message = new Message(true);
        try {
            //所有表必须按照格式写否则报错
            InputStream is = new FileInputStream(excelPath);
            // 处理xlsx文件
            if (excelPath.endsWith("xlsx")) {
                wb = new XSSFWorkbook(is);
                //判断文件打开是否成功

            } else if (excelPath.endsWith("xls")) {
                // 处理xls文件
                wb = new HSSFWorkbook(is);
            }

            //检查文件Excel内有没有sheet
            int numberOfSheets = wb.getNumberOfSheets();
            if (numberOfSheets == 0) {
                message.setSuccess(false);
                message.setMessageContent("Excel文件内没有表");
                return message;
            }

            //遍历Excel所有sheet留下有正确格式的sheet 放入ArrayList内
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = wb.getSheetAt(i);
                // 得到总行数
                int rowNum = sheet.getLastRowNum();
                // 读取标题
                if(sheet.getRow(0)==null){
                    message.setSuccess(false);
                    message.setMessageContent("每一个sheet都不能为空");
                    break;
                }
                Row row = sheet.getRow(0);
               /* // 标题不一致，返回空的list.
                String title0 = getCellFormatValue(row.getCell((short) 0));
                String title1 = getCellFormatValue(row.getCell((short) 1));
                if (title0.contains("数据项") && title1.contains("字段名")) {
                    message.setSuccess(true);
                } else {
                    message.setSuccess(false);
                    message.setMessageContent("所有sheet的第一列标题应为“字段名称”第二列标题应为“字段备注”");
                    break;
                }*/
            }
            is.close();
        } catch (Exception e) {
            message.setSuccess(false);
            message.setMessageContent("发生未知错误");
        }
        return message;
    }

    public Message write() {

        Message message = new Message(true);
        HashMap<String, String> map = null;
        ArrayList<Long> al = new ArrayList();
        Statement stmt=null;
        //连接数据库
        try {
             stmt = mysqlconn.createStatement();

            //数据库的连接是昂贵的 所以没有单独把方法抽取出来 检查所有sheet的名称在数据库中是不是有对应的表 没有返回
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {

                Sheet sheet = wb.getSheetAt(i);
                String tableName = sheet.getSheetName();
                //获取源数据表id语句
                String sql = "select max(id) from bd_catalog_table where table_name = 'SG_XS_" + tableName + "' and delete_flag = 0";
                ResultSet rs = stmt.executeQuery(sql);
                long tableId = 0;
                if (rs.next()) {
                    tableId = rs.getLong(1); //源表id
                    //保存到ArrayList中
                    al.add(tableId);
                } else {//如果找不到对应的表直接跳回主方法
                    message.setSuccess(false);
                    message.setMessageContent("没有找到Excel内" + tableName + "所对应的数据库的表");
                    stmt.close();
                    mysqlconn.close();
                    return message;
                }
            }

            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                Sheet sheet = wb.getSheetAt(i);
                map = readColumnAndRemarkFromExcel(sheet);
                //根据表id获取所有的字段。
                long tableId = al.get(i);
                String colSql = "select column_name from bd_catalog_column where catalog_table_id =" + tableId;
                ResultSet colrs = stmt.executeQuery(colSql);
                List<String> colnames = new ArrayList<String>();
                while (colrs.next()) {
                    String colName = colrs.getString(1);
                    colnames.add(colName);
                }
                //根据字段和表id添加备注。
               boolean allMatch=true;
                for (String colName : colnames) {
                    //更新备注语句。
                    if(map.get(colName)!=null) {
                        String updataSql = "update bd_catalog_column set remark='" + map.get(colName) + "' where catalog_table_id=" + tableId + " and column_name='" + colName + "'";
                        stmt.execute(updataSql);
                    }else{
                        allMatch=false;
                    }
                }
                if(allMatch){
                    message.setSuccess(true);
                    message.setMessageContent("导入成功");
                }else{
                    message.setSuccess(true);
                    message.setMessageContent("导入成功，但sheet内容与对应的表的字段名称不完全匹配,建议重新检查表名");
                }

            }

        } catch (Exception e){
            e.printStackTrace();
        }finally {
            //关闭连接
            try {
                stmt.close();
                mysqlconn.close();
            }catch (SQLException e){
                e.printStackTrace();
            }
        }
        return message;
    }

    private HashMap<String, String> readColumnAndRemarkFromExcel(Sheet sheet) {
        HashMap<String, String> result = new HashMap();

        String sheetName = sheet.getSheetName();
        // 得到总行数
        int rowNum = sheet.getLastRowNum();
        // 读取正文,只读取第一列,如果某一字段有空值，或者出现中文,返回空的list.
        for (int i = 1; i <= rowNum; i++) {
            Row row = sheet.getRow(i);
            String columnStr = getCellFormatValue(row.getCell((short) 1));
            // 判断第一列是否有空值
            if (columnStr.trim().equals("")) {
                continue;
            }

            String remarkStr = getCellFormatValue(row.getCell((short) 0));
            // 判断第二列是否有空值
            if (remarkStr.trim().equals("")) {
                continue;
            }
            result.put(columnStr, remarkStr);
        }


        return result;
    }


    /**
     * 根据HSSFCell类型设置数据
     *
     * @param cell
     * @return
     */

    private String getCellFormatValue(Cell cell) {
        String cellvalue = "";
        if (cell != null) {
            // 判断当前Cell的Type
            switch (cell.getCellType()) {
                // 如果当前Cell的Type为NUMERIC
                case HSSFCell.CELL_TYPE_NUMERIC:
                    cellvalue = String.valueOf(cell.getNumericCellValue());
                    break;
                // 如果当前Cell的Type为STRIN
                case HSSFCell.CELL_TYPE_STRING:
                    // 取得当前的Cell字符串
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN:
                    cellvalue = String.valueOf(cell.getBooleanCellValue());
                    break;
                // 默认的Cell值
                default:
                    cellvalue = " ";
            }
        } else {
            cellvalue = "";
        }
        return cellvalue;

    }

    public String getMysqlurl() {
        return mysqlurl;
    }

    public void setMysqlurl(String mysqlurl) {
        this.mysqlurl = mysqlurl;
    }

    public String getExcelPath() {
        return excelPath;
    }

    public void setExcelPath(String excelPath) {
        this.excelPath = excelPath;
    }

    public String getMysqlName() {
        return mysqlName;
    }

    public void setMysqlName(String mysqlName) {
        this.mysqlName = mysqlName;
    }

    public String getMysqlPassword() {
        return mysqlPassword;
    }

    public void setMysqlPassword(String mysqlPassword) {
        this.mysqlPassword = mysqlPassword;
    }

    public Connection getMysqlconn() {
        return mysqlconn;
    }

    public void setMysqlconn(Connection mysqlconn) {
        this.mysqlconn = mysqlconn;
    }
}

