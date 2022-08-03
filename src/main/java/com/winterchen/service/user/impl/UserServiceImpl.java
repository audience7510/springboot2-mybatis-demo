package com.winterchen.service.user.impl;

import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import com.winterchen.dao.UserDao;
import com.winterchen.model.UserDomain;
import com.winterchen.service.user.UserService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.util.List;

/**
 * Created by Administrator on 2017/8/16.
 */
@Slf4j
@Service(value = "userService")
public class UserServiceImpl implements UserService {

    @Autowired
    private UserDao userDao;//这里会报错，但是并不会影响

    @Override
    public int addUser(UserDomain user) {

        return userDao.insert(user);
    }

    /*
    * 这个方法中用到了我们开头配置依赖的分页插件pagehelper
    * 很简单，只需要在service层传入参数，然后将参数传递给一个插件的一个静态方法即可；
    * pageNum 开始页数
    * pageSize 每页显示的数据条数
    * */
    @Override
    public PageInfo<UserDomain> findAllUser(int pageNum, int pageSize) {
        //将参数传给这个方法就可以实现物理分页了，非常简单。
        PageHelper.startPage(pageNum, pageSize);
        List<UserDomain> userDomains = userDao.selectUsers();
        PageInfo result = new PageInfo(userDomains);
        return result;
    }

    @Override
    public void exportExcel(HttpServletResponse response) {
        try{
            //需要通过response给前端数据流，设置对应的response
            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-Disposition","attachment;filename="+"text.xlsx");

            //创建一张工作簿workbook
            Workbook workbook = new XSSFWorkbook();

            //在工作簿中创建一张表sheet
            Sheet sheet = workbook.createSheet("sheet1");
            //创建个cell的颜色背景
            CellStyle cellStyle = workbook.createCellStyle();
            //设置背景黄色
            cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            //水平对齐方式（居中）
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            //在表中创建一行row1，对应是sheet.createRow(0);
            Row row1 = sheet.createRow(0);
            int firstRow = 1;//需要合并的第一个单元格的行数
            int lastRow = 1;//需要合并的最后一个单元格的行数
            int firstCol = 1;//需要合并的第一个单元格的列数
            int lastCol = 5;//需要合并的最后一个单元格的列数
            CellRangeAddress cellRangeAddress1 = new CellRangeAddress(firstRow,lastRow,firstCol,lastCol);
            sheet.addMergedRegion(cellRangeAddress1);

            int firstRow2 = 3;//需要合并的第一个单元格的行数
            int lastRow2 = 3;//需要合并的最后一个单元格的行数
            int firstCol2 = 1;//需要合并的第一个单元格的列数
            int lastCol2 = 5;//需要合并的最后一个单元格的列数
            CellRangeAddress cellRangeAddress2 = new CellRangeAddress(firstRow2,lastRow2,firstCol2,lastCol2);
            sheet.addMergedRegion(cellRangeAddress2);
            //在表中创建第二行row2,对应是sheet.createRow(1);
            Row row2 = sheet.createRow(1);
            //第二行第一个单元格是张三，学好是123456
            Cell row2Cell1 = row2.createCell(1);
            row2Cell1.setCellStyle(cellStyle);
            row2Cell1.setCellValue("Leased Site Alarm Notification-drill（数据中心告警通知-演练）");

            Row row3 = sheet.createRow(2);
            Cell row3Cell1 = row3.createCell(1);
            row3Cell1.setCellValue("\u0052 东花园");
            Cell row3Cell2 = row3.createCell(2);
            row3Cell2.setCellValue("\u00A3 灵丘");
            Cell row3Cell3 = row3.createCell(3);
            row3Cell3.setCellValue("☑总部基地");
            Cell row3Cell4 = row3.createCell(4);
            row3Cell4.setCellValue("⬜总部基地");

            Row row4 = sheet.createRow(3);
            Cell row4Cell1 = row4.createCell(1);
            row4Cell1.setCellStyle(cellStyle);
            row4Cell1.setCellValue("☑总部基地 ☑东花园 ⬜灵丘 ");
            //使用数据流返回给前端
            OutputStream out = response.getOutputStream();
            workbook.write(out);
            out.close();

            //也直接使用文件流
            //FileOutputStream fileOutputStream = new FileOutputStream("E:\\workspace_2021(IDEA)\\export_excel\\test.xlsx");
            //workbook.write(fileOutputStream);
            //fileOutputStream.close();
        }catch (Exception e) {
            log.info(e.getMessage());
        }

    }
}
