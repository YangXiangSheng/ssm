package com;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;

/**
 * @ClassName: controller
 * @description:
 * @author: 杨祥胜
 * @create: 2021-06-18 21:55
 **/
@Controller
public class controller {
    @RequestMapping("/test")
    public String ss(HttpServletResponse response) throws IOException {
        XSSFWorkbook xssfWorkbook=null;
        try{
//            FileInputStream file=new FileInputStream("C:\\Users\\nick\\Desktop\\档案接收函.xlsx");
            ClassPathResource classPathResource=new ClassPathResource("/档案接收函.xlsx");
            InputStream file=classPathResource.getInputStream();
            xssfWorkbook=new XSSFWorkbook(file);
            XSSFSheet sheet=xssfWorkbook.getSheetAt(0);
            sheet.setForceFormulaRecalculation(true);
            XSSFCell archivesMgrCell=sheet.getRow(6).getCell(1);
            archivesMgrCell.setCellValue("李浩有限公司"+"：");
            XSSFCell nameCell=sheet.getRow(8).getCell(3);
            nameCell.setCellValue("李浩");
//            CellRangeAddress region = new CellRangeAddress(8, 8, 8, 10);
            XSSFCell cardIdCell=sheet.getRow(8).getCell(7);
            cardIdCell.setCellValue("535464564564564556");
//            sheet.addMergedRegion(region);
            String deadlineFormat="2010年"+"7月"+"5日";
//			System.out.println(deadline[0]+"年"+deadline[1]+"月"+deadline[2]+"日");
            XSSFCell deadlineCell=sheet.getRow(10).getCell(8);
            deadlineCell.setCellValue(deadlineFormat);

            XSSFCell nowDateCell=sheet.getRow(27).getCell(5);
            nowDateCell.setCellValue("2010年"+"7月"+"5日");
            XSSFCell lastNameCell=sheet.getRow(30).getCell(1);
            lastNameCell.setCellValue("李浩");
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        xssfWorkbook.write(os);
        byte[] content = os.toByteArray();
        InputStream is = new ByteArrayInputStream(content);
        // 设置response参数，可以打开下载页面
        response.reset();
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.addHeader("Content-Disposition", "attachment;filename="+new String("档案接收函".getBytes(),"ISO8859-1") + ".xlsx");

        ServletOutputStream sout = response.getOutputStream();
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;

        try {
            bis = new BufferedInputStream(is);
            bos = new BufferedOutputStream(sout);
            byte[] buff = new byte[3072];
            int bytesRead;
            // Simple read/write loop.
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (bis != null)
                bis.close();
            if (bos != null)
                bos.close();
        }
    } catch (Exception e) {
        e.printStackTrace();
    }
        return null;
    }
}
