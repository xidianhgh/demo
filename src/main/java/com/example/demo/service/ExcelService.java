package com.example.demo.service;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.UUID;

@Service
public class ExcelService {

    public Workbook create() throws Exception {
        // 创建excel工作簿
        HSSFWorkbook wb = new HSSFWorkbook();

        HSSFSheet sheet = wb.createSheet();
        for (int i = 0; i < 10; i++) {
            HSSFCell cell = sheet.createRow(i).createCell(0);
            cell.setCellValue("俄罗斯" + i);
//            cell.setCellStyle(cellStyle);
        }

        CellRangeAddress region = new CellRangeAddress(0, 3, 0, 0);
        sheet.addMergedRegion(region);
        CellStyle cellStyle = getStyle(wb);
        cellStyle.setFont(getFont(wb));
        sheet.getRow(region.getFirstRow()).getCell(region.getFirstColumn()).setCellStyle(cellStyle);

        return wb;
    }

    //生成并下载Excel
    public void download(Workbook workbook, HttpServletResponse response) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        String fileName = UUID.randomUUID().toString();
        try {
            workbook.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
        byte[] content = os.toByteArray();
        InputStream is = new ByteArrayInputStream(content);
        // 设置response参数
        response.reset();
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + new String((fileName + ".xls").getBytes(), "utf-8"));
        ServletOutputStream out = response.getOutputStream();
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        try {
            bis = new BufferedInputStream(is);
            bos = new BufferedOutputStream(out);
            byte[] buff = new byte[2048];
            int bytesRead;
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
        } catch (final IOException e) {
            throw e;
        } finally {
            if (bis != null)
                bis.close();
            if (bos != null)
                bos.close();
        }
    }

    //生成并下载Excel
    public void downloadx(Workbook workbook, HttpServletResponse response) throws IOException {
        String fileName = UUID.randomUUID().toString();
        // 设置response参数
        response.reset();
        response.setContentType("application/ms-excel;charset=UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename=".concat(String.valueOf(URLEncoder.encode("xxx.xls", "UTF-8"))));
        try {
            workbook.write(response.getOutputStream());
            response.getOutputStream().flush();
        } catch (Exception e) {
            System.out.println(e.getMessage());
        } finally {
            if (workbook != null) {
                workbook.close();
            }
            if (response.getOutputStream() != null) {
                response.getOutputStream().close();
            }
        }

    }

    public CellStyle getStyle(Workbook wb) throws Exception {
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    public Font getFont(HSSFWorkbook wb) {
        String str = "#02C874";
//处理把它转换成十六进制并放入一个数
        int[] color = new int[3];
        color[0] = Integer.parseInt(str.substring(1, 3), 16);
        color[1] = Integer.parseInt(str.substring(3, 5), 16);
        color[2] = Integer.parseInt(str.substring(5, 7), 16);
        HSSFPalette customPalette = wb.getCustomPalette();
        customPalette.setColorAtIndex(HSSFColor.BLACK.index, (byte) color[0], (byte) color[1], (byte) color[2]);

        Font font = wb.createFont();
        font.setColor(HSSFColor.BLACK.index);
        return font;
    }

    public Font getFont(XSSFWorkbook wb) {
        String str = "#003D79";
//处理把它转换成十六进制并放入一个数
        int r = Integer.parseInt(str.substring(1, 3), 16);
        int g = Integer.parseInt(str.substring(3, 5), 16);
        int b = Integer.parseInt(str.substring(5, 7), 16);
        Font font = wb.createFont();
        java.awt.Color color = new java.awt.Color(r, g, b);
        font.setColor(new XSSFColor(color).getIndex());
        return font;
    }


}
