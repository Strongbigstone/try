package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.util.*;

public class XlsPreviewer {

    private static final int MAX_PREVIEW_ROWS = 100; // 限制预览行数

    public String previewXls(String filePath) throws Exception {
        StringBuilder html = new StringBuilder();
        html.append("<html><body><table border='1'>");

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new HSSFWorkbook(fis)) { // 使用 HSSF 处理 .xls

            Sheet sheet = workbook.getSheetAt(0); // 读取第一个工作表
            Iterator<Row> rowIterator = sheet.iterator();

            int rowCount = 0;
            while (rowIterator.hasNext() && rowCount < MAX_PREVIEW_ROWS) {
                Row row = rowIterator.next();
                html.append("<tr>");

                // 处理空行（POI可能跳过空行）
                int maxCellIndex = row.getLastCellNum();
                for (int cn = 0; cn <= maxCellIndex; cn++) {
                    Cell cell = row.getCell(cn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    html.append("<td>").append(getCellValue(cell)).append("</td>");
                }

                html.append("</tr>");
                rowCount++;
            }

        } catch (Exception e) {
            return "解析失败: " + e.getMessage();
        }

        html.append("</table></body></html>");
        return html.toString();
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return "公式: " + cell.getCellFormula();
            default:
                return "";
        }
    }

    public static void main(String[] args) throws Exception {
        // 示例用法
        String result = new XlsPreviewer().previewXls("D:\\DK\\Desktop\\测试1.xls");
        System.out.println(result); // 输出 HTML 表格
        // 实际应用中可将 HTML 写入文件或通过 HTTP 返回
    }
}
