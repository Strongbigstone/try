package org.example;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelConverter1 extends JFrame {

    private JTextField inputFileField;
    private JButton convertButton;
    private JProgressBar progressBar;

    public ExcelConverter1() {
        super("Excel 文件转换工具");
        initUI();
    }

    public static void convertXlsToXlsx(File inputFile, File outputFile) throws Exception {
        try (InputStream in = new FileInputStream(inputFile);
             Workbook inputWorkbook = WorkbookFactory.create(in)) {

            // 收集图片信息
            Map<Sheet, List<PictureInfo>> sheetPicturesMap = collectAllPictures(inputWorkbook);

            try (Workbook outputWorkbook = new XSSFWorkbook()) {
                // 转换样式映射
                Map<CellStyle, CellStyle> styleCache = new HashMap<>();

                // 转换每个工作表
                for (int i = 0; i < inputWorkbook.getNumberOfSheets(); i++) {
                    Sheet inputSheet = inputWorkbook.getSheetAt(i);
                    Sheet outputSheet = outputWorkbook.createSheet(inputSheet.getSheetName());

                    // 设置列宽
//                    for (int col = 0; col <= inputSheet.getRow(0).getLastCellNum(); col++) {
//                        outputSheet.setColumnWidth(col, inputSheet.getColumnWidth(col));
//                    }
                    for (int row = 0; row < inputSheet.getLastRowNum(); row++) {
                        for (int col = 0; col < inputSheet.getRow(row).getLastCellNum(); col++) {
                            outputSheet.setColumnWidth(col, inputSheet.getColumnWidth(col));
                        }
                    }

                    // 复制行和单元格
                    copySheetContent(inputSheet, outputSheet, styleCache, inputWorkbook, outputWorkbook);

                    // 添加图片
                    addPicturesToSheet(sheetPicturesMap.get(inputSheet), inputSheet, outputSheet, outputWorkbook);
                }

                // 保存结果
                try (FileOutputStream out = new FileOutputStream(outputFile)) {
                    outputWorkbook.write(out);
                }
            }
        }
    }

    // 图片信息存储类
    private static class PictureInfo {
        byte[] imageData;
        String mimeType;
        ClientAnchor anchor;
        int width;
        int height;

        PictureInfo(byte[] imageData, String mimeType, ClientAnchor anchor, int width, int height) {
            this.imageData = imageData;
            this.mimeType = mimeType;
            this.anchor = anchor;
            this.width = width;
            this.height = height;
        }
    }

    // 收集所有图片信息
    private static Map<Sheet, List<PictureInfo>> collectAllPictures(Workbook workbook) {
        Map<Sheet, List<PictureInfo>> sheetPicturesMap = new HashMap<>();

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            List<PictureInfo> pictureList = new ArrayList<>();

            Drawing<?> drawing = getDrawingPatriarch(sheet);
            if (drawing != null) {
                for (Shape shape : drawing) {
                    if (shape instanceof Picture) {
                        Picture picture = (Picture) shape;
                        PictureData pictureData = picture.getPictureData();

                        try {
                            ClientAnchor anchor = picture.getClientAnchor();
                            byte[] data = pictureData.getData();

                            // 获取图片尺寸
                            int width = 0;
                            int height = 0;
                            try (ByteArrayInputStream bis = new ByteArrayInputStream(data)) {
                                BufferedImage bufferedImage = ImageIO.read(bis);
                                if (bufferedImage != null) {
                                    width = bufferedImage.getWidth();
                                    height = bufferedImage.getHeight();
                                }
                            } catch (Exception e) {
                                // 如果无法解析图片大小，使用默认值
                                width = anchor.getDx2() - anchor.getDx1();
                                height = anchor.getDy2() - anchor.getDy1();
                            }

                            pictureList.add(new PictureInfo(
                                    data,
                                    pictureData.getMimeType(),
                                    anchor,
                                    width,
                                    height
                            ));
                        } catch (Exception e) {
                            System.err.println("Error processing picture: " + e.getMessage());
                        }
                    }
                }
            }
            sheetPicturesMap.put(sheet, pictureList);
        }
        return sheetPicturesMap;
    }

    // 复制工作表内容
    private static void copySheetContent(Sheet inputSheet, Sheet outputSheet,
                                         Map<CellStyle, CellStyle> styleCache,
                                         Workbook inputWorkbook, Workbook outputWorkbook) {

        // 复制行
        for (Row inputRow : inputSheet) {
            if (inputRow == null) continue;

            Row outputRow = outputSheet.createRow(inputRow.getRowNum());
            outputRow.setHeight(inputRow.getHeight());

            // 复制单元格
            for (Cell inputCell : inputRow) {
                if (inputCell == null) continue;

                Cell outputCell = outputRow.createCell(inputCell.getColumnIndex());
                copyCell(inputCell, outputCell, styleCache, inputWorkbook, outputWorkbook);
            }
        }

        // 复制合并单元格
        for (int i = 0; i < inputSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = inputSheet.getMergedRegion(i);
            outputSheet.addMergedRegion(mergedRegion);
        }
    }

    // 复制单元格内容和样式
    private static void copyCell(Cell inputCell, Cell outputCell,
                                 Map<CellStyle, CellStyle> styleCache,
                                 Workbook inputWorkbook, Workbook outputWorkbook) {

        // 复制单元格值
        switch (inputCell.getCellType()) {
            case STRING:
                outputCell.setCellValue(inputCell.getStringCellValue());
                break;
            case NUMERIC:
                outputCell.setCellValue(inputCell.getNumericCellValue());
                break;
            case BOOLEAN:
                outputCell.setCellValue(inputCell.getBooleanCellValue());
                break;
            case FORMULA:
                outputCell.setCellFormula(inputCell.getCellFormula());
                break;
            case BLANK:
                outputCell.setBlank();
                break;
        }

        // 解决样式克隆问题
        CellStyle inputStyle = inputCell.getCellStyle();
        CellStyle outputStyle;

        if (styleCache.containsKey(inputStyle)) {
            outputStyle = styleCache.get(inputStyle);
        } else {
            // 创建对应新样式
            outputStyle = outputWorkbook.createCellStyle();
            // 复制样式属性（而不是克隆整个对象）
            copyCellStyle(outputStyle, inputStyle, outputWorkbook, inputWorkbook);
            styleCache.put(inputStyle, outputStyle);
        }
        outputCell.setCellStyle(outputStyle);
    }

    // 手动复制样式属性（避免跨工作簿克隆）
    private static void copyCellStyle(CellStyle targetStyle, CellStyle sourceStyle, Workbook outputWorkbook, Workbook inputWorkbook) {
        targetStyle.setAlignment(sourceStyle.getAlignment());
        targetStyle.setBorderTop(sourceStyle.getBorderTop());
        targetStyle.setBorderBottom(sourceStyle.getBorderBottom());
        targetStyle.setBorderLeft(sourceStyle.getBorderLeft());
        targetStyle.setBorderRight(sourceStyle.getBorderRight());
        targetStyle.setTopBorderColor(sourceStyle.getTopBorderColor());
        targetStyle.setBottomBorderColor(sourceStyle.getBottomBorderColor());
        targetStyle.setLeftBorderColor(sourceStyle.getLeftBorderColor());
        targetStyle.setRightBorderColor(sourceStyle.getRightBorderColor());
        targetStyle.setFillPattern(sourceStyle.getFillPattern());
        targetStyle.setFillForegroundColor(sourceStyle.getFillForegroundColor());
        targetStyle.setFillBackgroundColor(sourceStyle.getFillBackgroundColor());


        // 修复字体设置 - 使用正确的工作簿对象
        Font targetFont = outputWorkbook.createFont();
        Font sourceFont = inputWorkbook.getFontAt(sourceStyle.getFontIndex());
        targetFont.setBold(sourceFont.getBold());
        targetFont.setColor(sourceFont.getColor());
        targetFont.setFontHeight(sourceFont.getFontHeight());
        targetFont.setFontName(sourceFont.getFontName());
        targetFont.setItalic(sourceFont.getItalic());
        targetStyle.setFont(targetFont);
    }

    // 添加图片到工作表
    private static void addPicturesToSheet(List<PictureInfo> pictures,
                                           Sheet sourceSheet, Sheet outputSheet,
                                           Workbook outputWorkbook) {
        if (pictures == null || pictures.isEmpty()) return;

        CreationHelper creationHelper = outputWorkbook.getCreationHelper();
        Drawing<?> drawing = getOrCreateDrawing(outputSheet);

        for (PictureInfo picInfo : pictures) {
            // 创建锚点
            ClientAnchor newAnchor = createNewAnchor(creationHelper, picInfo.anchor, sourceSheet, outputSheet);

            // 添加图片到工作簿
            int pictureType = getImageType(picInfo.mimeType);
            int pictureIndex = outputWorkbook.addPicture(picInfo.imageData, pictureType);

            // 创建图片对象
            Picture picture = drawing.createPicture(newAnchor, pictureIndex);

            // 可选：调整图片大小
            adjustPictureSize(picture, picInfo, sourceSheet, outputSheet);
        }
    }

    // 创建新的锚点
    private static ClientAnchor createNewAnchor(CreationHelper helper, ClientAnchor original,
                                                Sheet sourceSheet, Sheet targetSheet) {
        ClientAnchor newAnchor = helper.createClientAnchor();

        // 基本位置复制
        newAnchor.setCol1(original.getCol1());
        newAnchor.setRow1(original.getRow1());
        newAnchor.setCol2(original.getCol2());
        newAnchor.setRow2(original.getRow2());

        // 默认偏移量
        int dx1 = original.getDx1();
        int dy1 = original.getDy1();
        int dx2 = original.getDx2();
        int dy2 = original.getDy2();

        // 计算缩放比例
        try {
            float rowRatio = getRowHeightRatio(sourceSheet, targetSheet);
            float colRatio = getColumnWidthRatio(original.getCol1(), sourceSheet, targetSheet);
//            float colRatio2 = getColumnWidthRatio(original.getCol2(), sourceSheet, targetSheet);

            // 应用缩放比例
            newAnchor.setDx1((int) (dx1 * colRatio));
            newAnchor.setDy1((int) (dy1 * rowRatio));
            newAnchor.setDx2((int) (dx2 * colRatio));
            newAnchor.setDy2((int) (dy2 * rowRatio));
        } catch (Exception e) {
            // 如果计算失败，使用原始偏移量
            newAnchor.setDx1(dx1);
            newAnchor.setDy1(dy1);
            newAnchor.setDx2(dx2);
            newAnchor.setDy2(dy2);
        }

        return newAnchor;
    }

    // 调整图片大小
    private static void adjustPictureSize(Picture picture, PictureInfo picInfo,
                                          Sheet sourceSheet, Sheet targetSheet) {
        try {
            // 获取源工作表的单元格尺寸
            float sourceColWidth = sourceSheet.getColumnWidthInPixels(picInfo.anchor.getCol1());
            float sourceRowHeight = sourceSheet.getDefaultRowHeightInPoints();

            // 获取目标工作表的单元格尺寸
            float targetColWidth = targetSheet.getColumnWidthInPixels(picInfo.anchor.getCol1());
            float targetRowHeight = targetSheet.getDefaultRowHeightInPoints();

            // 计算缩放比例
            float scaleX = targetColWidth / sourceColWidth;
            float scaleY = targetRowHeight / sourceRowHeight;

            // 应用缩放比例（1.0为100%）
            picture.resize(scaleX, scaleY);
        } catch (Exception e) {
            // 如果无法自动调整大小，保持原始比例
            picture.resize();
        }
    }

    // 获取绘图对象
    private static Drawing<?> getOrCreateDrawing(Sheet sheet) {
        Drawing<?> drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            drawing = sheet.createDrawingPatriarch();
        }
        return drawing;
    }

    // 获取图片类型
    private static int getImageType(String mimeType) {
        if (mimeType == null) return Workbook.PICTURE_TYPE_PNG;
        switch (mimeType.toLowerCase()) {
            case "image/jpeg": return Workbook.PICTURE_TYPE_JPEG;
            case "image/png": return Workbook.PICTURE_TYPE_PNG;
            case "image/emf": return Workbook.PICTURE_TYPE_EMF;
            case "image/wmf": return Workbook.PICTURE_TYPE_WMF;
            case "image/dib": return Workbook.PICTURE_TYPE_DIB;
            case "image/pict": return Workbook.PICTURE_TYPE_PICT;
            default: return Workbook.PICTURE_TYPE_PNG;
        }
    }

    // 跨版本兼容的绘图对象获取方法
    private static Drawing<?> getDrawingPatriarch(Sheet sheet) {
        if (sheet == null) return null;

        try {
            // POI 5.x+
            return sheet.getDrawingPatriarch();
        } catch (NoSuchMethodError ex) {
            // POI 3.x-4.x 兼容处理
            try {
                if (sheet instanceof org.apache.poi.hssf.usermodel.HSSFSheet) {
                    return ((org.apache.poi.hssf.usermodel.HSSFSheet) sheet).getDrawingPatriarch();
                } else {
                    return ((org.apache.poi.xssf.usermodel.XSSFSheet) sheet).getDrawingPatriarch();
                }
            } catch (Exception e) {
                return null;
            }
        }
    }

    // 计算行高比例
    private static float getRowHeightRatio(Sheet sourceSheet, Sheet targetSheet) {
        float sourceRowHeight = sourceSheet.getDefaultRowHeightInPoints();
        if (sourceRowHeight == 0) sourceRowHeight = 15; // 默认行高

        float targetRowHeight = targetSheet.getDefaultRowHeightInPoints();
        if (targetRowHeight == 0) targetRowHeight = 15; // 默认行高

        return targetRowHeight / sourceRowHeight;
    }

    // 计算列宽比例
    private static float getColumnWidthRatio(int columnIndex, Sheet sourceSheet, Sheet targetSheet) {
        int sourceColWidth = sourceSheet.getColumnWidth(columnIndex);
        if (sourceColWidth == 0) sourceColWidth = 2048; // 默认列宽（约8字符）

        int targetColWidth = targetSheet.getColumnWidth(columnIndex);
        if (targetColWidth == 0) targetColWidth = 2048; // 默认列宽

        return (float) targetColWidth / sourceColWidth;
    }

    // GUI 入口方法（可选）
    private void initUI() {
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1200, 600);
        setLocationRelativeTo(null);

        JPanel panel = new JPanel();
        panel.setLayout(new BorderLayout(30, 30));
        panel.setBorder(BorderFactory.createEmptyBorder(60, 60, 60, 60));

        // 输入区域
        JPanel inputPanel = new JPanel();
        inputPanel.setLayout(new BorderLayout(60, 60));
        inputFileField = new JTextField();
        inputFileField.setEditable(false);
        inputPanel.add(new JLabel("输入文件 (.xls):"), BorderLayout.WEST);
        inputPanel.add(inputFileField, BorderLayout.CENTER);

        JButton browseButton = new JButton("浏览...");
        browseButton.addActionListener(this::browseAction);
        inputPanel.add(browseButton, BorderLayout.EAST);

        // 按钮区域
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 30, 30));
        convertButton = new JButton("转换为 XLSX");
        convertButton.setEnabled(false);
        convertButton.addActionListener(this::convertAction);
        buttonPanel.add(convertButton);

        // 进度条
        progressBar = new JProgressBar(0, 600);
        progressBar.setStringPainted(true);
        progressBar.setString("等待操作...");

        // 布局
        panel.add(inputPanel, BorderLayout.NORTH);
        panel.add(buttonPanel, BorderLayout.CENTER);
        panel.add(progressBar, BorderLayout.SOUTH);

        add(panel);
    }

    private void browseAction(ActionEvent e) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel 97-2003 工作簿 (.xls)", "xls"));

        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            inputFileField.setText(selectedFile.getAbsolutePath());
            convertButton.setEnabled(true);
        }
    }

    private void convertAction(ActionEvent e) {
        String inputPath = inputFileField.getText();
        if (inputPath == null || inputPath.isEmpty()) {
            JOptionPane.showMessageDialog(this, "请选择要转换的Excel文件", "错误", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 构造输出文件路径 (xls 替换为 xlsx)
        File inputFile = new File(inputPath);
        String outputPath = inputPath.replace(".xls", ".xlsx");
        if (outputPath.equals(inputPath)) {
            outputPath = inputPath + ".xlsx";
        }

        File outputFile = new File(outputPath);

        // 检查输出文件是否已存在
        if (outputFile.exists()) {
            int choice = JOptionPane.showConfirmDialog(this,
                    "文件 " + outputFile.getName() + " 已存在。是否覆盖？",
                    "文件已存在", JOptionPane.YES_NO_OPTION);
            if (choice != JOptionPane.YES_OPTION) {
                return;
            }
        }

        // 在后台线程中执行转换
        new Thread(() -> {
            SwingUtilities.invokeLater(() -> {
                convertButton.setEnabled(false);
                progressBar.setString("转换中...");
                progressBar.setIndeterminate(true);
            });

            boolean success;

            try {
                convertXlsToXlsx(inputFile, outputFile);
                success = true;
            } catch (Exception exception) {
                exception.printStackTrace();
                success = false;
            }

            boolean finalSuccess = success;
            SwingUtilities.invokeLater(() -> {
                progressBar.setIndeterminate(false);
                if (finalSuccess) {
                    progressBar.setValue(300);
                    progressBar.setString("转换完成!");
                    JOptionPane.showMessageDialog(this,
                            "文件转换成功!\n输出文件: " + outputFile.getAbsolutePath(),
                            "转换完成", JOptionPane.INFORMATION_MESSAGE);
                } else {
                    progressBar.setValue(0);
                    progressBar.setString("转换失败");
                    JOptionPane.showMessageDialog(this,
                            "文件转换失败，请检查输入文件是否正确",
                            "错误", JOptionPane.ERROR_MESSAGE);
                }
                convertButton.setEnabled(true);
            });
        }).start();
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            ExcelConverter1 converter = new ExcelConverter1();
            converter.setVisible(true);
        });

    }
}
