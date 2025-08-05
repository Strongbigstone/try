package org.example;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.*;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelConverter0 extends JFrame {

    private JTextField inputFileField;
    private JButton convertButton;
    private JProgressBar progressBar;

    public ExcelConverter0() {
        super("Excel 文件转换工具");
        initUI();
    }

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

            boolean success = convertExcel(inputFile, outputFile);

            SwingUtilities.invokeLater(() -> {
                progressBar.setIndeterminate(false);
                if (success) {
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

    private boolean convertExcel(File inputFile, File outputFile) {
        try (InputStream in = new FileInputStream(inputFile);
             Workbook inputWorkbook = new HSSFWorkbook(in)) {

            try (Workbook outputWorkbook = new XSSFWorkbook()) {
                // 创建样式映射表
                Map<CellStyle, CellStyle> styleMap = new HashMap<>();

                int sheetCount = inputWorkbook.getNumberOfSheets();
                for (int i = 0; i < sheetCount; i++) {
                    Sheet inputSheet = inputWorkbook.getSheetAt(i);
                    Sheet outputSheet = outputWorkbook.createSheet(inputSheet.getSheetName());

                    int rowCount = inputSheet.getLastRowNum();
                    for (int j = 0; j <= rowCount; j++) {
                        Row inputRow = inputSheet.getRow(j);
                        if (inputRow == null) continue;

                        Row outputRow = outputSheet.createRow(j);

                        int cellCount = inputRow.getLastCellNum();
                        for (int k = 0; k < cellCount; k++) {
                            Cell inputCell = inputRow.getCell(k);
                            if (inputCell == null) continue;

                            Cell outputCell = outputRow.createCell(k, inputCell.getCellType());

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

                            if (styleMap.containsKey(inputStyle)) {
                                outputStyle = styleMap.get(inputStyle);
                            } else {
                                // 创建对应新样式
                                outputStyle = outputWorkbook.createCellStyle();
                                // 复制样式属性（而不是克隆整个对象）
                                copyCellStyle(outputStyle, inputStyle, outputWorkbook, inputWorkbook);
                                styleMap.put(inputStyle, outputStyle);
                            }

                            outputCell.setCellStyle(outputStyle);
                        }
                    }
                }

                try (FileOutputStream out = new FileOutputStream(outputFile)) {
                    outputWorkbook.write(out);
                }

                return true;
            }
        } catch (Exception ex) {
            ex.printStackTrace();
            return false;
        }
    }

    // 手动复制样式属性（避免跨工作簿克隆）
    private void copyCellStyle(CellStyle targetStyle, CellStyle sourceStyle, Workbook outputWorkbook, Workbook inputWorkbook) {
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


    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            ExcelConverter0 converter = new ExcelConverter0();
            converter.setVisible(true);
        });

    }
}