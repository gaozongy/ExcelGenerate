package org.example;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class Main {

    private static String targetDirectory = "./docx/result";
    private static LocalDateTime startTime = LocalDateTime.of(2024, 1, 1, 0, 0);
    private static LocalDateTime endTime = LocalDateTime.of(2024, 12, 31, 0, 0);

    public static void main(String[] args) {
        createFile();
        editFile();
    }

    private static void createFile() {

        while (!startTime.isAfter(endTime)) {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yy-MM-dd");
            String newFileName = "template" + startTime.format(formatter) + ".docx";
            startTime = startTime.plusDays(1);
            String sourceFilePath = "./docx/template24-09-17.docx";
            Path sourcePath = Paths.get(sourceFilePath);
            Path targetPath = Paths.get(targetDirectory, newFileName);

            try {
                Files.createDirectories(Paths.get(targetDirectory));
                Files.copy(sourcePath, targetPath, StandardCopyOption.REPLACE_EXISTING);
                replaceTitle(targetPath);
                System.out.println("文件已成功复制并重命名为: " + targetPath);
            } catch (FileAlreadyExistsException e) {
                System.out.println("目标文件已存在，请选择不同的名称。");
            } catch (IOException e) {
                System.out.println("复制文件时出错。");
            }
        }
    }

    private static void replaceTitle(Path filePath) {
        // 要替换的文本
        String targetText = "2024年01月01日";

        String fileName = filePath.getFileName().toString();
        fileName = fileName.replace("template", "").replace(".docx", "");
        String[] array = fileName.split("-");
        String replacementText = "20" + array[0] + "年" + array[1] + "月" + array[2] + "日";

        try (FileInputStream fis = new FileInputStream(filePath.toFile());
             XWPFDocument document = new XWPFDocument(fis)) {

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null) {
                    StringBuilder paragraphText = new StringBuilder();
                    for (XWPFRun run : runs) {
                        paragraphText.append(run.getText(0));
                    }

                    String fullText = paragraphText.toString();
                    if (fullText.contains(targetText)) {
                        fullText = fullText.replace(targetText, replacementText);

                        String fontFamily = runs.get(0).getFontFamily();
                        int fontSize = runs.get(0).getFontSize();

                        int runSize = runs.size();
                        for (int i = runSize - 1; i >= 0; i--) {
                            paragraph.removeRun(i);
                        }

                        // 将替换后的文本重新添加回段落
                        XWPFRun newRun = paragraph.createRun();
                        newRun.setFontFamily(fontFamily);
                        newRun.setFontSize(fontSize);
                        newRun.setText(fullText);
                    }
                }
            }

            // 保存修改后的文档（覆盖原文件）
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                document.write(fos);
            }

            System.out.println("文件已修改: " + filePath.getFileName());
        } catch (IOException e) {
            System.out.println("处理文件时出错: " + filePath.getFileName());
            e.printStackTrace();
        }
    }


    private static void replaceTitle2(Path filePath) {
        // 要替换的文本
        String targetText = "2024年";
        String replacementText = "2023年";

        try (FileInputStream fis = new FileInputStream(filePath.toFile());
             XWPFDocument document = new XWPFDocument(fis)) {

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null) {
                    StringBuilder paragraphText = new StringBuilder();
                    for (XWPFRun run : runs) {
                        paragraphText.append(run.getText(0));
                    }

                    String fullText = paragraphText.toString();
                    if (fullText.contains(targetText)) {
                        fullText = fullText.replace(targetText, replacementText);

                        String fontFamily = runs.get(0).getFontFamily();
                        int fontSize = runs.get(0).getFontSize();

                        int runSize = runs.size();
                        for (int i = runSize - 1; i >= 0; i--) {
                            paragraph.removeRun(i);
                        }

                        // 将替换后的文本重新添加回段落
                        XWPFRun newRun = paragraph.createRun();
                        newRun.setFontFamily(fontFamily);
                        newRun.setFontSize(fontSize);
                        newRun.setText(fullText);
                    }
                }
            }

            // 保存修改后的文档（覆盖原文件）
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                document.write(fos);
            }

            System.out.println("文件已修改: " + filePath.getFileName());
        } catch (Exception e) {
            System.out.println("处理文件时出错: " + filePath.getFileName());
            e.printStackTrace();
        }
    }

    private static void deleteTable(Path path) {
        try (FileInputStream fis = new FileInputStream(path.toFile());
             XWPFDocument document = new XWPFDocument(fis)) {

            // 获取文档中的所有表格
            List<XWPFTable> tables = document.getTables();

            // 假设要处理的表格是第一个表格
            if (!tables.isEmpty()) {
                XWPFTable table = tables.get(0);

                // 获取表格中的所有行
                List<XWPFTableRow> rows = table.getRows();

                // 检查表格是否有行
                if (!rows.isEmpty()) {
                    // 删除最后一行
                    table.removeRow(rows.size() - 1);
                    System.out.println("成功删除最后一行");
                } else {
                    System.out.println("表格没有可删除的行");
                }
            } else {
                System.out.println("文档中没有表格");
            }

            // 保存修改后的文档
            try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
                document.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void replaceValue(Path filePath) {

        System.out.println(filePath);
        Map<String, Integer> replaceText = new HashMap<>();
        replaceText.put("v1v", 1000);
        replaceText.put("v2v", 2000);
        replaceText.put("v3v", 3000);
        replaceText.put("v4v", 4000);
        replaceText.put("v5v", 5000);

        try (FileInputStream fis = new FileInputStream(filePath.toFile());
             XWPFDocument document = new XWPFDocument(fis)) {

            // 遍历文档中的所有表格
            List<XWPFTable> tables = document.getTables();
            for (XWPFTable table : tables) {
                // 遍历每一行
                for (XWPFTableRow row : table.getRows()) {
                    // 遍历每个单元格
                    for (XWPFTableCell cell : row.getTableCells()) {
                        // 遍历单元格中的每个段落
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            List<XWPFRun> runs = paragraph.getRuns();
                            if (runs != null) {
                                StringBuilder cellText = new StringBuilder();
                                for (XWPFRun run : runs) {
                                    cellText.append(run.getText(0));
                                }

                                // 检查是否包含要替换的文本
                                String fullText = cellText.toString();
                                for (String key : replaceText.keySet()) {
                                    if (fullText.contains(key)) {
                                        String fontFamily = runs.get(0).getFontFamily();
                                        int fontSize = runs.get(0).getFontSize();

                                        // 清除原有的 runs
                                        int runSize = runs.size();
                                        for (int i = runSize - 1; i >= 0; i--) {
                                            paragraph.removeRun(i);
                                        }

                                        // 进行替换
                                        double percentage = 0.02;
                                        Random random = new Random();
                                        int originalNumber = replaceText.get(key);
                                        int lowerBound = (int) (originalNumber * (1 - percentage));
                                        int upperBound = (int) (originalNumber * (1 + percentage));
                                        int randomNumber = random.nextInt(upperBound - lowerBound + 1) + lowerBound;
                                        fullText = fullText.replace(key, randomNumber + "");

                                        XWPFRun newRun = paragraph.createRun();
                                        newRun.setFontFamily(fontFamily);
                                        newRun.setFontSize(fontSize);
                                        newRun.setText(fullText);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // 保存修改后的文档
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                document.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void editFile() {
        try {
            Files.walk(Paths.get(targetDirectory))
                    .filter(Files::isRegularFile)
                    .filter(path -> path.toString().endsWith(".docx"))
                    .forEach(path -> replaceValue(path));
            System.out.println("所有文件处理完成！");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}