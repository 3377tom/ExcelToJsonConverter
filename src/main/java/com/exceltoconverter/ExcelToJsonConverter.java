package com.exceltoconverter;

import com.aspose.cells.Cell;
import com.aspose.cells.Row;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class ExcelToJsonConverter {
    public static void main(String[] args) {
        try {
            // 初始化Aspose Cells许可证
            initAsposeLicense();
            
            // 检查命令行参数
            if (args.length < 2) {
                System.out.println("Usage: java -jar ExcelToJsonConverter.jar <input-excel-file> <output-json-file>");
                return;
            }
            
            String inputFile = args[0];
            String outputFile = args[1];
            
            // 转换Excel到JSON
            convertExcelToJson(inputFile, outputFile);
            
            System.out.println("Conversion completed successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    /**
     * 初始化Aspose Cells许可证
     */
    private static void initAsposeLicense() {
        try {
            com.aspose.cells.License license = new com.aspose.cells.License();
            license.setLicense(new java.io.StringReader("<License> <Data> <Products> <Product>Aspose.Cells for Java</Product> </Products> <EditionType>Enterprise</EditionType> <SubscriptionExpiry>29991231</SubscriptionExpiry> <LicenseExpiry>29991231</LicenseExpiry> <SerialNumber>evilrule</SerialNumber> </Data> <Signature>evilrule</Signature> </License>"));
        } catch (Exception e) {
            System.out.println("License initialization failed, but continuing...");
        }
    }
    
    /**
     * 转换Excel文件到JSON格式
     */
    public static void convertExcelToJson(String inputFile, String outputFile) throws Exception {
        // 加载Excel文件
        Workbook workbook = new Workbook(inputFile);
        Worksheet worksheet = workbook.getWorksheets().get(0);
        
        // 获取总行数
        int totalRows = worksheet.getCells().getMaxDataRow() + 1;
        
        // 创建JSON数组存储所有题目
        JSONArray questionsArray = new JSONArray();
        
        // 遍历每一行（跳过标题行）
        for (int rowIndex = 1; rowIndex < totalRows; rowIndex++) {
            Row row = worksheet.getRows().get(rowIndex);
            
            // 解析题目信息
            JSONObject questionObj = parseQuestionRow(row);
            if (questionObj != null) {
                questionsArray.put(questionObj);
            }
        }
        
        // 创建最终的JSON对象
        JSONObject result = new JSONObject();
        result.put("questions", questionsArray);
        
        // 写入JSON文件
        java.io.FileWriter writer = new java.io.FileWriter(outputFile);
        writer.write(result.toString(2));
        writer.close();
        
        System.out.println("Successfully converted " + questionsArray.length() + " questions.");
    }
    
    /**
     * 解析一行Excel数据为题目JSON对象
     */
    private static JSONObject parseQuestionRow(Row row) {
        try {
            // 第一列：题型
            Cell typeCell = row.getCells().get(0);
            String type = typeCell.getStringValue().trim();
            
            // 第二列：题干
            Cell questionCell = row.getCells().get(1);
            String question = questionCell.getStringValue().trim();
            
            if (question.isEmpty()) {
                return null; // 跳过空行
            }
            
            // 第十列：答案
            Cell answerCell = row.getCells().get(9);
            String answer = answerCell.getStringValue().trim();
            
            // 创建题目对象
            JSONObject questionObj = new JSONObject();
            questionObj.put("id", System.currentTimeMillis() + row.getRowIndex()); // 临时ID
            
            // 设置题型
            String questionType = "";
            switch (type.toLowerCase()) {
                case "单选题":
                case "single":
                    questionType = "SINGLE";
                    break;
                case "多选题":
                case "multiple":
                    questionType = "MULTIPLE";
                    break;
                case "判断题":
                case "true_false":
                    questionType = "TRUE_FALSE";
                    break;
                case "简答题":
                case "short":
                    questionType = "SHORT";
                    break;
                default:
                    questionType = "SINGLE"; // 默认单选题
            }
            questionObj.put("type", questionType);
            questionObj.put("question", question);
            
            // 处理选项（第三到第九列：A-F选项）
            List<String> options = new ArrayList<>();
            if (!questionType.equals("SHORT")) { // 简答题没有选项
                for (int i = 2; i <= 8; i++) { // 第三列到第九列对应索引2-8
                    Cell optionCell = row.getCells().get(i);
                    String option = optionCell.getStringValue().trim();
                    if (!option.isEmpty()) {
                        options.add(option);
                    }
                }
                
                // 判断题特殊处理：如果没有选项，添加默认选项
                if (questionType.equals("TRUE_FALSE") && options.isEmpty()) {
                    options.add("正确");
                    options.add("错误");
                }
                
                // 设置选项
                if (!options.isEmpty()) {
                    questionObj.put("options", options);
                }
            }
            
            // 处理答案
            if (questionType.equals("TRUE_FALSE")) {
                // 判断题答案转换：a→TRUE, b→FALSE
                if (answer.equalsIgnoreCase("a")) {
                    questionObj.put("answer", "TRUE");
                } else if (answer.equalsIgnoreCase("b")) {
                    questionObj.put("answer", "FALSE");
                } else {
                    questionObj.put("answer", answer);
                }
            } else {
                questionObj.put("answer", answer);
            }
            
            return questionObj;
        } catch (Exception e) {
            System.out.println("Error parsing row " + row.getRowIndex() + ": " + e.getMessage());
            return null;
        }
    }
}