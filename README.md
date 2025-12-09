# Excel to JSON Converter

A Java application to convert Excel files to JSON format using Aspose Cells 20.12.

## 更新说明
- 2025-12-10: 添加GitHub Actions自动编译功能
- 使用JDK 17编译环境

## Features

- Convert Excel spreadsheets to JSON format
- Support for multiple question types:
  - 单选题 (SINGLE)
  - 多选题 (MULTIPLE)
  - 判断题 (TRUE_FALSE)
  - 简答题 (SHORT)
- Handle up to 6 options per question (A-F)
- Easy to use command-line interface

## Excel Format

The Excel file should follow this format:

| Column 1 | Column 2 | Column 3 | Column 4 | Column 5 | Column 6 | Column 7 | Column 8 | Column 9 | Column 10 |
|----------|----------|----------|----------|----------|----------|----------|----------|----------|-----------|
| Type     | Question | Option A | Option B | Option C | Option D | Option E | Option F |          | Answer    |

### Column Descriptions

1. **Type**: Question type (单选题, 多选题, 判断题, 简答题)
2. **Question**: The question text
3-8. **Options**: Options A-F for multiple choice and true/false questions
9. **Reserved**: Not used
10. **Answer**: The correct answer

### Special Handling

- **判断题**: Uses options A-B, answer is A for true, B for false
- **简答题**: No options, answer is stored directly

## Prerequisites

- Java 8 or higher
- Maven 3.6 or higher
- Aspose Cells 20.12 license (patched version included)

## Setup

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd ExcelToJsonConverter
   ```

2. Install dependencies:
   ```bash
   mvn install
   ```

3. Patch the Aspose Cells library:
   ```bash
   java -cp "target/classes:lib/aspose-cells-20.12.jar:lib/javassist-3.28.0-GA.jar" com.exceltoconverter.AsposeCells_20_12
   ```

## Usage

### Command Line

```bash
java -jar target/excel-to-json-converter-1.0-SNAPSHOT-jar-with-dependencies.jar <input-excel-file> <output-json-file>
```

### Example

```bash
java -jar target/excel-to-json-converter-1.0-SNAPSHOT-jar-with-dependencies.jar questions.xlsx questions.json
```

## Output Format

The output JSON file will have the following format:

```json
{
  "questions": [
    {
      "id": 1234567890123,
      "type": "SINGLE",
      "question": "This is a sample question?",
      "options": [
        "Option A",
        "Option B",
        "Option C",
        "Option D"
      ],
      "answer": "A"
    },
    {
      "id": 1234567890124,
      "type": "TRUE_FALSE",
      "question": "This is a true/false question?",
      "options": [
        "正确",
        "错误"
      ],
      "answer": "TRUE"
    },
    {
      "id": 1234567890125,
      "type": "SHORT",
      "question": "This is a short answer question?",
      "answer": "This is the answer"
    }
  ]
}
```

## Build

To build the project:

```bash
mvn package
```

This will create two JAR files in the `target` directory:
- `excel-to-json-converter-1.0-SNAPSHOT.jar` - Basic JAR without dependencies
- `excel-to-json-converter-1.0-SNAPSHOT-jar-with-dependencies.jar` - Fat JAR with all dependencies included

## License

This project uses Aspose Cells 20.12 with a patched license.

## Disclaimer

This tool is intended for educational and personal use only. Please ensure you have the appropriate licenses for Aspose Cells before using it in a production environment.
