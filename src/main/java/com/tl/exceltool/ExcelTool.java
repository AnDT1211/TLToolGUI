/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */
package com.tl.exceltool;

import com.tl.exceltool.service.ExcelService;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.commons.lang3.StringUtils;
import com.spire.xls.*;

/**
 *
 * @author andt
 */
public class ExcelTool {

    private final static String folderTKCTStep1 = "C:\\Users\\andt\\EC\\EC41_ESS_Andt\\TL\\testTKCT_step1";
    private final static String folderOutStep1 = "C:\\Users\\andt\\EC\\EC41_ESS_Andt\\TL\\outputTKCT_step1";
    private final static String PATH_TEMPLATE_EXCEL_FILE = "C:\\Users\\andt\\EC\\EC41_ESS_Andt\\TL\\template\\template.xlsx";

    private final static String outputFileNameStep1 = "Output_Step1.xlsx";

    private final static String PATH_INPUT_FILE_STEP2 = folderOutStep1 + "/" + outputFileNameStep1;
    private final static String PATH_OUTPUT_FOLDER_STEP2 = folderTKCTStep1;
    
    
    static List<String> sheetNames = new ArrayList<>() {
        {
            add("処理概要");
            add("IF編集仕様");
            add("ファイルレイアウト");
            add("処理概要 (SQL)①");
            add("処理概要 (SQL)②");
            add("処理概要 (SQL)③");
            add("処理概要 (SQL)④");
            add("処理概要 (SQL)⑤");
            add("処理概要 (SQL)⑥");
            add("処理概要 (SQL)⑦");
            add("処理概要 (SQL)⑧");
            add("処理概要 (SQL)⑨");
            add("処理概要 (SQL)⑩");
            add("処理概要 (SQL)⑪");
            add("処理概要 (SQL)⑫");
            add("処理概要 (SQL)⑬");
            add("処理概要 (SQL)⑭");
            add("処理概要 (SQL)⑮");
            add("処理概要 (SQL)⑯");
            add("処理概要 (SQL)⑰");
            add("処理概要 (SQL)⑱");
            add("処理概要 (SQL)⑲");
            add("処理概要 (SQL)⑳");
        }
    };

    /**
     * https://docs.aspose.com/cells/java/re-order-sheets-within-workbook/ <br>
     * reorder sheets
     */
    public static void reorder() throws Exception {

    }

    /**
     * https://www.e-iceblue.com/Tutorials/Java/Spire.XLS-for-Java/Program-Guide/Worksheet/Add-or-remove-worksheet-in-Java.html
     */
    public static void removeSheet() throws Exception {
        //Specify input and output paths
//        String inputFile = "sample.xlsx";
//        String outputFile = "output/AddWorksheet.xlsx";
        String filePath = "C:\\Users\\andt\\EC\\EC41_ESS_Andt\\TL\\testTKCT_step1\\バッチ処理詳細設計書_KYLK_B_62_バンドル申込者情報連係ファイル作成.xlsx";

        //Create a workbook and load a file
        Workbook workbook = new Workbook();

        //Load a sample Excel file
        workbook.loadFromFile(filePath);

        //Get the second worksheet and remove it
        Worksheet sheet1 = workbook.getWorksheets().get(1);
        sheet1.remove();

        //Save the Excel file
        workbook.saveToFile(filePath, ExcelVersion.Version2010);
        workbook.dispose();
    }

    /**
     * https://www.e-iceblue.com/Tutorials/Java/Spire.XLS-for-Java/Program-Guide/Worksheet/Java-rename-Excel-Sheet-and-Set-Tab-Color.html
     *
     * @throws Exception
     */
    public static void renameSheet() throws Exception {
        //Create a Workbook object
        Workbook workbook = new Workbook();

        // Load a sample Excel document
        String filePath = "C:\\Users\\andt\\EC\\EC41_ESS_Andt\\TL\\testTKCT_step1\\バッチ処理詳細設計書_KYLK_B_62_バンドル申込者情報連係ファイル作成.xlsx";
        workbook.loadFromFile(filePath);

        //Get the first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(3);
        //Rename the first worksheet and set its tab color
        worksheet.setName("DataXXX");

        //Save the document to file
        workbook.saveToFile(filePath, ExcelVersion.Version2010);
    }

    /**
     * https://www.e-iceblue.com/Tutorials/Java/Spire.XLS-for-Java/Program-Guide/Worksheet/Java-rename-Excel-Sheet-and-Set-Tab-Color.html
     */
    /*
    "C:\\Users\andt\\EC\\EC41_ESS_Andt\\TL\\testTKCT_step1\\バッチ処理詳細設計書_KYLK_B_55_再リース支払案内書作成.xlsx"
     */
    public static void duplicateSheetInWorkbook() throws Exception {
        //Create a Workbook object
        Workbook workbook = new Workbook();

        //Load the sample Excel file
        String filePath = "C:\\Users\\andt\\EC\\EC41_ESS_Andt\\TL\\testTKCT_step1\\バッチ処理詳細設計書_KYLK_B_62_バンドル申込者情報連係ファイル作成.xlsx";
        workbook.loadFromFile(filePath);

        //Get the first worksheet
        Worksheet originalSheet = workbook.getWorksheets().get(0);

        //Add a new worksheet
        Worksheet newSheet = workbook.getWorksheets().add(originalSheet.getName() + " - Copy");

        //Copy the worksheet to new sheet
        newSheet.copyFrom(originalSheet);

        //Save to file
        workbook.saveToFile(filePath);
    }

    /**
     * https://www.e-iceblue.com/Tutorials/Java/Spire.XLS-for-Java/Program-Guide/Worksheet/Copy-Worksheets-from-one-Workbook-to-another-in-Java.html#1
     */
    public static void copySheetBetweenWorkbook() throws Exception {
        Workbook sourceWorkbook = new Workbook();
        sourceWorkbook.loadFromFile("sample1.xlsx");
        Worksheet srcWorksheet = sourceWorkbook.getWorksheets().get(0);

        //Create a another Workbook
        Workbook targetWorkbook = new Workbook();
        targetWorkbook.loadFromFile("sample2.xlsx");
        Worksheet targetWorksheet = targetWorkbook.getWorksheets().add("added");

        //Copy the first worksheet of sample1 to the new added sheet of sample2
        targetWorksheet.copyFrom(srcWorksheet);
        String outputFile = "output/CopyWorksheet.xlsx";

        //Save the result file
        targetWorkbook.saveToFile(outputFile, ExcelVersion.Version2013);
        sourceWorkbook.dispose();
        targetWorkbook.dispose();
    }
    
    
    public static void main(String[] args) throws Exception {

        step1(folderTKCTStep1, folderOutStep1, outputFileNameStep1);
        step2(PATH_INPUT_FILE_STEP2, PATH_OUTPUT_FOLDER_STEP2);

//        duplicateSheetInWorkbook();
//        renameSheet();
    }

    
    private static void step2(String pathInputFileStep2, String pathOutputFolderStep2) throws Exception {
        Workbook targetWorkbook = new Workbook();
        Workbook sourceWorkbook = new Workbook();
        sourceWorkbook.loadFromFile(pathInputFileStep2);
        
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(pathOutputFolderStep2))) {
            OUTER: for (Path path : stream) {
                String noFile = path.getFileName().toString().substring(0, 3);
                
                INNER: for (int i = 0; i < sourceWorkbook.getWorksheets().size(); i++) {
                    Worksheet sourceSheet = sourceWorkbook.getWorksheets().get(i);
                    String sheetNameSource = sourceSheet.getName();
                    String noFileSheetNameSource = sheetNameSource.substring(0, 3);
                    
                    if (noFile.equals(noFileSheetNameSource)) {
                        targetWorkbook.loadFromFile(path.toString());
                        
                        for (int j = 0; j < targetWorkbook.getWorksheets().size(); j++) {
                            Worksheet SheetTarget = targetWorkbook.getWorksheets().get(j);
                            String sheetNameTarget = SheetTarget.getName();
                            
                            if (sheetNameTarget.replaceAll("\\s", "").strip().equals(sheetNameSource.replaceAll("\\s", "").strip().substring(4).strip())) {
                                SheetTarget.setName(sheetNameTarget + "(JP)");
                                sourceSheet.setName(sourceSheet.getName() + "(VN)");
                                targetWorkbook.getWorksheets().addCopyAfter(sourceSheet, targetWorkbook.getWorksheets().get(j));
                            }
                        }
                        targetWorkbook.saveToFile(path.toString(), ExcelVersion.Version2007);
                    }
                    targetWorkbook.dispose();
                }
            }
        } finally {
            sourceWorkbook.dispose();
        }
    }

    public static void step1(String pathFolderTKCTStep1, String pathFolderOutputStep1, String outputFileNameStep1) throws Exception {
        Workbook targetWorkbook = new Workbook();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(pathFolderTKCTStep1))) {

            //Create a another Workbook
            targetWorkbook.loadFromFile(PATH_TEMPLATE_EXCEL_FILE);

            int noFile = 0;
            for (Path path : stream) {
                noFile++;
                try {
                    String fileNameTKCT = path.getFileName().toString();
                    if (fileNameTKCT.startsWith("~$")) {
                        continue;
                    }
                    if (!Files.isDirectory(path) && path.getFileName().toString().endsWith(".xlsx")) {
                        // TODO here

                        String appendName = String.format("%03d_", noFile);  // following by files

                        /**
                         * Create file output
                         */
                        Workbook sourceWorkbook = new Workbook();
                        sourceWorkbook.loadFromFile(path.toString());

                        for (int i = 0; i < sourceWorkbook.getWorksheets().size(); i++) {
                            Worksheet sourceSheet = sourceWorkbook.getWorksheets().get(i);
                            String sheetName = sourceSheet.getName().replaceAll("\\s", "").strip();
                            if (sheetNames.contains(sheetName)) {
                                Worksheet targetWorksheet = targetWorkbook.getWorksheets().add(appendName + sourceSheet.getName());
                                targetWorksheet.copyFrom(sourceSheet);
                            }
                        }

                        //Save the result file
                        String outputFile = Path.of(folderOutStep1, outputFileNameStep1).toString();
                        targetWorkbook.saveToFile(outputFile, ExcelVersion.Version2007);
                        sourceWorkbook.dispose();

                        /**
                         * Change name of files
                         */
                        ExcelService.changeName(path, appendName);
                    }
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
        } finally {

            targetWorkbook.dispose();
            ExcelService.removeSheetSignature(Path.of(folderOutStep1, outputFileNameStep1));
            ExcelService.removeSheetAtIndex(Path.of(folderOutStep1, outputFileNameStep1), 0);
        }
    }

    public static void step1_temp(String pathFolderTKCTStep1, String pathFolderOutputStep1, String outputFileNameStep1) throws Exception {
        /*
        - sua ten tat ca TKCT o folder temp
        - tao file gom sheet
        - doc file
        
         */
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(pathFolderTKCTStep1))) {
            int noFile = 0;
            for (Path path : stream) {
                noFile++;
                try {
                    String fileNameTKCT = path.getFileName().toString();
                    if (fileNameTKCT.startsWith("~$")) {
                        continue;
                    }
                    if (!Files.isDirectory(path) && path.getFileName().toString().endsWith(".xlsx")) {
                        // TODO here

                        String appendName = String.format("%03d_", noFile);  // following by files
                        /**
                         * Change name of files
                         */
                        // ExcelService.changeName(path, appendName);

                        /**
                         * Create file output
                         */
                        Files.createDirectories(Path.of(pathFolderOutputStep1));

//                        Sheet sheetTmp = null;
//                        try (final Workbook workbook = new XSSFWorkbook(new FileInputStream(path.toFile()))) {
//                            final Iterator<Sheet> sheetIterator = workbook.sheetIterator();
//                            int idxSheet = 0;
////                            Map<Integer, S
//                            while (sheetIterator.hasNext()) {
//                                final Sheet sheet = sheetIterator.next();
//                                final String sheetName = sheet.getSheetName().trim();
//                                String newSheetName = String.format("%03d_", noFile);
//                                if (sheetNames.contains(sheetName)) {
////                                    workbook.setSheetName(idxSheet, newSheetName);
//                                    sheetTmp = sheet;
//                                    break;
//                                }
//                                
//                                
//                                
//                                idxSheet++;
//                            }
//                            workbook.cloneSheet(1);
//                        }
//                        try (FileInputStream file = new FileInputStream(path.toFile());
//                                final XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(file)) {
//                            
//                            XSSFSheet sheet_copy = workbook.cloneSheet(3);
//                            int num = workbook.getSheetIndex(sheet_copy);
//                            workbook.setSheetName(num, sheet_copy.getSheetName() + "X");
//                            file.close();
//                            
//
//                            FileOutputStream outputStream = new FileOutputStream(path.toFile());
//                            workbook.write(outputStream);
//                            workbook.close();
//                            outputStream.close();
//                        }
//                        ExcelService.createOutputFileStep1(Path.of(pathFolderOutputStep1), null);
                    }
                } catch (Exception ex) {
//                    System.out.println("Fail to read: " + fileNameTKCT);
//                    fileNameFailed.add(fileNameTKCT);
                    ex.printStackTrace();
                }
            }
        }
    }

}
