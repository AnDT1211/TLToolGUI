package com.tl.exceltool.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author andt
 */
public class ExcelService {

    /**
     * sua ten file bang cach append vao dau ten file <br>
     * path = 'abc.xlsx', appendFirst = '01' => '01abc.xlsx'
     *
     * @param path : path cua file muon sua ten
     * @param appendFirst : string de append vao dau
     */
    public static void changeName(Path path, String appendFirst) throws Exception {
        // rename
        String newName = appendFirst + path.getFileName();
        Files.move(path, Path.of(path.getParent().toString(), newName), StandardCopyOption.REPLACE_EXISTING);
    }
    
    /**
     * sua ten file bang cach xoa '001_' o dau <br>
     * path = '001_abc.xlsx' => 'abc.xlsx'
     *
     * @param path : path cua file muon sua ten
     * @param appendFirst : string de append vao dau
     */
    public static void changeNameToVN(Path path) throws Exception {
        // rename
        String fileName = path.getFileName().toString();  // abcd.xlsx
        String newName = fileName.substring(4, fileName.length() - 5) + "_VN.xlsx";
        Files.move(path, Path.of(path.getParent().toString(), newName), StandardCopyOption.REPLACE_EXISTING);
    }

    public static void createOutputFileStep1(Path newFilePath, Workbook workbook) throws Exception {
        try (workbook) {
            FileOutputStream outputStream = new FileOutputStream(newFilePath.toFile());
            workbook.write(outputStream);
        }
    }

    public static void removeSheetSignature(Path filePath) throws Exception {
        final String SIGNATURE_SHEET_NAME = "Evaluation Warning";
        // Open file
        try (final XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(filePath.toFile()))) {

            for (int i = workbook.getNumberOfSheets() - 1; i >= 0; i--) {
                XSSFSheet tmpSheet = workbook.getSheetAt(i);
                if (tmpSheet.getSheetName().equals(SIGNATURE_SHEET_NAME)) {
                    workbook.removeSheetAt(i);
                }
            }

            // Save the file
            try (FileOutputStream outFile = new FileOutputStream(filePath.toFile())) {
                workbook.write(outFile);
            }
        }
    }
    
    public static void removeSheetAtIndex(Path filePath, int idx) throws Exception {
        try (final XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(filePath.toFile()))) {

            workbook.removeSheetAt(idx);

            // Save the file
            try (FileOutputStream outFile = new FileOutputStream(filePath.toFile())) {
                workbook.write(outFile);
            }
        }
    }

}
