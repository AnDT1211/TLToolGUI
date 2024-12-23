/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.tl.exceltool.service;

import java.io.BufferedReader;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 *
 * @author andt
 */
public class ReadFileService {
    public static String readContentFromFile(Path path) throws Exception {
        StringBuilder content = new StringBuilder();
        try (BufferedReader wBR = Files.newBufferedReader(path, Charset.forName("UTF-8"));) {
            String wLine;
            while ((wLine = wBR.readLine()) != null) {
                content.append(wLine).append("\n");
            }

            wBR.close();
            return content.toString();
        }
    }
}
