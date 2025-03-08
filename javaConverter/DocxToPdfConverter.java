package javaConverter;

import java.io.*;
import java.nio.file.*;

public class DocxToPdfConverter {
    
    public static void convertDocxToPdf(String inputPath, String outputFolder) {
        
        Path input = Paths.get(inputPath).toAbsolutePath();
        
        if (!Files.isRegularFile(input)) {
            System.out.println("❌ Error: The file '" + input + "' does not exist.");
            return;
        }
        
        if (outputFolder == null || outputFolder.isEmpty()) {
            outputFolder = Paths.get("").toAbsolutePath().toString();
        } else {
            outputFolder = Paths.get(outputFolder).toAbsolutePath().toString();
            
        }
        
        Path outputDir = Paths.get(outputFolder);
        
        if (!Files.isDirectory(outputDir)) {
            System.out.println("❌ Error: The output folder '" + outputFolder + "' does not exist.");
            return;
        }
        
        String[] command = { "soffice", "--headless", "--convert-to", "pdf", "--outdir", outputFolder,
                        input.toString() };
        
        try {
            ProcessBuilder processBuilder = new ProcessBuilder(command);
            processBuilder.redirectErrorStream(true);
            Process process = processBuilder.start();
            int exitCode = process.waitFor();
            
            if (exitCode == 0) {
                System.out.println("✅ PDF saved in: " + outputFolder);
            } else {
                System.out.println("❌ Error: Conversion failed.");
                try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()))) {
                    String line;
                    while ((line = reader.readLine()) != null) {
                        System.out.println(line);
                    }
                }
            }
        } catch (IOException | InterruptedException e) {
            System.out.println("❌ Error: An error occurred while executing the command.");
            e.printStackTrace();
        }
    }
    
    public static void main(String[] args) {
        convertDocxToPdf("javaConverter/input.docx", "/Users/YasinKhan/coding/convert-docx-to-pdf/javaConverter");
    }
}