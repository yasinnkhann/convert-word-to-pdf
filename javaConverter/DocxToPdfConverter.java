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
        
        String os = System.getProperty("os.name").toLowerCase();
        
        if (os.contains("win")) {
            // Use Microsoft Word on Windows
            convertUsingWord(input.toString(), outputFolder);
        } else {
            // Use LibreOffice on macOS/Linux
            convertUsingLibreOffice(input.toString(), outputFolder);
        }
    }
    
    private static void convertUsingWord(String inputPath, String outputFolder) {
        String vbsScript = outputFolder + "\\convert.vbs";
        String outputFilePath = outputFolder + "\\"
                + Paths.get(inputPath).getFileName().toString().replace(".docx", ".pdf");
        
        String scriptContent = "Dim word\n" + "Set word = CreateObject(\"Word.Application\")\n"
                + "word.Visible = False\n" + "Dim doc\n" + "Set doc = word.Documents.Open(\"" + inputPath + "\")\n"
                + "doc.SaveAs \"" + outputFilePath + "\", 17\n" + "doc.Close\n" + "word.Quit";
        
        try {
            Files.write(Paths.get(vbsScript), scriptContent.getBytes());
            ProcessBuilder processBuilder = new ProcessBuilder("cscript", vbsScript);
            processBuilder.redirectErrorStream(true);
            Process process = processBuilder.start();
            int exitCode = process.waitFor();
            
            Files.delete(Paths.get(vbsScript));
            
            if (exitCode == 0) {
                System.out.println("✅ PDF saved in: " + outputFolder);
            } else {
                System.out.println("❌ Error: Conversion failed.");
            }
        } catch (IOException | InterruptedException e) {
            System.out.println("❌ Error: Could not convert using Microsoft Word.");
            e.printStackTrace();
        }
    }
    
    private static void convertUsingLibreOffice(String inputPath, String outputFolder) {
        String[] command = { "soffice", "--headless", "--convert-to", "pdf", "--outdir", outputFolder, inputPath };
        
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
            System.out.println("❌ Error: Could not convert using LibreOffice.");
            e.printStackTrace();
        }
    }
    
    public static void main(String[] args) {
        convertDocxToPdf("javaConverter/input.docx", "/Users/YasinKhan/coding/convert-docx-to-pdf/javaConverter");
    }
}