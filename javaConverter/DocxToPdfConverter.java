package javaConverter;

import java.io.*;
import java.nio.file.*;

public class DocxToPdfConverter {
    
    public static byte[] convertDocxToPdf(ByteArrayInputStream docxStream) throws IOException {
        // Step 1: Create temporary DOCX file
        Path tempDocx = Files.createTempFile("input", ".docx");
        Files.write(tempDocx, docxStream.readAllBytes());
        
        // Step 2: Create temporary PDF output file
        Path tempPdf = Paths.get(tempDocx.toString().replace(".docx", ".pdf"));
        
        // Step 3: Convert DOCX to PDF
        String os = System.getProperty("os.name").toLowerCase();
        boolean success = os.contains("win") ? convertUsingWord(tempDocx.toString(), tempPdf.toString())
                : convertUsingLibreOffice(tempDocx.toString(), tempPdf.toString());
        
        if (!success) {
            throw new IOException("Conversion failed.");
        }
        
        // Step 4: Read the converted PDF into a byte array
        byte[] pdfBytes = Files.readAllBytes(tempPdf);
        
        // Step 5: Cleanup
        Files.deleteIfExists(tempDocx);
        Files.deleteIfExists(tempPdf);
        
        return pdfBytes;
    }
    
    private static boolean convertUsingWord(String inputPath, String outputPath) {
        try {
            // Step 1: Generate temporary VBScript
            Path scriptPath = Files.createTempFile("convert", ".vbs");
            String vbsScript = generateVbsScript(inputPath, outputPath);
            Files.write(scriptPath, vbsScript.getBytes());
            
            // Step 2: Run the VBScript
            Process process = new ProcessBuilder("cscript", "//nologo", scriptPath.toString()).start();
            int exitCode = process.waitFor();
            
            // Cleanup script
            Files.deleteIfExists(scriptPath);
            
            return exitCode == 0;
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
            return false;
        }
    }
    
    private static String generateVbsScript(String inputPath, String outputPath) {
        return "Dim word, doc\n" + "Set word = CreateObject(\"Word.Application\")\n" + "word.Visible = False\n"
                + "Set doc = word.Documents.Open(\"" + inputPath.replace("\\", "\\\\") + "\")\n" + "doc.SaveAs2 \""
                + outputPath.replace("\\", "\\\\") + "\", 17 'wdFormatPDF\n" + "doc.Close False\n" + "word.Quit\n"
                + "WScript.Echo \"✅ PDF saved at: " + outputPath.replace("\\", "\\\\") + "\"\n";
    }
    
    private static boolean convertUsingLibreOffice(String inputPath, String outputPath) {
        String outputFolder = new File(outputPath).getParent();
        String[] command = { "soffice", "--headless", "--convert-to", "pdf", "--outdir", outputFolder, inputPath };
        
        try {
            Process process = new ProcessBuilder(command).start();
            int exitCode = process.waitFor();
            return exitCode == 0;
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
            return false;
        }
    }
    
    public static void main(String[] args) throws IOException {
        // Example: Load DOCX from a file into a ByteArrayInputStream
        Path docxPath = Paths.get("/Users/YasinKhan/coding/convert-docx-to-pdf/javaConverter/input.docx");
        ByteArrayInputStream docxStream = new ByteArrayInputStream(Files.readAllBytes(docxPath));
        
        // Convert to PDF
        byte[] pdfBytes = convertDocxToPdf(docxStream);
        
        // Save PDF for testing
        Files.write(Paths.get("/Users/YasinKhan/coding/convert-docx-to-pdf/javaConverter/input.pdf"), pdfBytes);
        System.out.println("✅ PDF saved successfully!");
    }
}