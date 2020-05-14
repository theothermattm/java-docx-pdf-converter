package com.rdtschools;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * Taken from https://rdtschools.com/how-to-covert-docx-file-to-pdf-using-apache-poi-library-in-java/
 */
public class Docx2PdfConversion {
    public static void main(String[] args) {
        try (InputStream is = new FileInputStream(new File(args[0]));
             OutputStream out = new FileOutputStream(new File(args[1]));) {
            long start = System.currentTimeMillis();
// 1) Load DOCX into XWPFDocument
            XWPFDocument document = new XWPFDocument(is);
// 2) Prepare Pdf options
            PdfOptions options = PdfOptions.create();
// 3) Convert XWPFDocument to Pdf
            PdfConverter.getInstance().convert(document, out, options);
            System.out.println("File converted to a PDF file in :: "
                    + (System.currentTimeMillis() - start) + " milli seconds");
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }
}