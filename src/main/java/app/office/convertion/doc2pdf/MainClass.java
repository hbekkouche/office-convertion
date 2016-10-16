package app.office.convertion.doc2pdf;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;


public class MainClass {

    public static void main(String[] args) {
        ClassLoader classLoader = new MainClass().getClass().getClassLoader();
        File file = new File(classLoader.getResource("demo.docx").getFile());
        System.out.println(file.getPath());
        try {
            Converter converter = MainClass.processArguments(file.getPath(), null, null);
            converter.convert();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static Converter processArguments(String inPath, String outPath, String type) throws Exception {
        Converter converter = null;
        boolean shouldShowMessages = true;


        if (inPath == null) {
            throw new IllegalArgumentException();
        }

        if (outPath == null) {
            outPath = changeExtensionToPDF(inPath);
        }


        String lowerCaseInPath = inPath.toLowerCase();

        InputStream inStream = getInFileStream(inPath);
        OutputStream outStream = getOutFileStream(outPath);

        if (type == null) {
            if (lowerCaseInPath.endsWith("doc")) {
                converter = new DocToPDFConverter(inStream, outStream, shouldShowMessages, true);
            } else if (lowerCaseInPath.endsWith("docx")) {
                converter = new DocxToPDFConverter(inStream, outStream, shouldShowMessages, true);
            } else if (lowerCaseInPath.endsWith("ppt")) {
                converter = new PptToPDFConverter(inStream, outStream, shouldShowMessages, true);
            } else if (lowerCaseInPath.endsWith("pptx")) {
                converter = new PptxToPDFConverter(inStream, outStream, shouldShowMessages, true);
            } else if (lowerCaseInPath.endsWith("odt")) {
                converter = new OdtToPDF(inStream, outStream, shouldShowMessages, true);
            } else {
                converter = null;
            }


        } else {

            switch (type) {
                case "doc":
                    converter = new DocToPDFConverter(inStream, outStream, shouldShowMessages, true);
                    break;
                case "docx":
                    converter = new DocxToPDFConverter(inStream, outStream, shouldShowMessages, true);
                    break;
                case "ppt":
                    converter = new PptToPDFConverter(inStream, outStream, shouldShowMessages, true);
                    break;
                case "pptx":
                    converter = new PptxToPDFConverter(inStream, outStream, shouldShowMessages, true);
                    break;
                case "odt":
                    converter = new OdtToPDF(inStream, outStream, shouldShowMessages, true);
                    break;
                default:
                    converter = null;
                    break;

            }


        }


        return converter;

    }


    //From http://stackoverflow.com/questions/941272/how-do-i-trim-a-file-extension-from-a-string-in-java
    public static String changeExtensionToPDF(String originalPath) {

//		String separator = System.getProperty("file.separator");
        String filename = originalPath;

//		// Remove the path upto the filename.
//		int lastSeparatorIndex = originalPath.lastIndexOf(separator);
//		if (lastSeparatorIndex == -1) {
//			filename = originalPath;
//		} else {
//			filename = originalPath.substring(lastSeparatorIndex + 1);
//		}

        // Remove the extension.
        int extensionIndex = filename.lastIndexOf(".");

        String removedExtension;
        if (extensionIndex == -1) {
            removedExtension = filename;
        } else {
            removedExtension = filename.substring(0, extensionIndex);
        }
        String addPDFExtension = removedExtension + ".pdf";

        return addPDFExtension;
    }


    protected static InputStream getInFileStream(String inputFilePath) throws FileNotFoundException {
        File inFile = new File(inputFilePath);
        FileInputStream iStream = new FileInputStream(inFile);
        return iStream;
    }

    protected static OutputStream getOutFileStream(String outputFilePath) throws IOException {
        File outFile = new File(outputFilePath);

        try {
            //Make all directories up to specified
            outFile.getParentFile().mkdirs();
        } catch (NullPointerException e) {
            //Ignore error since it means not parent directories
        }

        outFile.createNewFile();
        FileOutputStream oStream = new FileOutputStream(outFile);
        return oStream;
    }


}
