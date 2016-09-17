package wordtopdf;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.function.BooleanSupplier;

import java.io.OutputStream;

import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class Main {

    public static void main(String[] args) {

        try {

            boolean saveFO = true;

            String inputfilepath = "D:\\test.docx";


            String regex = null;
            PhysicalFonts.setRegex(regex);

            WordprocessingMLPackage wordMLPackage;

            System.out.print("Initializing the XMLPackage");

            wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputfilepath));

            // Refresh the values of DOCPROPERTY fields
            FieldUpdater updater = new FieldUpdater(wordMLPackage);
            updater.update(true);

            String outputfilepath = inputfilepath + ".pdf";

            // All methods write to an output stream
            OutputStream os = new java.io.FileOutputStream(outputfilepath);

// Set up font mapper (optional)
            //Mapper fontMapper = new IdentityPlusMapper();
            // wordMLPackage.setFontMapper(fontMapper);


            //PhysicalFont font
            //        = PhysicalFonts.get("Arial Unicode MS");

            FOSettings foSettings = Docx4J.createFOSettings();
            if (saveFO) {
                foSettings.setFoDumpFile(new java.io.File(inputfilepath + ".fo"));
            }
            foSettings.setWmlPackage(wordMLPackage);

            // Don't care what type of exporter you use
            Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);

            System.out.println("Saved: " + outputfilepath);

            // Clean up, so any ObfuscatedFontPart temp files can be deleted
            if (wordMLPackage.getMainDocumentPart().getFontTablePart()!=null) {
                wordMLPackage.getMainDocumentPart().getFontTablePart().deleteEmbeddedFontTempFiles();
            }
            // This would also do it, via finalize() methods
            updater = null;
            foSettings = null;
            wordMLPackage = null;



        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
