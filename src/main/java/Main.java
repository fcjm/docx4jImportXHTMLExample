import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class Main {

    /*
    ex1 - example with styles
    ex2 - example without styles
    ex3 - example without styles and one column (works)
    ex4 - example without styles and <html> + <body> tags
    ex5 - example without styles and without <p> tags within <td> tags
     */

    public static void main(String[] args) throws IOException, Docx4JException {
        int numExamples = new File("src/main/resources/input").list().length;
        String[] examples = new String[numExamples];

        // Read examples
        for (int i = 0; i < numExamples; i++) {
            File file = new File("src/main/resources/input/ex"+ (i+1) +".xhtml");
            String pathString = file.getPath();
            Path path = Path.of(pathString);

            examples[i] = Files.readString(path);
        }

        // Tranform example .xhtml files to .docx files
        for (int i = 0; i < numExamples; i++) {
            try (
                    ByteArrayOutputStream os = new ByteArrayOutputStream();
                    OutputStream output = new FileOutputStream("src/main/resources/output/ex"+ (i+1) +".docx");
            ) {
                // create DOCX
                WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

                // XHTML --> DOCX
                XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);
                wordMLPackage.getMainDocumentPart().getContent()
                        .addAll(XHTMLImporter.convert(examples[i], null));

                // write DOCX
                wordMLPackage.save(os);
                os.writeTo(output);
            }
        }
    }
}
