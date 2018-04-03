import java.io.File;
import java.io.FileOutputStream;
 
import org.apache.poi.xwpf.usermodel.XWPFDocument;
 
public class GenerateDoc {
 
	public static void main(String args[]) {
 
		XWPFDocument document = null;
		FileOutputStream fileOutputStream = null;
		try {
 
			document = new XWPFDocument();
			File fileToBeCreated = new File("C:\\kscodes_temp\\FirstWordFile.docx");
			fileOutputStream = new FileOutputStream(fileToBeCreated);
			document.write(fileOutputStream);
 
			System.out.println("Word Document Created Successfully !!!");
 
		} catch (Exception e) {
			System.out.println("We had an error while creating the Word Doc");
		} finally {
			try {
				if (document != null) {
					document.close();
				}
				if (fileOutputStream != null) {
					fileOutputStream.close();
				}
			} catch (Exception ex) {
			}
		}
 
	}
}