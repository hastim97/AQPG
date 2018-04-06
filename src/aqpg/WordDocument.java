/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package aqpg;


import com.itextpdf.text.Chunk;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import java.util.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Random;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;


/**
 *
 * @author Hasti Mehta
 */
public class WordDocument {
    
    
    static String branch,year,sem,subject,level;
    static int mModule,marks;
    static ArrayList<String> randQuestions=new ArrayList<String>();
    static int count=1;
    static int filecount=1;
    Connection cn;
    Statement st;
    static Random randGenerator;
    
    public WordDocument(String branch,String year,String sem,String subject,int marks,String level){
        this.branch=branch;
        this.year=year;
        this.sem=sem;
        this.subject=subject;
        this.marks=marks;
        this.level=level;
        randGenerator=new Random();
        count=1;
        randQuestions.clear();
        try{
            Class.forName("com.mysql.jdbc.Driver");
            cn=(com.mysql.jdbc.Connection) DriverManager.getConnection("jdbc:mysql://localhost:3306/SEProj?zeroDateTimeBehavior=convertToNull","Hasti","hasti");
            st=(com.mysql.jdbc.Statement) cn.createStatement();
            JOptionPane.showMessageDialog(null,"Connected");
            String query="SELECT question from questions where sem='"+sem+"' and year='"+year+"' and branch='"+branch+"' and subject='"+subject+"' and level='"+level+"'";
            ResultSet rs=st.executeQuery(query);
            while(rs.next()){
                try{
                    String q=rs.getString("question");
                    randQuestions.add(q);           
                }
                catch(Exception e){}
            }
            System.out.println(randQuestions);
        }
        catch(Exception e){
            System.out.println(e);
        }
        
        //Blank Document
        XWPFDocument document = new XWPFDocument(); 
        
        //Write the Document in file system
        try{
            FileOutputStream out = new FileOutputStream(new File(subject+" Question Paper.docx"));
           
            //create paragraph
            XWPFParagraph paragraph1 = document.createParagraph();
            paragraph1.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = paragraph1.createRun();
            run.setBold(true);
            run.setFontSize(20);
            run.setText(subject+" \t\t Total Marks: "+marks);
            
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run1 = paragraph.createRun();
            run1.setBold(true);
            run1.setFontSize(16);
            run1.setText("Q"+count+". Answer the following Questions (5 marks each)");
            int i=marks;
            XWPFParagraph para=document.createParagraph();
            XWPFRun run2=para.createRun();
            run2.setBold(false);
            run2.setFontSize(12);
            String generatedQuestions="";
            while(i>0){
                int index = randGenerator.nextInt(randQuestions.size());
                run2.setText("Q"+count+". "+randQuestions.get(index));
                generatedQuestions=generatedQuestions+","+randQuestions.get(index);
                run2.addBreak();
                randQuestions.remove(index);
                System.out.println("Index :"+index);
                i=i-5;
                count++;
            }
            document.write(out);
            out.close();
            System.out.println(".docx written successully");
            JOptionPane.showMessageDialog(null,"Word Document created Successfully!");
            String query="INSERT into past_questions (branch,sem,year,subject,marks,questions) values ('"+branch+"','"+sem+"','"+year+"','"+subject+"',"+marks+",'"+generatedQuestions+"')";
            st.executeUpdate(query);
            createPDF();
        }
        catch(Exception e){
            System.out.println(e);
        }
    }
    
    public void createPDF() throws FileNotFoundException, DocumentException{
        String ext = FilenameUtils.getExtension("C:\\Users\\Hasti Mehta\\Documents\\NetBeansProjects\\AQPG\\"+subject+" Question Paper.docx");
        String output = "";
        if ("docx".equalsIgnoreCase(ext)) {
            output = readDocxFile("C:\\Users\\Hasti Mehta\\Documents\\NetBeansProjects\\AQPG\\"+subject+" Question Paper.docx");
        } else if ("doc".equalsIgnoreCase(ext)) {
            output = readDocFile("C:\\Users\\Hasti Mehta\\Documents\\NetBeansProjects\\AQPG\\"+subject+" Question Paper.doc");
        } else {
            System.out.println("INVALID FILE TYPE. ONLY .doc and .docx are permitted.");
        }
        writePdfFile(output); 
        JOptionPane.showMessageDialog(null,"PDF Created!");
    }
    
    public static String readDocFile(String fileName) {
        String output = "";
        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            HWPFDocument doc = new HWPFDocument(fis);
            WordExtractor we = new WordExtractor(doc);
            String[] paragraphs = we.getParagraphText();
            for (String para : paragraphs) {
                output = output + "\n" + para.toString() + "\n";
            }
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return output;
    }

    public static String readDocxFile(String fileName) {
        String output = "";
        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            XWPFDocument document = new XWPFDocument(fis);
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (XWPFParagraph para : paragraphs) {
                output = output + "\n" + para.getText() + "\n";
            }
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return output;
    }
    
    public static void writePdfFile(String output) throws FileNotFoundException, DocumentException {
        File file = new File("C:\\Users\\Hasti Mehta\\Documents\\NetBeansProjects\\AQPG\\"+subject+" Question Paper.pdf");
        FileOutputStream fileout = new FileOutputStream(file);
        Document document = new Document();
        PdfWriter.getInstance(document, fileout);
        document.addTitle("My Converted PDF");
        document.open();
        String[] splitter = output.split("\\n");
        for (int i = 0; i < splitter.length; i++) {
            Chunk chunk = new Chunk(splitter[i]);
            document.add(chunk);
            Paragraph paragraph = new Paragraph();
            paragraph.add("");
            document.add(paragraph);
        }
        document.close();

    }
    
    public static void main(String[] args)throws Exception  {
   
   }
}
