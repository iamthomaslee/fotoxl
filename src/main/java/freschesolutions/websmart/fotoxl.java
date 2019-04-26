package freschesolutions.websmart;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.Date;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Node;
import org.w3c.dom.Element;

public class fotoxl {
    public static void main(String[]args) {
        Date a = new Date();
        long aT = a.getTime();
        try {
            runTest(args[0], args[1]);
        } catch (Exception e) {
            e.printStackTrace();
        }
        Date b = new Date();
        long bT = b.getTime();
        long msBetween = Math.abs(bT-aT);
        System.out.println(msBetween+"ms");
    }

    private static void runTest(String inputFileName, String outputFileName) throws IOException {
        Workbook wb = new XSSFWorkbook();

        Sheet sheet = wb.createSheet("new sheet");

        // create header row
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("Student No");
        row.createCell(1).setCellValue("First Name");
        row.createCell(2).setCellValue("Last Name");
        row.createCell(3).setCellValue("Nick Name");
        row.createCell(4).setCellValue("Mark");

        try {
            File inputFile = new File(inputFileName);
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder;

            dBuilder = dbFactory.newDocumentBuilder();

            Document doc = dBuilder.parse(inputFile);
            doc.getDocumentElement().normalize();

            XPath xPath =  XPathFactory.newInstance().newXPath();

            String expression = "/class/student";
            NodeList nodeList = (NodeList) xPath.compile(expression).evaluate(
                    doc, XPathConstants.NODESET);

            for (int i = 0; i < nodeList.getLength(); i++) {
                row = sheet.createRow(i+1);
                Node nNode = nodeList.item(i);
                if (nNode.getNodeType() == Node.ELEMENT_NODE) {
                    Element eElement = (Element) nNode;
                    row.createCell(0).setCellValue(eElement.getAttribute("rollno"));
                    row.createCell(1).setCellValue(eElement.getElementsByTagName("firstname").item(0).getTextContent());
                    row.createCell(2).setCellValue(eElement.getElementsByTagName("lastname").item(0).getTextContent());
                    row.createCell(3).setCellValue(eElement.getElementsByTagName("nickname").item(0).getTextContent());
                    row.createCell(4).setCellValue(eElement.getElementsByTagName("marks").item(0).getTextContent());
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Write the output to a file
        try (OutputStream fileOut = new FileOutputStream(outputFileName)) {
            wb.write(fileOut);
        }
    }
}
