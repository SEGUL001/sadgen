package com.company;
/*
 * SADs Generator
 *
 * Usage: in console -> java sadGen {applicationId}
 *
 * for example: java sadGen 239
 *
 */


import java.io.*;
import java.sql.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.List;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;

/**
 * Build a bar chart from a template docx
 */
public class Main {

    public static void main(String[] args) {

        //
        //ARGS -> ID_APP PREPARED_BY
        //

        try {
            File fileToBeRead = new File("SAD_auto_template.docx");
            FileInputStream fileInputStream = new FileInputStream(fileToBeRead);
            XWPFDocument xwpfDocument = new XWPFDocument(fileInputStream);
            XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor(xwpfDocument);

            ResultSet properties = getPropertiesDataSet();

            ResultSet application = getApplicationDataSet((Long.parseLong(args[0])));


            while(application.next()){
                String application_name = application.getString("ApplicationName");
                String application_id = application.getString("ApplicationId");
                while(properties.next()){
                    String property =  properties.getString("column_access");
                    xwpfDocument = replaceText(xwpfDocument, properties.getString("document_reference"),application.getString(property));
                }

                String pattern = "MM-dd-yyyy";
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);

                String date = simpleDateFormat.format(new java.util.Date());


                xwpfDocument = replaceText(xwpfDocument, "$DATE$",date);


                xwpfDocument.write(new FileOutputStream("output\\FES NewCo_SAD_"+application_id+"_"+application_name+"V1.0.docx"));
            }




            //aqui deberia hacer pull de las keys desde un archivo externo.

            //xwpfDocument = replaceText(xwpfDocument, "$APPLICATION_NAME$", application_name);


            /*File fileToBeRead2 = new File("C:\\dev\\sadGen\\"+application_name);
            FileInputStream fileInputStream2 = new FileInputStream(fileToBeRead2);
            XWPFDocument xwpfDocument2 = new XWPFDocument(fileInputStream2);
            XWPFWordExtractor xwpfWordExtractor2 = new XWPFWordExtractor(xwpfDocument2);

            System.out.println(xwpfWordExtractor2.getText());*/
            System.out.println("ready");
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    private static XWPFDocument replaceText(XWPFDocument doc, String findText, String replaceText) {

        // cambia el texto en el header

        for (XWPFHeader header : doc.getHeaderList()) {
            for (IBodyElement bodyElement : header.getBodyElements()) {
                if (bodyElement instanceof XWPFTable){
                    for (XWPFTableRow xwpfTableRow : ((XWPFTable) bodyElement).getRows()) {
                        for (XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()) {
                            for (XWPFParagraph xwpfParagraph : xwpfTableCell.getParagraphs()) {
                                for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {
                                    String text = xwpfRun.getText(0);
                                    if (text != null) {
                                        if (replaceText != null && text.contains(findText)) {
                                            text = text.replace(findText, replaceText);
                                            xwpfRun.setText(text, 0);
                                        }
                                    }

                                }
                            }

                        }
                }
                }
            }
        }



        for(XWPFTable xwpfTable: doc.getTables()){
            List<XWPFTableRow> tables = xwpfTable.getRows();

            if (tables!=null){
                for (XWPFTableRow xwpfTableRow : tables){
                    for(XWPFTableCell xwpfTableCell : xwpfTableRow.getTableCells()){
                        for(XWPFParagraph xwpfParagraph: xwpfTableCell.getParagraphs()){
                            for(XWPFRun xwpfRun: xwpfParagraph.getRuns()){
                                String text = xwpfRun.getText(0);
                                if (text != null){
                                    if( replaceText!=null && text.contains(findText)){
                                        text = text.replace(findText, replaceText);
                                        xwpfRun.setText(text,0);
                                    }
                                    else{
                                        text = text.replace(findText, "");
                                        xwpfRun.setText(text,0);
                                    }
                                }

                            }
                        }

                    }
                }
            }
            //analiza el resto del documento sin tablas y reemplaza las keys


            for (XWPFParagraph p : doc.getParagraphs()) {
                List<XWPFRun> runs = p.getRuns();
                if (runs != null) {
                    for (XWPFRun r : runs) {
                        String text = r.getText(0);
                        if (text != null && text.contains(findText)) {
                            if( replaceText!=null){
                                if(findText.equals("$APPLICATION_VENDOR$"))  text = text.replace(findText, "Vendor: "+replaceText);
                                else text = text.replace(findText, replaceText);
                                r.setText(text,0);
                            }
                            else {
                                text = text.replace(findText, "");
                                r.setText(text,0);
                            }

                        }
                    }
                }
            }
        }
        return doc;
    }

    private static ResultSet getApplicationDataSet(Long appId){


        // variables
        Connection connection = null;
        Statement statement = null;
        ResultSet resultSet = null;

        // Step 1: Loading or
        // registering Oracle JDBC driver class
        try {

            Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        }
        catch(ClassNotFoundException cnfex) {

            System.out.println("Problem in loading or "
                    + "registering MS Access JDBC driver");
            cnfex.printStackTrace();
        }

        // Step 2: Opening database connection
        try {

            String msAccDB = "FESFENOCsharepointdata.accdb";
            String dbURL = "jdbc:ucanaccess://"
                    + msAccDB;

            // Step 2.A: Create and
            // get connection using DriverManager class
            connection = DriverManager.getConnection(dbURL);

            // Step 2.B: Creating JDBC Statement
            statement = connection.createStatement();

            String query = "SELECT * FROM MasterApplicationList where applicationId=?";

            PreparedStatement prp = connection.prepareStatement(query);

            prp.setLong(1, appId);

            resultSet = prp.executeQuery();
            return resultSet;



        }
        catch(SQLException sqlex){
            sqlex.printStackTrace();
        }
        finally {
            // Step 3: Closing database connection
            try {
                if(null != connection) {
                    // cleanup resources, once after processing
                    //resultSet.close();
                    //statement.close();

                    // and then finally close connection
                    connection.close();
                }
            }
            catch (SQLException e) {
                e.printStackTrace();
            }


        }
        return null;
    }
    private static ResultSet getPropertiesDataSet(){


        // variables
        Connection connection = null;
        ResultSet resultSet = null;

        // Step 1: Loading or
        // registering Oracle JDBC driver class
        try {

            Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        }
        catch(ClassNotFoundException cnfex) {

            System.out.println("Problem in loading or "
                    + "registering MS Access JDBC driver");
            cnfex.printStackTrace();
        }

        // Step 2: Opening database connection
        try {

            String msAccDB = "FESFENOCsharepointdata.accdb";
            String dbURL = "jdbc:ucanaccess://"
                    + msAccDB;

            // Step 2.A: Create and
            // get connection using DriverManager class
            connection = DriverManager.getConnection(dbURL);

            // Step 2.B: Creating JDBC Statement
            String query = "SELECT * FROM data_maping";

            Statement prp = connection.createStatement();

            resultSet = prp.executeQuery(query);

            return resultSet;



        }
        catch(SQLException sqlex){
            sqlex.printStackTrace();
        }
        finally {
            // Step 3: Closing database connection
            try {
                if(null != connection) {
                    // cleanup resources, once after processing
                    //resultSet.close();
                    //statement.close();

                    // and then finally close connection
                    connection.close();
                }
            }
            catch (SQLException e) {
                e.printStackTrace();
            }


        }
        return null;
    }
}


