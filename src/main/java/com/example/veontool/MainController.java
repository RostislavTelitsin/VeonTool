package com.example.veontool;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.Node;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;


public class MainController {

    public File reportFile;

    ArrayList<String> regions = new ArrayList<>();

    ArrayList<String> normlBSCs = new ArrayList<>();
    ArrayList<String> normalEgbts = new ArrayList<>();
    ArrayList<String> normalNodeB = new ArrayList<>();

    HashMap<String, String> commertialBSCs = new HashMap<>();
    HashMap<String, String> commertialEgbts = new HashMap<>();
    HashMap<String, String> commertialNodeb = new HashMap<>();
    HashMap<String, String> commertialEnodeb = new HashMap<>();

    HashMap<String, Integer> bscRegionList = new HashMap<>();
    HashMap<String, Integer> egbtsRegionList = new HashMap<>();
    HashMap<String, Integer> nodebRegionList = new HashMap<>();
    HashMap<String, Integer> enodbRegionList = new HashMap<>();

    double x  ;
    double y  ;

    public XSSFWorkbook wbRep;
    public XSSFSheet sheetRepBsc;
    public XSSFSheet sheetRepMain;
    public XSSFSheet sheetBSC;
    public XSSFSheet sheetEgbts;
    public XSSFSheet sheetNodeb;
    public XSSFSheet sheetEnodeb;

    int bscRegionColumnNum=666;
    int neNameColumnNum=666;
    int bscCommentsColumnNum=666;

    String licData;

    @FXML
    void dragged(MouseEvent event) {
        Node node = (Node) event.getSource();
        Stage stage = (Stage) node.getScene().getWindow();

        stage.setX(event.getScreenX()+x);
        stage.setY(event.getScreenY()+y);

    }

    @FXML
    void pressed(MouseEvent event) {
        Node node = (Node) event.getSource();
        Stage stage = (Stage) node.getScene().getWindow();
        x = stage.getX() - event.getScreenX();
        y = stage.getY() - event.getScreenY();
    }



    @FXML
    public Label warningLabel;

    @FXML
    public Label repNameLabel;

    @FXML
    public TextField dataText;


    @FXML
    public ProgressBar progressBar;

    @FXML
    public void chooseFile() {
        FileChooser fileChooser = new FileChooser();
        reportFile = fileChooser.showOpenDialog(null);
        progressBar.setVisible(false);
        if (reportFile != null) {
            if(!reportFile.getName().contains(".xlsx")) { warningLabel.setText("Что ты суёшь? Где XLSX?"); }
            repNameLabel.setText("");
            repNameLabel.setText(reportFile.getName());
        } else {;
            repNameLabel.setText("");
            repNameLabel.setText("не выбран");
        }
    }

    @FXML
    public void makeReport() {


        Task repTask = repThread();

        new Thread(repTask).start();


    }

    @FXML
    private void bbbb () {
        Task tt = t();
        
        new Thread(tt).start();
    }

    public Task t() {
        return new Task() {
            @Override
            protected Object call() throws Exception {
                for (int i = 0 ; i<100; i++) {
                    progressBar.setVisible(true);
                    Double ii = Double.valueOf(i);
                    ii /=100;
                    progressBar.setProgress(ii);
                    Thread.sleep(20);

                }
                return true;
            }
        };
    }
//aa
    public Task repThread() {
        return new Task() {
            @Override
            protected Object call() throws Exception {

                if (reportFile !=null) {


                    Platform.runLater(new Runnable() {
                        @Override public void run() {
                            warningLabel.setText("");
                        }
                    });
                    progressBar.setVisible(true);
                    progressBar.setProgress(0);
                    wbRep = new XSSFWorkbook(reportFile);
                    sheetRepBsc = wbRep.getSheet("BSC list");
                    sheetRepMain = wbRep.getSheet("3 Bars");
                    sheetBSC = wbRep.getSheet("BSC");
                    sheetEgbts = wbRep.getSheet("eGBTS");
                    sheetNodeb = wbRep.getSheet("NodeB");
                    sheetEnodeb = wbRep.getSheet("eNodeB");



                    makeNormalNeList();

                    makeCommertialBscList();


                    HashMap<String , Integer> regionColumNum = new HashMap<>();
                    for (String r: regions) {
                        for (int i = 0; i< sheetRepMain.getRow(0).getLastCellNum(); i++) {
                            try {
                                String currentData = sheetRepMain.getRow(0).getCell(i).getStringCellValue();
                                if(currentData.equals(r)) {
                                    regionColumNum.put(r, i);
                                }
                            }catch (RuntimeException e){}
                        }
                    }


                    for (String r: regions) {
                        int counterBSC   = 0;
                        int counterEgbts = 0;
                        int counterNodeB = 0;
                        int counterEnodeB = 0;

                        for (String k : commertialBSCs.keySet()) if(r.equals(commertialBSCs.get(k)))   {counterBSC++;}
                        for (String k : commertialEgbts.keySet()) if(r.equals(commertialEgbts.get(k))) {counterEgbts++;}
                        for (String k : commertialNodeb.keySet()) if(r.equals(commertialNodeb.get(k))) {counterNodeB++;}
                        for (String k : commertialEnodeb.keySet()) if(r.equals(commertialEnodeb.get(k))) {counterEnodeB++;}

                        egbtsRegionList.put(r, counterEgbts);
                        nodebRegionList.put(r, counterNodeB);
                        bscRegionList.put(r, counterBSC);
                        enodbRegionList.put(r, counterEnodeB);
                    }
                    progressBar.setProgress(.5);
                    for (String r: regions) {

                        XSSFCell cell = sheetRepMain.getRow(2).createCell(regionColumNum.get(r)+3);
                        cell.setCellValue(enodbRegionList.get(r));
                        //cell.setCellType(CellType.STRING);

                        XSSFCell cell1 = sheetRepMain.getRow(3).createCell(regionColumNum.get(r)+3);
                        cell1.setCellValue(bscRegionList.get(r));
                        //cell1.setCellType(CellType.STRING);

                        XSSFCell cell2 = sheetRepMain.getRow(4).createCell(regionColumNum.get(r)+3);
                        cell2.setCellValue(nodebRegionList.get(r));
                        //cell2.setCellType(CellType.STRING);

                        XSSFCell cell3 = sheetRepMain.getRow(5).createCell(regionColumNum.get(r)+3);
                        cell3.setCellValue(egbtsRegionList.get(r));
                        //cell3.setCellType(CellType.STRING);

                    }

                    progressBar.setProgress(.7);

                    System.out.println("буду сохранять");

                    Calendar cal = Calendar.getInstance();
                    cal.add(Calendar.MINUTE, 1);
                    Date currentTimePlusOneMinute = cal.getTime();
                    System.out.println(currentTimePlusOneMinute.toString().substring(11,19).replace(":", ""));

                    String repPath = reportFile.getPath().substring(0, reportFile.getPath().length()-4) + "_" + currentTimePlusOneMinute.toString().substring(11,19).replace(":", "") + "_new.xlsx";

                    try (OutputStream fileOut = new FileOutputStream(repPath)) {

                        wbRep.write(fileOut);


                    }
                    progressBar.setProgress(1);
                    System.out.println("готово");

                } else {
                    Platform.runLater(new Runnable() {
                        @Override
                        public void run() {
                            warningLabel.setText("Не все файлы выбраны");
                        }
                    });
                }
                return true;
            }
        };
    }


    public void makeNormalNeList() throws IOException, InvalidFormatException {
        licData = dataText.getText();
        for (int i = 0; i<=sheetEgbts.getLastRowNum(); i++) {
            Row row = sheetEgbts.getRow(i);
            try {
                Boolean isActivated =row.getCell(4).getStringCellValue().equals("Yes");
                String expDate = row.getCell(3).getStringCellValue();
                Boolean ifNormalLicense =
                        isActivated && (expDate.equals("PERMANENT") || expDate.contains(licData));
                if (ifNormalLicense) {
                    normalEgbts.add(row.getCell(1).getStringCellValue());
                }
            } catch (RuntimeException e){}
        }

        for (int i = 0; i<=sheetNodeb.getLastRowNum(); i++) {
            Row row = sheetNodeb.getRow(i);
            try {
                Boolean isActivated =row.getCell(4).getStringCellValue().equals("Yes");
                String expDate = row.getCell(3).getStringCellValue();
                Boolean ifNormalLicense =
                        isActivated && (expDate.equals("PERMANENT") || expDate.contains(licData));
                if (ifNormalLicense) {
                    normalNodeB.add(row.getCell(1).getStringCellValue());
                }
            } catch (RuntimeException e){}
        }

        for (int i = 0; i<=sheetBSC.getLastRowNum(); i++) {
            Row row = sheetBSC.getRow(i);
            try {
                Boolean isActivated =row.getCell(4).getStringCellValue().equals("Yes");
                String expDate = row.getCell(3).getStringCellValue();
                Boolean ifNormalLicense =
                        isActivated && (expDate.equals("PERMANENT") || expDate.contains(licData));
                if (ifNormalLicense) {
                    normlBSCs.add(row.getCell(1).getStringCellValue());
                }
            } catch (RuntimeException e){}
        }


        for (int i = 0; i<=sheetEnodeb.getLastRowNum(); i++) {
            Row row = sheetEnodeb.getRow(i);
            try {
                String expDate = row.getCell(6).getStringCellValue();
                String neType = row.getCell(2).getStringCellValue();
                Boolean ifNormalLicense = (expDate.equals("PERMANENT") || expDate.contains(licData)) || neType.equals("DBS3900 IBS");
                if (ifNormalLicense) {
                    commertialEnodeb.put(row.getCell(1).getStringCellValue(), row.getCell(0).getStringCellValue());
                }
            } catch (RuntimeException e){}
        }


    }

    public void makeCommertialBscList() {
         for (int i = 0; i< sheetRepBsc.getRow(0).getLastCellNum(); i++) {
             String currentData = sheetRepBsc.getRow(0).getCell(i).getStringCellValue();
             switch (currentData) {
                 case "Home Subnet" : {
                     bscRegionColumnNum=i;
                     break;
                 }
                 case "NE Name" : {
                     neNameColumnNum=i;
                     break;}
                 case "BSC Comments" : {
                     bscCommentsColumnNum=i;
                     break;}
             }
         }

         if (bscRegionColumnNum==666 && neNameColumnNum==666 && bscCommentsColumnNum==666) {
             Platform.runLater(new Runnable() {
                 @Override public void run() {
                     warningLabel.setText("Нет информации в 'BSC list'");
                 }
             });
         } else {
             for (int i = 1; i<= sheetRepBsc.getLastRowNum(); i++) {
                 try {
                     Cell currentRegValue = sheetRepBsc.getRow(i).getCell(bscRegionColumnNum);
                     Cell currentBSCValue = sheetRepBsc.getRow(i).getCell(neNameColumnNum);
                     Cell commentBSCValue = sheetRepBsc.getRow(i).getCell(bscCommentsColumnNum);

                     if (!regions.contains(currentRegValue.getStringCellValue())) regions.add(currentRegValue.getStringCellValue());
                     if (commentBSCValue == null || commentBSCValue.getCellType()== CellType.BLANK) {
                         if (normlBSCs.contains(currentBSCValue.getStringCellValue())) commertialBSCs.put(currentBSCValue.getStringCellValue(), currentRegValue.getStringCellValue());
                         if (normalEgbts.contains(currentBSCValue.getStringCellValue())) commertialEgbts.put(currentBSCValue.getStringCellValue(), currentRegValue.getStringCellValue());
                         if (normalNodeB.contains(currentBSCValue.getStringCellValue())) commertialNodeb.put(currentBSCValue.getStringCellValue(), currentRegValue.getStringCellValue());
                     }
                 }
                 catch (Exception e) {};
             }

             System.out.println("Regions:");
             for (String r:regions) {
                 System.out.println(r);
             }
         }
     }
}