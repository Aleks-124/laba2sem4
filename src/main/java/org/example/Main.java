package org.example;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    private static JTextField textField;

    public static void main(String[] args) {
        JFrame frame = new JFrame("Excel Data Importer");
        frame.setSize(400, 200);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        textField = new JTextField();
        textField.setBounds(100, 100, 200, 30);

        JButton importButton = new JButton("Import Excel File");
        importButton.setBounds(100, 50, 200, 30);

        importButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setCurrentDirectory(new File(System.getProperty("user.home") + "/Desktop"));
                int result = fileChooser.showOpenDialog(null);

                if (result == JFileChooser.APPROVE_OPTION) {
                    File file = fileChooser.getSelectedFile();
                    try {
                        FileInputStream fis = new FileInputStream(file);
                        XSSFWorkbook workbook = new XSSFWorkbook(fis);
                        XSSFSheet selectedSheet = workbook.getSheet(textField.getText());
                        HashMap<String, double[]> commonMap = InputModule.InputModule(selectedSheet);
                        var mapArray = new HashMap<String, HashMap<String, Double>>();
                        for (String key : commonMap.keySet()){
                            System.out.println(key);
                            System.out.println(commonMap.get(key));
                        }
                        for (String key : commonMap.keySet()){
                            mapArray.put(key, new HashMap<String, Double>());
                        }
                         for (String key : commonMap.keySet()){
                             Calculation calculationExample = new Calculation(mapArray.get(key),commonMap.get(key));
                             mapArray.put(key,calculationExample.makeCalculation(mapArray.get(key),commonMap.get(key)));
                         }
                        ExportModule.writeExcelFile(mapArray,Covarianc.getCovariance(commonMap));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        System.out.println("Ошибка при импорте и чтении данных из файла Excel.");
                        String errorMessage = "Ошибка при импорте и чтении данных из файла Excel";

                        // Отображение информационного сообщения
                        JOptionPane.showMessageDialog(null, errorMessage, "Ошибка", JOptionPane.INFORMATION_MESSAGE);
                    }
                }
            }
        });

        frame.add(textField);
        frame.add(importButton);
        frame.setLayout(null);
        frame.setVisible(true);
    }
}
