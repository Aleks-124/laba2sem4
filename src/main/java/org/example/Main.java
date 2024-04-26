package org.example;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {


    public static void main(String[] args) {
        JTextField textField;
        JFrame frame = new JFrame("Excel Data Importer");
        frame.setSize(400, 200);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        textField = new JTextField();
        textField.setBounds(100, 100, 200, 30);

        JButton importButton = new JButton("Import Excel File");
        importButton.setBounds(100, 30, 200, 30);
        JButton exportButton = new JButton("Экспортировать файл");
        exportButton.setBounds(100, 65, 200, 30);

        // Добавление кнопки в окно
        frame.add(exportButton);

        
        importButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setCurrentDirectory(new File(System.getProperty("user.dir")));
                int result = fileChooser.showOpenDialog(null);

                if (result == JFileChooser.APPROVE_OPTION) {
                    File file = fileChooser.getSelectedFile();
                    try {
                        FileInputStream fis = new FileInputStream(file);
                        XSSFWorkbook workbook = new XSSFWorkbook(fis);
                        XSSFSheet selectedSheet = workbook.getSheet(textField.getText());
                        HashMap<String, double[]> commonMap = InputModule.InputModule(selectedSheet);
                        for (Map.Entry<String, double[]> entry : commonMap.entrySet()) {
                            String key = entry.getKey();
                            double[] values = entry.getValue();

                            // Вывод ключа и значений в одной строке
                            System.out.println(key + ": " + Arrays.toString(values));
                         }
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

                        exportButton.addActionListener(new ActionListener() {
                            @Override
                            public void actionPerformed(ActionEvent e) {
                                // Вызов функции экспорта из ExportModule
                                ExportModule.writeExcelFile(mapArray, Covarianc.getCovariance(commonMap));
                                frame.dispose();
                            }
                        });


                    } catch (Exception ex) {
                        ex.printStackTrace();
                        System.out.println("Ошибка при импорте и чтении данных из файла Excel.");
                        String errorMessage = "Ошибка при импорте и чтении данных из файла Excel";

                        // Отображение информационного сообщения
                        JOptionPane.showMessageDialog(null, errorMessage, "Ошибка", JOptionPane.INFORMATION_MESSAGE);
                    }
                }
            }
        })
        ;

        frame.add(textField);
        frame.add(importButton);
        frame.setLayout(null);
        frame.setVisible(true);
    }
}
