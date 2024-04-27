package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;

public class InputModule {
    public static HashMap<String, double[]> InputModule(Sheet selectedSheet) {
        int columnCount = selectedSheet.getRow(0).getLastCellNum();

        // Создаем HashMap для хранения данных по странам
        HashMap<String, double[]> countryData = new HashMap<>();

        // Проходимся по каждому столбцу, начиная со 2 элемента, и считываем данные по странам
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            // Получаем название страны из заголовка столбца
            String countryName = selectedSheet.getRow(0).getCell(columnIndex).getStringCellValue();

            // Создаем массив для хранения значений показателей
            double[] values = new double[selectedSheet.getLastRowNum()];

            // Проходимся по каждой строке, начиная со второй, и считываем значения показателей
            for (int rowIndex = 1; rowIndex <= selectedSheet.getLastRowNum(); rowIndex++) {
                Row row = selectedSheet.getRow(rowIndex);

                // Получаем значение показателя из ячейки в текущем столбце
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                    values[rowIndex - 1] = cell.getNumericCellValue();
                }
                else if (cell.getCellType() == CellType.FORMULA) {
                    values[rowIndex - 1] = cell.getNumericCellValue();
                }
            }

            // Добавляем данные страны в HashMap
            countryData.put(countryName, values);
        }

        return countryData;
    }
}
