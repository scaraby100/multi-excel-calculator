/*
 * Copyright 2018 Alessandro Patriarca.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package xyz.scarabya.multi.excel.calculator.engine;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

/**
 *
 * @author Alessandro Patriarca
 */
public class ExcelRW
{
    private Workbook actualWb;

    private final static String COL_REGEX = "([a-z]+)"; //Column
    private final static String ROW_REGEX = "(\\d+)"; //Row

    private final static Pattern PATTERN = Pattern.compile(
            COL_REGEX + ROW_REGEX, Pattern.CASE_INSENSITIVE | Pattern.DOTALL);

    private Sheet actualSheet;
    private Row actualRow;

    public void loadExcelFile(File excelFile) throws IOException,
            InvalidFormatException
    {
        actualWb = WorkbookFactory.create(excelFile);
        actualSheet = null;
        actualRow = null;
    }

    public double getCellValueAt(String coordinates)
    {
        String[] sheetCellSplit = coordinates.split("\\[|\\]");
        if (actualSheet == null
                || !actualSheet.getSheetName().equals(sheetCellSplit[0]))
        {
            actualSheet = actualWb.getSheet(sheetCellSplit[0]);
            actualRow = null;
        }
        Matcher matcher = PATTERN.matcher(sheetCellSplit[1]);
        if (matcher.find())
        {
            int columnNum = CellReference.convertColStringToIndex(matcher.group(1));
            int rowNum = Integer.parseInt(matcher.group(2)) - 1;
            if (actualRow == null || actualRow.getRowNum() != rowNum)
                actualRow = actualSheet.getRow(rowNum);
            return actualRow.getCell(columnNum).getNumericCellValue();
        }
        return 0;
    }

    public void setCellValue(String coordinates, double value)
    {
        String[] sheetCellSplit = coordinates.split("\\[|\\]");
        if (actualSheet == null
                || !actualSheet.getSheetName().equals(sheetCellSplit[0]))
        {
            actualSheet = actualWb.getSheet(sheetCellSplit[0]);
            actualRow = null;
        }
        Matcher matcher = PATTERN.matcher(sheetCellSplit[1]);
        if (matcher.find())
        {
            int columnNum = CellReference.convertColStringToIndex(matcher.group(1));
            int rowNum = Integer.parseInt(matcher.group(2)) - 1;
            if (actualRow == null || actualRow.getRowNum() != rowNum)
                actualRow = createOrGetRow(actualSheet, rowNum);
            createOrSetCellValue(actualRow, columnNum, value);
        }
    }

    private Row createOrGetRow(Sheet sheet, int rowNum)
    {
        Row row = sheet.getRow(rowNum);
        if (row == null)
            row = sheet.createRow(rowNum);
        return row;
    }

    private void createOrSetCellValue(Row row, int columnNum, double value)
    {
        Cell cell = row.getCell(columnNum);
        if (cell == null)
            cell = row.createCell(columnNum);
        cell.setCellValue(value);
    }

    public void saveExcelFile(File excelFile) throws FileNotFoundException, IOException
    {
        File tempExcelFile = new File(excelFile.getAbsolutePath() + "_TMP");
        try (FileOutputStream fileOut = new FileOutputStream(
                tempExcelFile.getAbsolutePath()))
        {
            actualWb.write(fileOut);
        }
        actualWb.close();
        excelFile.delete();
        tempExcelFile.renameTo(excelFile);
    }
}
