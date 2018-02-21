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

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import xyz.scarabya.multi.excel.calculator.domain.Result;

/**
 *
 * @author Alessandro Patriarca
 */
public class Calculator
{
    private final static Logger LOGGER
            = Logger.getLogger(Logger.GLOBAL_LOGGER_NAME);
    
    private final Map<String, List<String>> mappings;
    private final Map<String, Result> resultsMap;
    private final File configFile, inputFolder, outputFile;

    public Calculator(File configFile, File inputFolder, File outputFile)
    {
        this.configFile = configFile;
        this.inputFolder = inputFolder;
        this.outputFile = outputFile;

        mappings = new HashMap<>();
        resultsMap = new HashMap<>();
    }

    public void loadMappings() throws IOException
    {
        String line;
        List<String> outputs;
        if (configFile.exists())
            try (BufferedReader br = new BufferedReader(
                    new FileReader(configFile)))
            {
                while ((line = br.readLine()) != null)
                {
                    String[] inOut = line.split(";");
                    if (!resultsMap.containsKey(inOut[1]))
                        resultsMap.put(inOut[1], new Result(0));
                    if (mappings.containsKey(inOut[0]))
                    {
                        outputs = mappings.get(inOut[0]);
                        if (!outputs.contains(inOut[1]))
                            outputs.add(inOut[1]);
                    }
                    else
                    {
                        outputs = new ArrayList<>();
                        outputs.add(inOut[1]);
                        mappings.put(inOut[0], outputs);
                    }
                }
            }

    }

    public void readAndSum() throws IOException, InvalidFormatException
    {
        ExcelRW reader = new ExcelRW();
        for (File excelFile : inputFolder.listFiles())
        {
            LOGGER.log(Level.INFO, "Reading {0}", excelFile.getName());
            reader.loadExcelFile(excelFile);
            for (String mapping : mappings.keySet())
            {
                double value = reader.getCellValueAt(mapping);
                for (String outputs : mappings.get(mapping))
                    resultsMap.get(outputs).sum(value);

            }
        }
    }

    public void writeResults() throws IOException, InvalidFormatException
    {
        ExcelRW writer = new ExcelRW();
        LOGGER.log(Level.INFO, "Loading output Excel: {0}",
                outputFile.getName());
        writer.loadExcelFile(outputFile);
        LOGGER.log(Level.INFO, "Writing new results");
        for (String resultDest : resultsMap.keySet())
            writer.setCellValue(resultDest, resultsMap.get(resultDest)
                    .getResultValue());
        LOGGER.log(Level.INFO, "Saving file");
        writer.saveExcelFile(outputFile);
    }
}
