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
package xyz.scarabya.multi.excel.calculator;

import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import xyz.scarabya.multi.excel.calculator.domain.FileNotSelectedException;
import xyz.scarabya.multi.excel.calculator.engine.Calculator;
import xyz.scarabya.multi.excel.calculator.log.LightLogger;

/**
 *
 * @author Alessandro Patriarca
 */
public class Main
{
    private final static Logger LOGGER
            = Logger.getLogger(Logger.GLOBAL_LOGGER_NAME);

    /**
     * @param args the command line arguments
     * @throws java.io.IOException
     */
    public static void main(String[] args) throws IOException
    {
        LightLogger.setup(args);

        LOGGER.log(Level.INFO, "Multi-excel-calculator log file."
                + " Current log level: {0}", LOGGER.getLevel().toString());

        File configFile, excelFolder, outputFile;
        if (args.length < 4)
            try
            {
                configFile = showFileChooser("Seleziona ll file di configurazione "
                        + "da utilizzare", "config");

                excelFolder = showFileChooser("Seleziona la cartella contenente "
                        + "i file Excel da processare", null);

                outputFile = showFileChooser("Seleziona ll file Excel in cui "
                        + "salvare i risultati", "xlsx");
            }
            catch (FileNotSelectedException e)
            {
                LOGGER.log(Level.INFO, "User requested cancel of current"
                        + " operation");
                throw e;
            }
        else
        {
            configFile = new File(args[1]);

            excelFolder = new File(args[2]);

            outputFile = new File(args[3]);
        }

        Calculator calculator = new Calculator(configFile, excelFolder,
                outputFile);

        try
        {
            LOGGER.log(Level.INFO, "Reading the mappings file");
            calculator.loadMappings();
        }
        catch (IOException ex)
        {
            LOGGER.log(Level.SEVERE, "Error reading the mappings file!", ex);
        }

        try
        {
            LOGGER.log(Level.INFO, "Reading Excel files");
            calculator.readAndSum();
        }
        catch (IOException | InvalidFormatException ex)
        {
            LOGGER.log(Level.SEVERE, "Error reading Excel files!", ex);
        }

        try
        {
            LOGGER.log(Level.INFO, "Writing results");
            calculator.writeResults();
        }
        catch (IOException | InvalidFormatException ex)
        {
            LOGGER.log(Level.SEVERE, "Error writing results!", ex);
        }

        LOGGER.log(Level.INFO, "All done, bye!");

        JOptionPane.showMessageDialog(null, "Elaborazione completata e salvata"
                + " nel file " + outputFile.getName(), "Completato",
                JOptionPane.INFORMATION_MESSAGE);
    }

    private static File showFileChooser(String message, String extension)
    {
        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView()
                .getHomeDirectory());
        jfc.setDialogTitle(message);
        if (extension != null)
        {
            jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
            jfc.setFileFilter(new FileNameExtensionFilter("File "
                    + extension.toUpperCase(), extension));
            if (extension.equals("config"))
                jfc.showSaveDialog(null);
            else
                jfc.showOpenDialog(null);
            if (jfc.getSelectedFile() == null)
                throw new FileNotSelectedException();
            String filename = jfc.getSelectedFile().getAbsolutePath();
            if (!filename.endsWith("." + extension))
                return new File(filename + "." + extension);
        }
        else
        {
            jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            jfc.showOpenDialog(null);
            if (jfc.getSelectedFile() == null)
                throw new FileNotSelectedException();
        }

        return jfc.getSelectedFile();
    }

}
