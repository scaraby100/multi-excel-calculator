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
        LightLogger.setup();
        
        File configFile = showFileChooser("Seleziona ll file di configurazione "
                + "da utilizzare", "config");

        File excelFolder = showFileChooser("Seleziona la cartella contenente "
                + "i file Excel da processare", null);

        File outputFile = showFileChooser("Seleziona ll file Excel in cui "
                + "salvare i risultati", "xlsx");

        Calculator calculator = new Calculator(configFile, excelFolder,
                outputFile);
        
        try
        {
            calculator.loadMappings();
        }
        catch (IOException ex)
        {
            LOGGER.log(Level.SEVERE, "Error reading the configuration file!", ex);
        }

        try
        {
            calculator.readAndSum();
        }
        catch (IOException | InvalidFormatException ex)
        {
            LOGGER.log(Level.SEVERE, "Error reading Excel files!", ex);
        }

        try
        {
            calculator.writeResults();
        }
        catch (IOException | InvalidFormatException ex)
        {
            LOGGER.log(Level.SEVERE, "Error writing results!", ex);
        }
        
        JOptionPane.showMessageDialog(null, "Elaborazione completata e salvata"
                + " nel file "+outputFile.getName(), "Completato",
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
            jfc.showSaveDialog(null);
            String filename = jfc.getSelectedFile().getAbsolutePath();
            if (!filename.endsWith("." + extension))
                return new File(filename + "." + extension);
        }
        else
        {
            jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            jfc.showOpenDialog(null);
        }
        return jfc.getSelectedFile();
    }

}