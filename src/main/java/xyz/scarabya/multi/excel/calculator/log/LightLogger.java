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
package xyz.scarabya.multi.excel.calculator.log;

import java.io.IOException;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

/**
 *
 * @author Alessandro Patriarca
 */
public class LightLogger
{
    static private FileHandler fileTxt;
    static private SimpleFormatter formatterTxt;

    private static final String DEF_LEVEL = "INFO";

    static public void setup(String[] args) throws IOException
    {
        String logLevel = DEF_LEVEL;
        if (args.length > 0)
            logLevel = args[0];
        System.setProperty("java.util.logging.SimpleFormatter.format",
                "%4$s: %5$s%n");
        Logger logger = Logger.getLogger(Logger.GLOBAL_LOGGER_NAME);

        logger.setLevel(Level.parse(logLevel));

        fileTxt = new FileHandler("calc_log.txt");

        formatterTxt = new SimpleFormatter();
        fileTxt.setFormatter(formatterTxt);
        logger.addHandler(fileTxt);
    }
}
