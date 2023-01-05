/*
 * Copyright (c) 2023. 01. 04. 17:57. Created by Andras Laczo. All rights reserved.
 */

package andras.laczo.services;
;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public final class OutputFileCreator {

    public static File createFile(File outputFolder) {
        // check if output exists
        String fileName = "/output";
        String fileExt = ".xlsx";
        String pathString = outputFolder.toString() + fileName + fileExt;
        Path path = Paths.get(pathString);

        //create file that not exists
        for (int i=1; Files.exists(path); i++) {
            pathString = outputFolder.toString() + fileName + i + fileExt;
            path = Paths.get(pathString);
        }

        File outFile = new File (path.toString());

        return outFile;

    }
}
