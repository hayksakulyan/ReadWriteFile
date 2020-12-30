package Readers;

import Interfaces.FileReader;
import Users.User;
import Writers.ExcelWriter;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;

import java.io.File;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

public class ExcelReader implements FileReader {
    public static void main(String[] args) throws IOException, WriteException {
        ExcelReader er = new ExcelReader();
        ExcelWriter ew = new ExcelWriter(er.fileReader("/home/hayk/temp/hayk.xls"));
    }
    @Override
    public Set<User> fileReader(String path) {
        Workbook workbook = null;
        Set<User> userCleanList = new HashSet<>();
        try {
            workbook = Workbook.getWorkbook(new File(path));
            Sheet sheet = workbook.getSheet(0);
            int totalRows = sheet.getRows();
            for (int row = 1; row < totalRows; row++) {
                String id = sheet.getCell(0, row).getContents();
                String name = sheet.getCell(1, row).getContents();
                String email = sheet.getCell(2, row).getContents();
                    userCleanList.add(new User(id, name, email));
            }
        } catch (IOException | BiffException e) {
            e.printStackTrace();
        }
        return userCleanList;
    }
}
