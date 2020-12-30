package Writers;

import Users.User;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.File;
import java.io.IOException;
import java.util.Set;

public class ExcelWriter {
    public ExcelWriter(Set<User> userCleanList) throws WriteException, IOException {
        WritableWorkbook workbook = Workbook.createWorkbook(new File("/home/hayk/temp/cleanList.xls"));
        WritableSheet sheet = workbook.createSheet("Sheet1", 0);
        Label cellId = new Label(0,0, "Id");
        sheet.addCell(cellId);
        Label cellName = new Label(1,0, "Name");
        sheet.addCell(cellName);
        Label cellEmail = new Label(2,0, "Email");
        sheet.addCell(cellEmail);
        int iterator = 0;
        for(User row:userCleanList){
            iterator++;
            Label labelId = new Label(0,iterator,row.getId());
            sheet.addCell(labelId);
            Label labelName = new Label(1,iterator,row.getName());
            sheet.addCell(labelName);
            Label labelEmail = new Label(2,iterator,row.getEmail());
            sheet.addCell(labelEmail);
        }
        workbook.write();
        workbook.close();
    }
}
