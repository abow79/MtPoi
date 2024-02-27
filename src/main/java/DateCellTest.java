import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class DateCellTest{
    public static void main(String[] args)throws Exception{
        XSSFWorkbook wb=new XSSFWorkbook();
        CreationHelper creationHelper=wb.getCreationHelper();
        try{
            //OutputStream os=new FileOutputStream("C:\\Users\\st1\\Desktop\\Datecell.xlsx");
            XSSFSheet sheet=wb.createSheet();
            Row row= sheet.createRow(0);
            Cell cell=row.createCell(0);
            CellStyle cellStyle=wb.createCellStyle();
            cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("YYMMDD"));
            cell= row.createCell(1);
            cell.setCellValue(new Date());
            cell.setCellStyle(cellStyle);
            //formatter.formatCellValue()可以直接获取单元格内容
            DataFormatter df=new DataFormatter();
            String result=df.formatCellValue(cell).toString();
            System.out.println(result);
            List<String> a=new ArrayList<String>(List.of("abow","kevin","dora"));
            System.out.println(a.getClass());
            //wb.write(os);
        }catch (Exception e){
            System.out.println(e.getMessage());
        }

    }
}
