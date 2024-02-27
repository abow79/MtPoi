import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Arrays;
import java.util.*;
import javax.swing.*;

public class ExcelToJava {
    public static void main(String[] args) throws IOException {
        StringBuilder input=new StringBuilder();
        String tablename = "";
        XSSFWorkbook workbook;
        File target;
        try{
            String pathName =args[0];
            //String pathName ="C:\\Users\\st1\\Desktop\\新增資料夾 (4)";
            File dirFile=new File(pathName);
            if(!dirFile.exists()){
                System.out.println("Directory does not exist!");
                return;
            }
            if(!dirFile.isDirectory()){
                System.out.println("Please input a Directory path!");
                return;
            }
            //下面這行是獲取目錄下所有的檔案名和目錄名
            String[] filelist= dirFile.list();
            for(int z=0;z<filelist.length;z++) {
                if(filelist[z].contains(".")) {
                    String fileType = filelist[z].substring(filelist[z].lastIndexOf(".")+1);
                    if ("XLS".equalsIgnoreCase(fileType) || "XLSX".equalsIgnoreCase(fileType)) {
                        target = new File(dirFile.getPath(), filelist[z]);
                        //System.out.println(target.getName());
                        FileInputStream file = new FileInputStream(target);
                        workbook = new XSSFWorkbook(file);
                        XSSFSheet sheet = workbook.getSheet(target.getName().substring(0, target.getName().lastIndexOf(".")));
                        tablename = sheet.getSheetName();
                        input.append("public class " + tablename + " extends MTMessage{")
                                .append("\n\n");
                        input.append("public " + tablename + "() {")
                                .append("\n")
                                .append("this.messageType = this.getClass();")
                                .append("\n")
                                .append("}")
                                .append("\n\n");
                        for (Row a : sheet) {
                            if (a.getRowNum() == 0) {
                                continue;
                            }
                            if(cellJudge(file,a.getCell(4)).equals("Mandatory") || cellJudge(file,a.getCell(4)).equals("Optional")) {
                                input.append("@Presence(PresenceType." + cellJudge(file, a.getCell(4)) + ")");
                                input.append("\n");
                            }else if(cellJudge(file,a.getCell(4)).equals("M")) {
                                input.append("@Presence(PresenceType.Mandatory)");
                                input.append("\n");
                            }else if(cellJudge(file,a.getCell(4)).equals("O")) {
                                input.append("@Presence(PresenceType.Optional)");
                                input.append("\n");
                            }

                            input.append("@ColumnId("+"\""+ cellJudge(file,a.getCell(0))+"\""+ ")");
                            input.append("\n");

                            for (int i = 5; i <= 16; i++) {
                                if (cellJudge(file,a.getCell(i)) != "") {
                                    String tagName=sheet.getRow(0).getCell(i).toString();
                                    if(i==9 || i==12 || i==16 ){
                                        input.append("@" + tagName.substring(0, tagName.indexOf("\n"))+ "(" +"\""+cellJudge(file,a.getCell(i))+"\"" + ")");
                                        input.append("\n");
                                    } else if(i==10 || i==11){
                                        input.append("@" + tagName.substring(0, tagName.indexOf("\n"))+ "(" +"literal = \""+cellJudge(file,a.getCell(i))+"\"" + ")");
                                        input.append("\n");
                                    } else if (i==8) {
                                        input.append("@" + tagName.substring(0, tagName.indexOf("\n"))+ "(" +"pos = "+cellJudge(file,a.getCell(i))+")");
                                        input.append("\n");
                                    } else {
                                        input.append("@" + tagName.substring(0, tagName.indexOf("\n")) + "(" + cellJudge(file, a.getCell(i)) + ")");
                                        input.append("\n");
                                    }
                                }
                            }

                            if(cellJudge(file,a.getCell(3)).contains("List")){
                                input.append("@ListItemType(" + cellJudge(file, a.getCell(3)).substring(a.getCell(3).toString().indexOf("<")+1,a.getCell(3).toString().indexOf(">")) + ".class)");
                                input.append("\n");
                            }

                            if (cellJudge(file,a.getCell(3)).equals("List<String>")) {
                                input.append("protected " + cellJudge(file,a.getCell(3)) + " col" + cellJudge(file,a.getCell(0)) + toCamelcase(cellJudge(file,a.getCell(1))) + " = new ArrayList<>();");
                                input.append("\n");
                            } else {
                                input.append("protected " + cellJudge(file,a.getCell(3)) + " col" + cellJudge(file,a.getCell(0)) + toCamelcase(cellJudge(file,a.getCell(1))) + ";");
                                input.append("\n");
                            }

                            input.append("public " + cellJudge(file,a.getCell(3)) + " get" + cellJudge(file,a.getCell(0)) + toCamelcase(cellJudge(file,a.getCell(1))) + "() {");
                            input.append("\n");

                            if (cellJudge(file,a.getCell(3)).equals("List<String>")) {
                                input.append("\treturn Collections.unmodifiableList(col" + cellJudge(file,a.getCell(0)) + toCamelcase(cellJudge(file,a.getCell(1)))+");");
                                input.append("\n");
                                input.append("}");
                                input.append("\n\n");
                            } else {
                                input.append("\treturn col" + cellJudge(file,a.getCell(0)) + toCamelcase(cellJudge(file,a.getCell(1))) + ";");
                                input.append("\n");
                                input.append("}");
                                input.append("\n\n");
                            }

                        }
                    }
                }else {
                    continue;
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        input.append("}");
        File f1 = new File("C:\\Users\\st1\\Desktop\\測試區\\"+tablename+".java");
        FileWriter result = new FileWriter(f1,false);
        result.write(input.toString());
        result.close();
    }

    private static String cellJudge(FileInputStream in,Cell x) throws IOException {
        String result="";
        if(x==null){
            return "";
        }else if(x.getCellType().equals(CellType.NUMERIC)){
            result= NumberToTextConverter.toText(x.getNumericCellValue());
        }else if(x.getCellType().equals(CellType.FORMULA)){
            DataFormatter df=new DataFormatter();
            XSSFWorkbook workbook = new XSSFWorkbook(in);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            result = df.formatCellValue(x, evaluator);
        }else {
            result=x.toString();
        }
        return result;
    }


    private static String cellToColumn(Cell c) {
        String name=c.getAddress().toString();
        name=name.substring(0,1);
        return name;
    }



    private static String toCamelcase(String cellToString){
        cellToString=cellToString.trim();
        StringBuilder builder=new StringBuilder();
            String[] result=cellToString.split(" ");
            for(int i=0;i< result.length;i++){
                String word=result[i];
                if(i==0){
                    word=word.isEmpty()?word: word.toLowerCase();
                }else {
                    word=word.isEmpty()?word: Character.toUpperCase(word.charAt(0))+word.substring(1).toLowerCase();
                }
                builder.append(word);
            }
        return builder.toString();
    }
}
