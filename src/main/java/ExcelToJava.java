import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

public class ExcelToJava {
    public static void main(String[] args) throws IOException {
        StringBuilder input=new StringBuilder();
        String tablename=null;
        try{
            FileInputStream file = new FileInputStream(new File("C:\\Users\\st1\\Desktop\\MT910toPOI.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
//            DataFormatter ;
//            FormulaEvaluator;
            XSSFSheet sheet = workbook.getSheet("MT910");
            tablename=sheet.getSheetName();
            input.append("public class "+tablename+" extends MTMessage{")
                    .append("\n");
            input.append("public "+tablename+"() {")
                    .append("\n")
                    .append("this.messageType = this.getClass();")
                    .append("\n")
                    .append("}")
                    .append("\n");
            for (Row a : sheet) {
                //System.out.println(a);
                if (a.getRowNum() == 0) {
                    continue;
                }
                        input.append("@Presence(" + cellJudge(a.getCell(4)) + ")");
                        input.append("\n");

                        input.append("@ColumnId(" + cellJudge(a.getCell(0)) + ")");
                        input.append("\n");

                    for(int i=5;i<=16;i++){
                        if(cellJudge(a.getCell(i))!=""){
                            input.append("@"+sheet.getRow(0).getCell(i).toString()+"("+a.getCell(i).toString()+")");
                            input.append("\n");
                        }
                    }

                        if (cellJudge(a.getCell(3)).equals("List<String>")) {
                            input.append("protected " + cellJudge(a.getCell(3)) + " col" + cellJudge(a.getCell(0)) + "_" + toCamelcase(cellJudge(a.getCell(1))) + " = new ArrayList<>();");
                            input.append("\n");
                        } else {
                            input.append("protected " + cellJudge(a.getCell(3)) + " col" + cellJudge(a.getCell(0)) + "_" + toCamelcase(cellJudge(a.getCell(1))) + ";");
                            input.append("\n");
                        }

                        input.append("public " + cellJudge(a.getCell(3)) + " get" + cellJudge(a.getCell(0)) + "_" + toCamelcase(cellJudge(a.getCell(1))) + "() {");
                        input.append("\n");

                        if (cellJudge(a.getCell(3)).equals("List<String>")) {
                            input.append("return Collections.unmodifiableList(col" + cellJudge(a.getCell(0)) + "_" + toCamelcase(cellJudge(a.getCell(1))));
                            input.append("\n");
                            input.append("}");
                            input.append("\n");
                        } else {
                            input.append("return col" + cellJudge(a.getCell(0)) + "_" + toCamelcase(cellJudge(a.getCell(1))) + ";");
                            input.append("\n");
                            input.append("}");
                            input.append("\n");
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

    private static String cellJudge(Cell x){
        if(x==null){
            return "";
        }else{
            return x.toString();
        }
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
