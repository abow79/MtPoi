import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;


public class PoiForm extends JFrame{
    private JButton btnOk;
    private JButton btnClear;
    private JTextField textField1;
    private JPanel Panel1;
    private JTextField textField2;
    private JButton inputDirectoryButton;
    private JButton outputDirectoryButton;

    private int numClicks=0;

    String configFile = "config.properties";

    static Properties properties = new Properties();


    StringBuilder input=new StringBuilder();
    String tablename = "";
    XSSFWorkbook workbook;
    File target;

    int versionNumber=0;

    public PoiForm() throws Exception{
        btnOk.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    String path = textField1.getText();
//                    System.setProperty("input_value",path);
//                    System.setProperty("output_value",textField2.getText());
                    properties.setProperty("input_value",path);
                    properties.setProperty("output_value",textField2.getText());
                    properties.store(new FileOutputStream("./properties.config"),null);
                    File dirFile = new File(path);
                    if (!dirFile.exists()) {
                        JOptionPane.showMessageDialog(null,"Directory does not exist!");
                        return;
                    }
                    if (!dirFile.isDirectory()) {
                        JOptionPane.showMessageDialog(null,"Please input a Directory path!");
                        return;
                    }
                    String[] filelist= dirFile.list();
                    for(int z=0;z<filelist.length;z++) {
                        if(filelist[z].contains(".") && filelist[z].substring(0,2).equalsIgnoreCase("MT")) {
                            String fileType = filelist[z].substring(filelist[z].lastIndexOf(".")+1);
                            if ("XLS".equalsIgnoreCase(fileType) || "XLSX".equalsIgnoreCase(fileType)) {
                                target = new File(dirFile.getPath(), filelist[z]);
                                //System.out.println(target.getName());
                                FileInputStream file = new FileInputStream(target);
                                workbook = new XSSFWorkbook(file);
                                XSSFSheet sheet = workbook.getSheet(target.getName().substring(0, target.getName().lastIndexOf(".")));
                                tablename = sheet.getSheetName();
                                input.append("package swiftmessages.messagetype;")
                                     .append("\n\n")
                                     .append("import swiftmessages.annotations.*;")
                                     .append("\n")
                                     .append("import swiftmessages.enums.PresenceType;")
                                     .append("\n\n")
                                     .append("import java.math.BigDecimal;")
                                     .append("\n")
                                     .append("import java.util.ArrayList;")
                                     .append("\n")
                                     .append("import java.util.Collections;")
                                     .append("\n")
                                     .append("import java.util.Date;")
                                     .append("\n")
                                     .append("import java.util.List;")
                                     .append("\n")
                                     .append("import java.time.LocalDate;")
                                     .append("\n")
                                     .append("import java.time.OffsetTime;")
                                     .append("\n");
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
                                    if(cellJudge(file,a.getCell(5)).equals("Mandatory") || cellJudge(file,a.getCell(5)).equals("Optional")) {
                                        input.append("@Presence(PresenceType." + cellJudge(file, a.getCell(5)) + ")");
                                        input.append("\n");
                                    }else if(cellJudge(file,a.getCell(5)).equals("M")) {
                                        input.append("@Presence(PresenceType.Mandatory)");
                                        input.append("\n");
                                    }else if(cellJudge(file,a.getCell(5)).equals("O")) {
                                        input.append("@Presence(PresenceType.Optional)");
                                        input.append("\n");
                                    }

                                    input.append("@ColumnId("+"\""+ cellJudge(file,a.getCell(0))+"\""+ ")");
                                    input.append("\n");

                                    for (int i = 6; i <= 18; i++) {
                                        if (cellJudge(file,a.getCell(i)) != "") {
                                            String tagName=sheet.getRow(0).getCell(i).toString();
                                            if(i == 10 || i == 13 || i == 17 || i == 16 || i == 18){
                                                input.append("@" + tagName.substring(0, tagName.indexOf("\n"))+ "(" +"\""+cellJudge(file,a.getCell(i))+"\"" + ")");
                                                input.append("\n");
                                            } else if(i == 11 || i == 12){
                                                input.append("@" + tagName.substring(0, tagName.indexOf("\n"))+ "(" +"literal = \""+cellJudge(file,a.getCell(i))+"\"" + ")");
                                                input.append("\n");
                                            } else {
                                                input.append("@" + tagName.substring(0, tagName.indexOf("\n")) + "(" + cellJudge(file, a.getCell(i)) + ")");
                                                input.append("\n");
                                            }
                                        }
                                    }

                                    if (cellJudge(file, a.getCell(4)).matches("List<.+>")) {
                                        input.append("@ListItemType(" + cellJudge(file, a.getCell(4)).substring(a.getCell(4).toString().indexOf("<") + 1, a.getCell(4).toString().indexOf(">")) + ".class)");
                                        input.append("\n");
                                    }

                                    if(cellJudge(file,a.getCell(1)).length()!=0 && cellJudge(file,a.getCell(2)).length()==0) {
                                        if (cellJudge(file, a.getCell(4)).matches("List<.+>")) {
                                            input.append("protected " + cellJudge(file, a.getCell(4)) + " col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) + " = new ArrayList<>();");
                                            input.append("\n");
                                        } else {
                                            input.append("protected " + cellJudge(file, a.getCell(4)) + " col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) + ";");
                                            input.append("\n");
                                        }
                                    } else if (cellJudge(file,a.getCell(1)).length()==0 && cellJudge(file,a.getCell(2)).length()!=0) {
                                        if (cellJudge(file, a.getCell(4)).matches("List<.+>")) {
                                            input.append("protected " + cellJudge(file, a.getCell(4)) + " col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(2))) + " = new ArrayList<>();");
                                            input.append("\n");
                                        } else {
                                            input.append("protected " + cellJudge(file, a.getCell(4)) + " col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(2))) + ";");
                                            input.append("\n");
                                        }
                                    } else if (cellJudge(file,a.getCell(1)).length()!=0 && cellJudge(file,a.getCell(2)).length()!=0) {
                                        if (cellJudge(file, a.getCell(4)).matches("List<.+>")) {
                                            input.append("protected " + cellJudge(file, a.getCell(4)) + " col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) +"_"+toCamelcase(cellJudge(file, a.getCell(2)))+" = new ArrayList<>();");
                                            input.append("\n");
                                        } else {
                                            input.append("protected " + cellJudge(file, a.getCell(4)) + " col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) +"_"+toCamelcase(cellJudge(file, a.getCell(2)))+";");
                                            input.append("\n");
                                        }
                                    }

                                    if(cellJudge(file,a.getCell(1)).length()!=0 && cellJudge(file,a.getCell(2)).length()==0) {
                                        input.append("public " + cellJudge(file, a.getCell(4)) + " get" + cellJudge(file, a.getCell(0)) + "() {");
                                        input.append("\n");

                                        if (cellJudge(file, a.getCell(4)).equals("List<String>")) {
                                            input.append("\treturn Collections.unmodifiableList(col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) + ");");
                                            input.append("\n");
                                            input.append("}");
                                            input.append("\n\n");
                                        } else {
                                            input.append("\treturn col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) + ";");
                                            input.append("\n");
                                            input.append("}");
                                            input.append("\n\n");
                                        }

                                        input.append("public " + cellJudge(file, a.getCell(4)) + " get" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) + "() {");
                                        input.append("\n");

                                        if (cellJudge(file, a.getCell(4)).equals("List<String>")) {
                                            input.append("\treturn Collections.unmodifiableList(col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) + ");");
                                            input.append("\n");
                                            input.append("}");
                                            input.append("\n\n");
                                        } else {
                                            input.append("\treturn col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) + ";");
                                            input.append("\n");
                                            input.append("}");
                                            input.append("\n\n");
                                        }
                                    } else if (cellJudge(file,a.getCell(1)).length()==0 && cellJudge(file,a.getCell(2)).length()!=0) {
                                        input.append("public " + cellJudge(file, a.getCell(4)) + " get" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(2))) + "() {");
                                        input.append("\n");

                                        if (cellJudge(file, a.getCell(4)).equals("List<String>")) {
                                            input.append("\treturn Collections.unmodifiableList(col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(2))) + ");");
                                            input.append("\n");
                                            input.append("}");
                                            input.append("\n\n");
                                        } else {
                                            input.append("\treturn col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(2))) + ";");
                                            input.append("\n");
                                            input.append("}");
                                            input.append("\n\n");
                                        }
                                    } else if (cellJudge(file,a.getCell(1)).length()!=0 && cellJudge(file,a.getCell(2)).length()!=0) {

                                        input.append("public " + cellJudge(file, a.getCell(4)) + " get" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(2)))+"() {");
                                        input.append("\n");

                                        if (cellJudge(file, a.getCell(4)).equals("List<String>")) {
                                            input.append("\treturn Collections.unmodifiableList(col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) +"_"+toCamelcase(cellJudge(file, a.getCell(2)))+ ");");
                                            input.append("\n");
                                            input.append("}");
                                            input.append("\n\n");
                                        } else {
                                            input.append("\treturn col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) +"_"+toCamelcase(cellJudge(file, a.getCell(2)))+";");
                                            input.append("\n");
                                            input.append("}");
                                            input.append("\n\n");
                                        }


                                        input.append("public " + cellJudge(file, a.getCell(4)) + " get" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1)))+"_"+toCamelcase(cellJudge(file, a.getCell(2))) +"() {");
                                        input.append("\n");

                                        if (cellJudge(file, a.getCell(4)).equals("List<String>")) {
                                            input.append("\treturn Collections.unmodifiableList(col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) +"_"+toCamelcase(cellJudge(file, a.getCell(2)))+ ");");
                                            input.append("\n");
                                            input.append("}");
                                            input.append("\n\n");
                                        } else {
                                            input.append("\treturn col" + cellJudge(file, a.getCell(0)) + "_" + toCamelcase(cellJudge(file, a.getCell(1))) +"_"+toCamelcase(cellJudge(file, a.getCell(2)))+";");
                                            input.append("\n");
                                            input.append("}");
                                            input.append("\n\n");
                                        }

                                    }

                                }
                            }
                        }else {
                            continue;
                        }

                        input.append("}");
                        File f1 = new File(textField2.getText()+"\\"+tablename+".java");
                        if(f1.exists()){
                            versionNumber++;
                            try {
                                f1= new File(textField2.getText()+"\\"+tablename+"("+versionNumber+")"+".java");
                                f1.createNewFile();
                            } catch (IOException ex) {
                                throw new RuntimeException(ex);
                            }
                        }
                        FileWriter result = null;
                        try {
                            result = new FileWriter(f1,false);
                            result.write(input.toString());
                            input.delete(0,input.length());
                            result.close();
                        } catch (IOException ex) {
                            throw new RuntimeException(ex);
                        }
                    }
                    }catch (Exception ex2){
                    ex2.printStackTrace();
                }
                }

        });


        btnClear.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                numClicks++;
                if(numClicks%2==1){
                    textField1.setText("");
                } else if (numClicks%2==0) {
                    textField2.setText("");
                }
            }
        });
        inputDirectoryButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jFileChooser=new JFileChooser();
                jFileChooser.setDialogTitle("請選擇輸入目錄");
                jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int option=jFileChooser.showOpenDialog(null);
                if(option==JFileChooser.APPROVE_OPTION){
                    File file= jFileChooser.getSelectedFile();
                    textField1.setText(file.getPath());
                }
            }
        });
        outputDirectoryButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jFileChooser2=new JFileChooser();
                jFileChooser2.setDialogTitle("請選擇輸出目錄");
                jFileChooser2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int option=jFileChooser2.showOpenDialog(null);
                if(option==JFileChooser.APPROVE_OPTION){
                    File file2=jFileChooser2.getSelectedFile();
                    textField2.setText(file2.getPath());
                }
            }
        });
    }

    public static void main(String[] args) throws Exception {
        JFrame frame = new JFrame("PoiForm");
        frame.setSize(600,300);
        PoiForm poiForm = new PoiForm();
        File f=new File("./properties.config");
        if(!f.exists()){
            f.createNewFile();
        }
        properties.load(new FileReader("./properties.config"));
        //if(!properties.getProperty("input_value").equals("") && !properties.getProperty("output_value").equals("")) {
        String input_value = properties.getProperty("input_value");
        String output_value = properties.getProperty("output_value");
        //poiForm.textField1.setText(input_value);
        //poiForm.textField2.setText(output_value);
        //}
        frame.setContentPane(poiForm.Panel1);
        List<Component> result=new ArrayList<>();
        for(Component sc:((JPanel)frame.getContentPane()).getComponents()) {
            if(sc instanceof JTextField){
                result.add(sc);
            }
        }
        JTextField textField2=(JTextField) result.get(0);
        textField2.setText(output_value);

        JTextField textField1=(JTextField) result.get(1);
        textField1.setText(input_value);



        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }



    private static String cellJudge(FileInputStream in, Cell x) throws IOException {
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
