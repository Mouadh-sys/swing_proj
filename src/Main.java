import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;


public class Main {

    public static void main(String[] args) {
        JFrame frame = new JFrame("Table");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);
        frame.setLayout(new BorderLayout());

//////////////////////////////////////////////////////////////////////////////////////////////////////
        JPanel panel = new JPanel();
        JTextField filePathField = new JTextField(30);
        JButton loadButton = new JButton("Load File");
        panel.add(filePathField);
        panel.add(loadButton);
        panel.setBackground(new Color(0,100,0));
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        DefaultTableModel model = new DefaultTableModel();
        JTable table = new JTable(model);
        frame.add(new JScrollPane(table), BorderLayout.CENTER);
        frame.add(panel, BorderLayout.NORTH);
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        loadButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String filePath = filePathField.getText();
                Path path = Paths.get(filePath);
                if (filePath == null || filePath.isEmpty()) {
                    JOptionPane.showMessageDialog(frame, "Invalid path", "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                String file = path.getFileName().toString();
                int dot_Index = file.lastIndexOf('.');
                if (dot_Index == -1 || dot_Index == file.length() - 1) {
                    JOptionPane.showMessageDialog(frame, "No file extension found", "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                String extension = file.substring(dot_Index + 1);
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if (extension.equals("csv")) {
                    try {
                        handleCSV(filePath, model);
                    } catch (FileNotFoundException ex) {
                        throw new RuntimeException(ex);
                    }
                } else if (extension.equals("xlsx") || extension.equals("xls")) {
                    try {
                        handleExcel(filePath, model);
                    } catch (IOException ex) {
                        throw new RuntimeException(ex);
                    }
                } else {
                    JOptionPane.showMessageDialog(frame, "Only csv/excel file", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        frame.setVisible(true);
    }
    ////////////////////////////////////////////////////////////////////////////////////
    public static void handleCSV (String filePath, DefaultTableModel model) throws FileNotFoundException {
        BufferedReader bReader = new BufferedReader(new FileReader(filePath));
        try  {
            String line;
            boolean Head = true;
            while ((line = bReader.readLine()) != null) {
                String[] values = line.split(",");
                if (Head) {
                    for (String value : values) {
                        model.addColumn(value);
                    }
                    Head = false;
                } else {
                    model.addRow(values);
                }
            }
        } catch (IOException e) {
            System.err.println("error");

        }
    }
    //////////////////////////////////////////////////////////////////////////////
    public static void handleExcel (String filePath, DefaultTableModel model) throws IOException {
        FileInputStream ExcelFile = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(ExcelFile);
        Sheet sheet = workbook.getSheetAt(0);



        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                model.addColumn(cell.getStringCellValue());
            }
        }

        // /////////////////
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                String[] rowData = new String[row.getLastCellNum()];
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    rowData[j] = (cell != null) ? cell.toString() : "";
                }
                model.addRow(rowData);
            }
            }
        }


    }


