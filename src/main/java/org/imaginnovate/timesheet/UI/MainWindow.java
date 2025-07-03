package org.imaginnovate.timesheet.UI;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;
import org.imaginnovate.timesheet.services.FileProcessService;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.Collections;
import java.util.Vector;

public class MainWindow extends JFrame {

    private final DefaultTableModel tableModel;
    private final JButton formatButton;
    private final FileProcessService fileProcessService = new FileProcessService();

    public MainWindow() {
        super("Time Sheet Formatter");
        tableModel = new DefaultTableModel();
        JTable table = new JTable(tableModel);
        JScrollPane scrollPane = new JScrollPane(table);

        JButton openButton = new JButton("Open CSV File");

        formatButton = new JButton("Format");
        formatButton.setEnabled(false);
        JPanel buttonPanel = new JPanel();

        buttonPanel.add(openButton);
        buttonPanel.add(formatButton);
        formatButton.addActionListener(this::exportToExcel);

        openButton.addActionListener(this::openCsvFile);

        JPanel panel = new JPanel(new BorderLayout());
        panel.add(buttonPanel, BorderLayout.NORTH);
        panel.add(scrollPane, BorderLayout.CENTER);
        add(panel);

        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(800, 600);
        setLocationRelativeTo(null);
        setVisible(true);
    }

    private void exportToExcel(ActionEvent actionEvent) {
        System.out.println("Input received, generating report..");
        // Create a modal dialog
        JDialog dialog = new JDialog(this, "Processing", Dialog.ModalityType.APPLICATION_MODAL);
        dialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
        dialog.add(new JLabel("Please wait, processing..."), BorderLayout.CENTER);
        dialog.setSize(200, 100);
        dialog.setLocationRelativeTo(this);

        new Thread(() -> {
            dialog.setVisible(true); // Show the dialog (blocks the EDT)
        }).start();

        new Thread(() -> {
            boolean status  = fileProcessService.process(tableModel);
            dialog.dispose();
            if(status){
                JOptionPane.showMessageDialog(this, "Time Sheet generated successfully. - " + FileProcessService.OUTPUT_FILE_PATH, "Success", JOptionPane.INFORMATION_MESSAGE);
            }else{
                JOptionPane.showMessageDialog(this, "Time Sheet generation failed.", "Failed", JOptionPane.ERROR_MESSAGE);
            }
        }).start();
    }

    private void openCsvFile(ActionEvent e) {
        JFileChooser fileChooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("CSV Files (*.csv)", "csv");
        fileChooser.setFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false);
        int option = fileChooser.showOpenDialog(this);
        if (option == JFileChooser.APPROVE_OPTION) {
            File csvFile = fileChooser.getSelectedFile();
            loadCsvToTable(csvFile);
        }
    }

    private void loadCsvToTable(File file) {
        try {
            tableModel.setRowCount(0);
            tableModel.setColumnCount(0);
            boolean isFirstLine = true;
            CSVReader reader = new CSVReader(new FileReader(file.getPath()));
            String[] lineItems;
            while ((lineItems = reader.readNext()) != null) {
                if (isFirstLine) {
                    for (String col : lineItems) {
                        tableModel.addColumn(col);
                    }
                    isFirstLine = false;
                } else {
                    Vector<String> row = new Vector<>();
                    Collections.addAll(row, lineItems);
                    tableModel.addRow(row);
                }
            }
            formatButton.setEnabled(tableModel.getRowCount() > 0);
        } catch (CsvValidationException | IOException e) {
            throw new RuntimeException(e);
        }
    }
}

