import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.dnd.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class OrderCounter extends JFrame {
    private JButton chooseFileButton;
    private JButton processButton;
    private JTable resultsTable;
    private File selectedFile;

    public OrderCounter() {
        setTitle("Product Quantity Counter");
        setLayout(new BorderLayout());
        setSize(600, 400);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        chooseFileButton = new JButton("Choose Excel File");
        processButton = new JButton("Process File");
        processButton.setEnabled(false);
        resultsTable = new JTable(new DefaultTableModel(new Object[]{"Product Name", "Variation Name", "Quantity"}, 0));

        DefaultTableCellRenderer centerRenderer = new DefaultTableCellRenderer();
        centerRenderer.setHorizontalAlignment(SwingConstants.CENTER);
        resultsTable.getColumnModel().getColumn(1).setCellRenderer(centerRenderer);
        resultsTable.getColumnModel().getColumn(2).setCellRenderer(centerRenderer);

        chooseFileButton.addActionListener(e -> chooseFile());
        processButton.addActionListener(e -> processFile());

        JPanel panel = new JPanel();
        panel.add(chooseFileButton);
        panel.add(processButton);

        enableDragAndDrop();
        add(panel, BorderLayout.NORTH);
        add(new JScrollPane(resultsTable), BorderLayout.CENTER);
    }

    private void chooseFile() {
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            selectedFile = fileChooser.getSelectedFile();
            processButton.setEnabled(true);
        }
    }

    private void enableDragAndDrop() {
        new DropTarget(this, new DropTargetListener() {
            public void dragEnter(DropTargetDragEvent dtde) {}
            public void dragOver(DropTargetDragEvent dtde) {}
            public void dropActionChanged(DropTargetDragEvent dtde) {}
            public void dragExit(DropTargetEvent dte) {}
            public void drop(DropTargetDropEvent dtde) {
                dtde.acceptDrop(DnDConstants.ACTION_COPY);
                try {
                    java.util.List<File> droppedFiles = (java.util.List<File>) dtde.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);
                    if (!droppedFiles.isEmpty()) {
                        selectedFile = droppedFiles.get(0);
                        processButton.setEnabled(true);
                        chooseFileButton.setText(selectedFile.getName());
                    }
                } catch (UnsupportedFlavorException | IOException e) {
                    JOptionPane.showMessageDialog(OrderCounter.this, "Error processing dropped file: " + e.getMessage());
                }
            }
        });
    }

    private void processFile() {
        Map<String, Map<String, Integer>> productQuantities = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(selectedFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            int variationCol = -1, productCol = -1, quantityCol = -1;

            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                for (Cell cell : headerRow) {
                    String header = cell.getStringCellValue();
                    if ("Variation Name".equalsIgnoreCase(header)) variationCol = cell.getColumnIndex();
                    else if ("Product Name".equalsIgnoreCase(header)) productCol = cell.getColumnIndex();
                    else if ("Quantity".equalsIgnoreCase(header)) quantityCol = cell.getColumnIndex();
                }
            }

            if (productCol == -1 || variationCol == -1 || quantityCol == -1) {
                JOptionPane.showMessageDialog(this, "Error: Could not find required columns in the file.");
                return;
            }

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;

                String product = getStringCellValue(row, productCol);
                String variation = getStringCellValue(row, variationCol);
                int quantity = getNumericCellValue(row, quantityCol);

                productQuantities
                        .computeIfAbsent(product, k -> new HashMap<>())
                        .merge(variation, quantity, Integer::sum);
            }

            displayResults(productQuantities);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(this, "Error reading file: " + e.getMessage());
        }
    }

    private void displayResults(Map<String, Map<String, Integer>> productQuantities) {
        DefaultTableModel model = (DefaultTableModel) resultsTable.getModel();
        model.setRowCount(0);
        for (Map.Entry<String, Map<String, Integer>> productEntry : productQuantities.entrySet()) {
            String productName = productEntry.getKey();
            for (Map.Entry<String, Integer> variationEntry : productEntry.getValue().entrySet()) {
                String variationName = variationEntry.getKey();
                Integer quantity = variationEntry.getValue();
                model.addRow(new Object[]{productName, variationName, quantity});
            }
        }
    }

    private String getStringCellValue(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        return cell != null && cell.getCellType() == CellType.STRING ? cell.getStringCellValue() : "";
    }

    private int getNumericCellValue(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) return 0;
        try {
            return cell.getCellType() == CellType.NUMERIC ? (int) cell.getNumericCellValue()
                    : Integer.parseInt(cell.getStringCellValue());
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            OrderCounter app = new OrderCounter();
            app.setVisible(true);
        });
    }
}
