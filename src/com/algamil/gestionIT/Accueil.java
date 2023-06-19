/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.algamil.gestionIT;
import java.sql.ResultSet;
import static com.algamil.BDD.Configuration.*;
import java.awt.Desktop;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.text.MessageFormat;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.OrientationRequested;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import net.proteanit.sql.DbUtils;
import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.type.OrientationEnum;
import net.sf.jasperreports.view.JasperViewer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author FaridMO
 */
public final class Accueil extends javax.swing.JFrame {
    ResultSet res;
    static String SQL;
    
    String idB,table="postepc", colonnes="department,service,sector,typeH,name,hostname";
    String depart,service,sector,typeH,name,hostname;
    String[] colonne = {"department","service","sector","typeH","name","hostname"};
    String cols="pid,department,service,sector,typeH,name,hostname";
    String dp,dp1;

    
    /**
     * Creates new form Accueil
     */
    public Accueil() {
        //run();
        initComponents();
        //maTable.getColumnModel().getColumn(0).setPreferredWidth(200);
        //setLocationRelativeTo(null);
        setResizable(false);
        deleteB.setEnabled(false);
        editB.setEnabled(false);
        
        table();

    }
    
    /**
     * Methode Permettant de charger et recupérer à chaque lancement du programme les données de la BD
     */
    private void table(){
        String[] col = {"pid","department","service","sector","typeH","name","hostname"};
        res = select(col, table);
        maTable.setModel(DbUtils.resultSetToTableModel(res));
    }
    
    public void recuperation(){
        depart = departC.getSelectedItem().toString();
        service = ServiceC.getSelectedItem().toString();
        sector = SectorC.getSelectedItem().toString();
        typeH = TypeC.getSelectedItem().toString();
        name = nameC.getText();
        hostname = hostnameC.getText();
     }
    
    public void actualiser(){
        departC.setSelectedIndex(0);
        ServiceC.setSelectedIndex(0);
        SectorC.setSelectedIndex(0);
        TypeC.setSelectedIndex(0);
        nameC.setText("");
        hostnameC.setText("");
    }
    
    public void viderTable(){
        SQL="Truncate table "+table;
        executionUpdate(SQL);
        table();
    }
    
    public void writeToExcell(JTable table, Path path) throws FileNotFoundException, IOException {
            new WorkbookFactory();
            Workbook wb = new XSSFWorkbook(); //Excell workbook
            Sheet sheet = wb.createSheet(); //WorkSheet
            Row row = sheet.createRow(2); //Row created at line 3
            DefaultTableModel model = (DefaultTableModel) maTable.getModel();
            
            

            Row headerRow = sheet.createRow(0); //Create row at line 0
            for(int headings = 0; headings < model.getColumnCount(); headings++){ //For each column
                headerRow.createCell(headings).setCellValue(model.getColumnName(headings));//Write column name
            }

            for(int rows = 0; rows < model.getRowCount(); rows++){ //For each table row
                for(int cols = 0; cols < table.getColumnCount(); cols++){ //For each table column
                    row.createCell(cols).setCellValue(model.getValueAt(rows, cols).toString()); //Write value
                }

                //Set the row to the next one in the sequence 
                row = sheet.createRow((rows + 3)); 
            }
           
            FileOutputStream fos = new FileOutputStream(path.toString());
            if(fos!=null){
                JOptionPane.showMessageDialog(this, "Fichier exporté avec succès !");
                wb.write(fos);
                fos.close();

            }else{
               JOptionPane.showMessageDialog(this, "Erreur lors de l'exportation du fichier ! !");
            }   
}
    
    private void openFile(){
        JFileChooser fchoose = new JFileChooser();
        fchoose.showOpenDialog(this);
        
       try{
            File f = fchoose.getSelectedFile();
            String path = f.getAbsolutePath();
            path = path+".xlsx";
            pathFile.setText(path);
       }catch(Exception e){
           JOptionPane.showMessageDialog(this,"Fichier non crée !");
           
       }
        
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jPanel2 = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jPanel3 = new javax.swing.JPanel();
        jPanel4 = new javax.swing.JPanel();
        toPDF = new javax.swing.JButton();
        viewJ = new javax.swing.JButton();
        toExcel = new javax.swing.JButton();
        pathFile = new javax.swing.JTextField();
        jLabel8 = new javax.swing.JLabel();
        searchB1 = new javax.swing.JTextField();
        saveB = new javax.swing.JButton();
        searchB2 = new javax.swing.JTextField();
        jLabel9 = new javax.swing.JLabel();
        jLabel10 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        departC = new javax.swing.JComboBox<>();
        ServiceC = new javax.swing.JComboBox<>();
        SectorC = new javax.swing.JComboBox<>();
        TypeC = new javax.swing.JComboBox<>();
        nameC = new javax.swing.JTextField();
        hostnameC = new javax.swing.JTextField();
        jScrollPane1 = new javax.swing.JScrollPane();
        maTable = new javax.swing.JTable();
        addB = new javax.swing.JButton();
        editB = new javax.swing.JButton();
        deleteB = new javax.swing.JButton();
        jButton1 = new javax.swing.JButton();
        cancelB = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBackground(new java.awt.Color(0, 0, 0));

        jPanel2.setBackground(new java.awt.Color(34, 29, 40));
        jPanel2.setBorder(new javax.swing.border.SoftBevelBorder(javax.swing.border.BevelBorder.RAISED));

        jLabel1.setBackground(new java.awt.Color(0, 0, 51));
        jLabel1.setFont(new java.awt.Font("Franklin Gothic Heavy", 1, 24)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(51, 156, 218));
        jLabel1.setText("GESTION  IT");
        jLabel1.setBorder(new javax.swing.border.MatteBorder(null));

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 177, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(439, 439, 439))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(16, 16, 16)
                .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, 40, Short.MAX_VALUE)
                .addGap(16, 16, 16))
        );

        jPanel3.setBackground(new java.awt.Color(0, 0, 51));

        jPanel4.setBackground(new java.awt.Color(0, 102, 102));
        jPanel4.setBorder(javax.swing.BorderFactory.createTitledBorder("Options"));

        toPDF.setIcon(new javax.swing.ImageIcon(getClass().getResource("/com/algamil/images/icons8-print-24.png"))); // NOI18N
        toPDF.setText("Print");
        toPDF.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                toPDFActionPerformed(evt);
            }
        });

        viewJ.setIcon(new javax.swing.ImageIcon(getClass().getResource("/com/algamil/images/icons8-detective-24.png"))); // NOI18N
        viewJ.setText("View");
        viewJ.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                viewJActionPerformed(evt);
            }
        });

        toExcel.setIcon(new javax.swing.ImageIcon(getClass().getResource("/com/algamil/images/icons8-excel-24.png"))); // NOI18N
        toExcel.setText("Excel");
        toExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                toExcelActionPerformed(evt);
            }
        });

        pathFile.setEditable(false);
        pathFile.setBackground(new java.awt.Color(204, 204, 204));
        pathFile.setEnabled(false);

        jLabel8.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jLabel8.setForeground(new java.awt.Color(255, 255, 255));
        jLabel8.setText("Search :");

        searchB1.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyReleased(java.awt.event.KeyEvent evt) {
                searchB1KeyReleased(evt);
            }
        });

        saveB.setIcon(new javax.swing.ImageIcon(getClass().getResource("/com/algamil/images/icons8-add-file-24.png"))); // NOI18N
        saveB.setText("Create File");
        saveB.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                saveBActionPerformed(evt);
            }
        });

        searchB2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                searchB2ActionPerformed(evt);
            }
        });
        searchB2.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyReleased(java.awt.event.KeyEvent evt) {
                searchB2KeyReleased(evt);
            }
        });

        jLabel9.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        jLabel9.setForeground(new java.awt.Color(255, 255, 255));
        jLabel9.setText("By Hostname");

        jLabel10.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        jLabel10.setForeground(new java.awt.Color(255, 255, 255));
        jLabel10.setText("By Department");

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(pathFile, javax.swing.GroupLayout.PREFERRED_SIZE, 272, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(7, 7, 7)
                .addComponent(saveB)
                .addGap(103, 103, 103)
                .addComponent(jLabel10, javax.swing.GroupLayout.PREFERRED_SIZE, 105, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(98, 98, 98)
                .addComponent(jLabel9, javax.swing.GroupLayout.PREFERRED_SIZE, 105, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addGap(6, 6, 6)
                .addComponent(toPDF, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(toExcel, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(viewJ, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(98, 98, 98)
                .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 64, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(28, 28, 28)
                .addComponent(searchB1, javax.swing.GroupLayout.PREFERRED_SIZE, 120, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(59, 59, 59)
                .addComponent(searchB2, javax.swing.GroupLayout.PREFERRED_SIZE, 139, javax.swing.GroupLayout.PREFERRED_SIZE))
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addGap(2, 2, 2)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(saveB, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(pathFile, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addGap(23, 23, 23)
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel10)
                            .addComponent(jLabel9))))
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(toPDF)
                    .addComponent(toExcel)
                    .addComponent(viewJ)
                    .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addGap(1, 1, 1)
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(searchB1, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(searchB2, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)))))
        );

        jLabel2.setFont(new java.awt.Font("Arial", 1, 14)); // NOI18N
        jLabel2.setForeground(new java.awt.Color(255, 255, 255));
        jLabel2.setText("Department :");

        jLabel3.setFont(new java.awt.Font("Arial", 1, 14)); // NOI18N
        jLabel3.setForeground(new java.awt.Color(255, 255, 255));
        jLabel3.setText("Service :");

        jLabel4.setFont(new java.awt.Font("Arial", 1, 14)); // NOI18N
        jLabel4.setForeground(new java.awt.Color(255, 255, 255));
        jLabel4.setText("Sector :");

        jLabel5.setFont(new java.awt.Font("Arial", 1, 14)); // NOI18N
        jLabel5.setForeground(java.awt.Color.white);
        jLabel5.setText("Type :");

        jLabel6.setFont(new java.awt.Font("Arial", 1, 14)); // NOI18N
        jLabel6.setForeground(java.awt.Color.white);
        jLabel6.setText("Name :");

        jLabel7.setFont(new java.awt.Font("Arial", 1, 14)); // NOI18N
        jLabel7.setForeground(java.awt.Color.white);
        jLabel7.setText("Hostname :");

        departC.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Choisir...", "Commerciale", "Administration", "Comptabilité", "Quincaillerie", "Hypermarché", "AlGamil2", "30Mile", "Concassage", "Magasin" }));

        ServiceC.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Choisir...", "PDG", "Achat", "Development", "Bureautique", "Comptabilité", "Facture", "DG", "HR", "Finance", "Engineering", "FrontSide", "MiddleSide", "BackSide", "OutSide", "2nd Floor", "3rd Floor", "Clientel", "Papeterie", "Zebra", "Meuble", "Stock", "Trésorerie", "EnGros", "Reception" }));

        SectorC.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Choisir...", "PDG", "Achat", "Development", "Bureautique", "Facture", "DG", "HR", "Recouvrement", "Communication", "Trésorerie", "Civil", "GPS", "IT", "Control", "Caisse", "Electricity", "Post16", "ZebraHM", "ZebraQLC", "Vendeur", "Pauto", "Stock", "Paint", "Encadrement", "Reception1", "Reception2", "Meuble", "Post25", "Intern", "Extern", "EnGros", "Aluminium", "Standard" }));

        TypeC.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Choisir...", "Desktop", "Laptop" }));

        maTable.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null, null, null, null},
                {null, null, null, null, null, null, null},
                {null, null, null, null, null, null, null},
                {null, null, null, null, null, null, null}
            },
            new String [] {
                "PID", "Department", "Service", "Sector", "Type", "Name", "Hostname"
            }
        ) {
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false
            };

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        maTable.setColumnSelectionAllowed(true);
        maTable.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        maTable.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                maTableMouseClicked(evt);
            }
        });
        jScrollPane1.setViewportView(maTable);
        maTable.getColumnModel().getSelectionModel().setSelectionMode(javax.swing.ListSelectionModel.SINGLE_SELECTION);
        if (maTable.getColumnModel().getColumnCount() > 0) {
            maTable.getColumnModel().getColumn(0).setResizable(false);
            maTable.getColumnModel().getColumn(0).setPreferredWidth(2);
            maTable.getColumnModel().getColumn(1).setResizable(false);
            maTable.getColumnModel().getColumn(1).setPreferredWidth(6);
            maTable.getColumnModel().getColumn(2).setResizable(false);
            maTable.getColumnModel().getColumn(2).setPreferredWidth(6);
            maTable.getColumnModel().getColumn(3).setResizable(false);
            maTable.getColumnModel().getColumn(3).setPreferredWidth(6);
            maTable.getColumnModel().getColumn(4).setResizable(false);
            maTable.getColumnModel().getColumn(4).setPreferredWidth(5);
            maTable.getColumnModel().getColumn(5).setResizable(false);
            maTable.getColumnModel().getColumn(5).setPreferredWidth(5);
            maTable.getColumnModel().getColumn(6).setResizable(false);
            maTable.getColumnModel().getColumn(6).setPreferredWidth(5);
        }

        addB.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        addB.setText("Add Host");
        addB.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                addBActionPerformed(evt);
            }
        });

        editB.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        editB.setText("Edit");
        editB.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                editBActionPerformed(evt);
            }
        });

        deleteB.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        deleteB.setText("Delete");
        deleteB.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                deleteBActionPerformed(evt);
            }
        });

        jButton1.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        jButton1.setText("Clear Table");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        cancelB.setFont(new java.awt.Font("Segoe UI", 1, 12)); // NOI18N
        cancelB.setText("Cancel");
        cancelB.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cancelBActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addGap(20, 20, 20)
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel4)
                            .addComponent(jLabel5)
                            .addComponent(jLabel6)
                            .addComponent(jLabel7))
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGap(19, 19, 19)
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(SectorC, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(TypeC, javax.swing.GroupLayout.PREFERRED_SIZE, 118, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(nameC, javax.swing.GroupLayout.PREFERRED_SIZE, 118, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel3Layout.createSequentialGroup()
                                        .addGap(82, 82, 82)
                                        .addComponent(addB)
                                        .addGap(37, 37, 37)
                                        .addComponent(editB)
                                        .addGap(38, 38, 38)
                                        .addComponent(deleteB)
                                        .addGap(39, 39, 39)
                                        .addComponent(jButton1)
                                        .addGap(27, 27, 27)
                                        .addComponent(cancelB))
                                    .addGroup(jPanel3Layout.createSequentialGroup()
                                        .addGap(33, 33, 33)
                                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(jScrollPane1)))))
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGap(18, 18, 18)
                                .addComponent(hostnameC, javax.swing.GroupLayout.PREFERRED_SIZE, 118, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 0, Short.MAX_VALUE))))
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addComponent(jLabel3)
                                .addGap(37, 37, 37)
                                .addComponent(ServiceC, javax.swing.GroupLayout.PREFERRED_SIZE, 118, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGap(2, 2, 2)
                                .addComponent(jLabel2)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(departC, javax.swing.GroupLayout.PREFERRED_SIZE, 118, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGap(30, 30, 30)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(departC, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel2))
                        .addGap(38, 38, 38)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(ServiceC, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel3)))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel3Layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jPanel4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGap(11, 11, 11)
                        .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 215, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(28, 28, 28)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(addB)
                            .addComponent(editB)
                            .addComponent(deleteB)
                            .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                .addComponent(jButton1)
                                .addComponent(cancelB))))
                    .addGroup(jPanel3Layout.createSequentialGroup()
                        .addGap(38, 38, 38)
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel3Layout.createSequentialGroup()
                                .addGap(67, 67, 67)
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                    .addComponent(jLabel5)
                                    .addComponent(TypeC, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGap(38, 38, 38)
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                    .addComponent(jLabel6)
                                    .addComponent(nameC, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGap(48, 48, 48)
                                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                    .addComponent(jLabel7)
                                    .addComponent(hostnameC, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                            .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                .addComponent(SectorC, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addComponent(jLabel4)))))
                .addContainerGap(88, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addContainerGap())
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void addBActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_addBActionPerformed
        recuperation();
        String valeurs[]={depart,service,sector,typeH,name,hostname};
        
        if(departC.getSelectedIndex()!=0 && ServiceC.getSelectedIndex() !=0 && SectorC.getSelectedIndex()!=0 && TypeC.getSelectedIndex()!=0 &&
            nameC.getText()!=null && hostnameC.getText()!=null    ){
            
            insert(table, colonne, valeurs);
            System.out.println("Poste Ajouté");
        }else{
            JOptionPane.showMessageDialog(this, "Veuillez remplir tous les champs svp !");
        }
         actualiser();
         table();
         
    }//GEN-LAST:event_addBActionPerformed

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        viderTable();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void maTableMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_maTableMouseClicked
        DefaultTableModel model = (DefaultTableModel) maTable.getModel();
        int selectedRow = maTable.getSelectedRow();
        departC.setSelectedItem(model.getValueAt(selectedRow, 1).toString());
        ServiceC.setSelectedItem(model.getValueAt(selectedRow, 2).toString());
        SectorC.setSelectedItem(model.getValueAt(selectedRow, 3).toString());
        TypeC.setSelectedItem(model.getValueAt(selectedRow, 4).toString());
        nameC.setText(model.getValueAt(selectedRow, 5).toString());
        hostnameC.setText(model.getValueAt(selectedRow, 6).toString());
        editB.setEnabled(true);
        deleteB.setEnabled(true);
        addB.setEnabled(false);
    }//GEN-LAST:event_maTableMouseClicked

    private void editBActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_editBActionPerformed
        DefaultTableModel model = (DefaultTableModel) maTable.getModel();
        recuperation();
        idB=String.valueOf(model.getValueAt(maTable.getSelectedRow(), 0));
        String[] colon = {"department","service","sector","typeH","name","hostname"};
        String val[]={depart,service,sector,typeH,name,hostname};
        if(departC.getSelectedIndex()!=0 && ServiceC.getSelectedIndex() !=0 && SectorC.getSelectedIndex()!=0 && TypeC.getSelectedIndex()!=0 &&
            nameC.getText()!=null && hostnameC.getText()!=null){
            updateTable(table, colon,val, idB);
            JOptionPane.showMessageDialog(this,"Poste Modifié !");
        }else{
            JOptionPane.showMessageDialog(this, "Veuillez remplir tous les champs svp !");
        }
        actualiser();
        table();
        
    }//GEN-LAST:event_editBActionPerformed

    private void cancelBActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cancelBActionPerformed
                actualiser();
                addB.setEnabled(true);
                editB.setEnabled(false);
                deleteB.setEnabled(false);
                pathFile.setText(null);
                table();
    }//GEN-LAST:event_cancelBActionPerformed

    private void deleteBActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_deleteBActionPerformed
        DefaultTableModel model = (DefaultTableModel)maTable.getModel();
        int ligne = maTable.getSelectedRow();
        String idBC = String.valueOf(model.getValueAt(ligne, 0));
        if(maTable.getSelectedRowCount()==1){
            deleteB.setEnabled(true);
            deleteColumn(table, idBC);
            System.out.println("Poste Supprimé !!!!!");
            idReload(table);
            if(maTable.getRowCount()==0){
                JOptionPane.showMessageDialog(this, "Table Vide !!!");         
            }
        }else{
            deleteB.setEnabled(false);
        }
        
        actualiser();
        table();
    }//GEN-LAST:event_deleteBActionPerformed

    private void toPDFActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_toPDFActionPerformed
        MessageFormat header = new MessageFormat("Liste des postes informatiques");
        MessageFormat footer = new MessageFormat("GroupAlGamil");
        try {
            PrintRequestAttributeSet set = new HashPrintRequestAttributeSet();
            set.add(OrientationRequested.PORTRAIT);
            maTable.print(JTable.PrintMode.FIT_WIDTH,header,footer,true,set,true);
          
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Erreur d'impression !");
        }
    }//GEN-LAST:event_toPDFActionPerformed

    private void toExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_toExcelActionPerformed
        Desktop desktop = Desktop.getDesktop();
        File file = new File(pathFile.getText().toString());
        Path path = Path.of( pathFile.getText().toString());
        
        try {
            writeToExcell(maTable, path);
               if(file.exists()){
                   desktop.open(file);
                }
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        
        

        
        
    }//GEN-LAST:event_toExcelActionPerformed

    private void saveBActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_saveBActionPerformed
       openFile();
    }//GEN-LAST:event_saveBActionPerformed

    private void searchB2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_searchB2ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_searchB2ActionPerformed

    private void searchB1KeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_searchB1KeyReleased
        dp = searchB1.getText();
        if(searchB1.getText()!=null){
            maTable.setModel(DbUtils.resultSetToTableModel(searchByDepartment(table,cols,dp)));
        }else{
            table();
        }
        
        
    }//GEN-LAST:event_searchB1KeyReleased

    private void searchB2KeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_searchB2KeyReleased
        dp1 = searchB2.getText();
        if(searchB2.getText()!=null){
            maTable.setModel(DbUtils.resultSetToTableModel(searchByHostname(table,cols,dp1)));
        }else{
            table();
        }
    }//GEN-LAST:event_searchB2KeyReleased

    private void viewJActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_viewJActionPerformed
        String reportpath = "C:\\Users\\FaridMO\\Documents\\NetBeansProjects\\GestIT\\report.jrxml";
        
        try {
            JasperReport jr = JasperCompileManager.compileReport(reportpath);
            JasperPrint jp = JasperFillManager.fillReport(jr, null, dbConn());
            JasperViewer.viewReport(jp);
            closeConn();
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(rootPane, ex);
        }
    }//GEN-LAST:event_viewJActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Accueil.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Accueil.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Accueil.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Accueil.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Accueil().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JComboBox<String> SectorC;
    private javax.swing.JComboBox<String> ServiceC;
    private javax.swing.JComboBox<String> TypeC;
    private javax.swing.JButton addB;
    private javax.swing.JButton cancelB;
    private javax.swing.JButton deleteB;
    private javax.swing.JComboBox<String> departC;
    private javax.swing.JButton editB;
    private javax.swing.JTextField hostnameC;
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable maTable;
    private javax.swing.JTextField nameC;
    private javax.swing.JTextField pathFile;
    private javax.swing.JButton saveB;
    private javax.swing.JTextField searchB1;
    private javax.swing.JTextField searchB2;
    private javax.swing.JButton toExcel;
    private javax.swing.JButton toPDF;
    private javax.swing.JButton viewJ;
    // End of variables declaration//GEN-END:variables
}
