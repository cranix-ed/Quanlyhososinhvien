/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JPanel.java to edit this template
 */
package QuanLyHoSoSinhVien;

import ConnectDatabase.ConnectDB;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Vector;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Cranux
 */
public class qlhocphi extends javax.swing.JPanel {

    /**
     * Creates new form qlmonhoc
     */
    public qlhocphi() {
        initComponents();
        load_hocphi();
        load_sv();
        txtHPdadong.setText("0");
    }

    public JTable getTblHocphi() {
        return tblHocphi;
    }
    
    public void reload() {
         tblHocphi.setModel(new DefaultTableModel());
        load_hocphi();
    }
    
    Connection conn = null;
    Map<String, String> masv = new HashMap<>();
    Map<String, String> tensv = new HashMap<>();
    private void load_sv() {
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT * FROM sinhvien";
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(sql);

            while (rs.next()) {
                tensv.put(rs.getString("masinhvien"), rs.getString("hoten"));
                masv.put(rs.getString("masinhvien"), rs.getString("idsinhvien"));
            }

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void load_hocphi() {
        try {
            conn = ConnectDB.KetnoiDB();

            Statement statement = conn.createStatement();
            String query = "SELECT \n"
                    + "    hocphi.idhocphi,\n"
                    + "    sinhvien.masinhvien AS masinhvien,\n"
                    + "    sinhvien.hoten AS tensinhvien,\n"
                    + "    hocphi.tinchi,\n"
                    + "    hocphi.tonghocphi,\n"
                    + "    hocphi.hocphidadong,\n"
                    + "    hocphi.nohocphi,\n"
                    + "    hocphi.hanchot,\n"
                    + "    hocphi.trangthai\n"
                    + "FROM \n"
                    + "    hocphi\n"
                    + "JOIN \n"
                    + "    sinhvien ON hocphi.idsinhvien = sinhvien.idsinhvien";
            ResultSet resultset = statement.executeQuery(query);

//            tblHocphi.removeAll();
            String[] tdb = {"Mã học phí","Mã sinh viên", "Tên sinh viên", "Tín chỉ đăng ký", "Tổng học phí", "Học phí đã đóng", "Nợ học phí", "Hạn đóng", "Trạng thái"};
            DefaultTableModel model = new DefaultTableModel(tdb, 0);

            int i = 0;
            while (resultset.next()) {
                
                Vector vector = new Vector();
                
                vector.add(resultset.getString("idhocphi"));
                vector.add(resultset.getString("masinhvien"));
                vector.add(resultset.getString("tensinhvien"));
                vector.add(resultset.getString("tinchi"));
                vector.add(resultset.getString("tonghocphi"));
                vector.add(resultset.getString("hocphidadong"));
                vector.add(resultset.getString("nohocphi"));
                vector.add(resultset.getString("hanchot"));
                vector.add(resultset.getString("trangthai"));
                model.addRow(vector);
            }
            tblHocphi.setModel(model);
            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        txtMahocphi.setEnabled(false);
    }
    
    private static CellStyle DinhdangHeader(XSSFSheet sheet) {
        // Create font
        Font font = sheet.getWorkbook().createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 12); // font size
        font.setColor(IndexedColors.WHITE.getIndex()); // text color

        // Create CellStyle
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        cellStyle.setFillForegroundColor(IndexedColors.DARK_GREEN.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setWrapText(true);
        return cellStyle;
    }
    
    private void ReadExcel(String tenfilepath) {
        try {
            FileInputStream fis = new FileInputStream(tenfilepath);
            //Tạo đối tượng Excel
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0); //Lấy sheet đầu tiên của file
            //Lấy ra các dòng bảng bảng
            Iterator<Row> itr = sheet.iterator();
            //Đọc dữ liệu
            itr.next();
            while (itr.hasNext()) {//Lặp đến hết các dòng trong excel
                Row row = itr.next();//Lấy dòng tiếp theo
                double idhp, hpdong;
                String masinhvien, tensinhvien, sotinchi, hocphidadong;
                Date hanchot;
//                Date ngs;

                masinhvien = row.getCell(0).getStringCellValue();
                
                tensinhvien = row.getCell(1).getStringCellValue();
                idhp = row.getCell(2).getNumericCellValue();
                sotinchi = String.valueOf((int) idhp);
                hpdong = row.getCell(3).getNumericCellValue();
                hocphidadong = String.valueOf((int) hpdong);
                hanchot = new Date(row.getCell(4).getDateCellValue().getTime());
                

                Themhocphi(masinhvien, tensinhvien, sotinchi, hocphidadong, hanchot);
            }
        } catch (Exception e) {
            e.printStackTrace();
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
        jLabel1 = new javax.swing.JLabel();
        jPanel2 = new javax.swing.JPanel();
        btThem = new javax.swing.JButton();
        btSua = new javax.swing.JButton();
        btXoa = new javax.swing.JButton();
        btXuatExcel = new javax.swing.JButton();
        btNhapExcel = new javax.swing.JButton();
        jPanel3 = new javax.swing.JPanel();
        jComboBox1 = new javax.swing.JComboBox<>();
        jTextField1 = new javax.swing.JTextField();
        jPanel4 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        tblHocphi = new javax.swing.JTable();
        jPanel5 = new javax.swing.JPanel();
        txtNghandong = new com.toedter.calendar.JDateChooser();
        jLabel2 = new javax.swing.JLabel();
        txtTenSV = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        txtHPdadong = new javax.swing.JTextField();
        jLabel12 = new javax.swing.JLabel();
        txtMaSV = new javax.swing.JTextField();
        jLabel13 = new javax.swing.JLabel();
        txtMahocphi = new javax.swing.JTextField();

        jLabel1.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("QUẢN LÝ HỌC PHÍ");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(426, 426, 426)
                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 243, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, 43, Short.MAX_VALUE)
                .addContainerGap())
        );

        jPanel2.setBorder(javax.swing.BorderFactory.createTitledBorder("Thao tác"));

        btThem.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\Documents\\NetBeansProjects\\BTL_Nhom4\\src\\main\\resources\\image\\icons8_plus_+_48px_1.png")); // NOI18N
        btThem.setText("Thêm");
        btThem.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        btThem.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        btThem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btThemActionPerformed(evt);
            }
        });

        btSua.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\Documents\\NetBeansProjects\\BTL_Nhom4\\src\\main\\resources\\image\\icons8_edit_property_48px.png")); // NOI18N
        btSua.setText("Sửa");
        btSua.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        btSua.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        btSua.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btSuaActionPerformed(evt);
            }
        });

        btXoa.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\Documents\\NetBeansProjects\\BTL_Nhom4\\src\\main\\resources\\image\\icons8_trash_can_48px.png")); // NOI18N
        btXoa.setText("Xóa");
        btXoa.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        btXoa.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        btXoa.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btXoaActionPerformed(evt);
            }
        });

        btXuatExcel.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\OneDrive\\Hình ảnh\\Icon\\Fatcow-Farm-Fresh-Excel-exports.32.png")); // NOI18N
        btXuatExcel.setText("Xuất excel");
        btXuatExcel.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        btXuatExcel.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        btXuatExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btXuatExcelActionPerformed(evt);
            }
        });

        btNhapExcel.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\OneDrive\\Hình ảnh\\Icon\\Fatcow-Farm-Fresh-Excel-imports.32.png")); // NOI18N
        btNhapExcel.setText("Nhập excel");
        btNhapExcel.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        btNhapExcel.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        btNhapExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btNhapExcelActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(btThem)
                .addGap(18, 18, 18)
                .addComponent(btSua)
                .addGap(18, 18, 18)
                .addComponent(btXoa)
                .addGap(18, 18, 18)
                .addComponent(btXuatExcel)
                .addGap(18, 18, 18)
                .addComponent(btNhapExcel)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(btNhapExcel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btXuatExcel, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, jPanel2Layout.createSequentialGroup()
                        .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(btSua)
                            .addComponent(btThem)
                            .addComponent(btXoa))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );

        jPanel3.setBorder(javax.swing.BorderFactory.createTitledBorder("Tìm kiếm"));

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jTextField1.setText("jTextField1");

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addGap(27, 27, 27)
                .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(114, 114, 114)
                .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(344, Short.MAX_VALUE))
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(16, Short.MAX_VALUE))
        );

        tblHocphi.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        tblHocphi.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                tblHocphiMouseClicked(evt);
            }
        });
        jScrollPane1.setViewportView(tblHocphi);

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane1)
                .addContainerGap())
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 356, Short.MAX_VALUE)
                .addContainerGap())
        );

        jPanel5.setBorder(javax.swing.BorderFactory.createTitledBorder("Nhập dữ liệu"));

        jLabel2.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel2.setText("Tên sinh viên");

        jLabel3.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel3.setText("Hạn đóng học phí");

        jLabel7.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel7.setText("Học phí đã đóng");

        jLabel12.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel12.setText("Mã sinh viên");

        txtMaSV.addInputMethodListener(new java.awt.event.InputMethodListener() {
            public void caretPositionChanged(java.awt.event.InputMethodEvent evt) {
                txtMaSVCaretPositionChanged(evt);
            }
            public void inputMethodTextChanged(java.awt.event.InputMethodEvent evt) {
                txtMaSVInputMethodTextChanged(evt);
            }
        });
        txtMaSV.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtMaSVActionPerformed(evt);
            }
        });

        jLabel13.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel13.setText("Mã học phí");

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addComponent(jLabel13, javax.swing.GroupLayout.PREFERRED_SIZE, 91, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(txtMahocphi, javax.swing.GroupLayout.PREFERRED_SIZE, 179, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addComponent(jLabel7)
                        .addGap(18, 18, 18)
                        .addComponent(txtHPdadong, javax.swing.GroupLayout.PREFERRED_SIZE, 180, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addComponent(jLabel12)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(txtMaSV, javax.swing.GroupLayout.PREFERRED_SIZE, 179, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel5Layout.createSequentialGroup()
                        .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 121, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(28, 28, 28)
                        .addComponent(txtNghandong, javax.swing.GroupLayout.PREFERRED_SIZE, 152, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(61, 61, 61)
                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(txtTenSV, javax.swing.GroupLayout.PREFERRED_SIZE, 190, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(69, 69, 69))
        );
        jPanel5Layout.setVerticalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtTenSV, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 34, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel12, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtMaSV, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel13, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtMahocphi, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(20, 20, 20)
                .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(jLabel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(txtHPdadong, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(txtNghandong, javax.swing.GroupLayout.DEFAULT_SIZE, 34, Short.MAX_VALUE))
                .addContainerGap(16, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addComponent(jPanel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanel5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
    }// </editor-fold>//GEN-END:initComponents

    private void btThemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btThemActionPerformed
        // TODO add your handling code here:
        String msv = txtMaSV.getText().trim();
        String idsv = masv.get(msv);
        String tensv = txtTenSV.getText().trim();
        
        Double hp = Double.parseDouble(txtHPdadong.getText().trim());
        Date ngbatdau = new Date(txtNghandong.getDate().getTime());
        

//        if (!checkTenlop()) {
//            JOptionPane.showMessageDialog(this, "Lớp học đã tồn tại");
//            return;
//        }
        try {
            conn = ConnectDB.KetnoiDB();
//            String sql = "Insert Tacgia values('" + mtg + "',N'" + ttg + "','" + ngs + "',N'" + gt + "',"
//                    + "'" + dt + "','" + email + "',N'" + dc + "')";
            String sqli = "INSERT INTO hocphi (idsinhvien, hocphidadong, hanchot) VALUES (?,?,?)";
            PreparedStatement st = conn.prepareStatement(sqli);

            st.setString(1, idsv);
            st.setDouble(2, hp);
            st.setDate(3, ngbatdau);

            st.execute();

            conn.close();
            JOptionPane.showMessageDialog(this, "Thêm mới thành công");
            load_hocphi();
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btThemActionPerformed

    private void txtMaSVActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtMaSVActionPerformed
        // TODO add your handling code here:
//        String msv = evt.getActionCommand();
        String msv = txtMaSV.getText().trim();
        txtTenSV.setText(tensv.get(msv));
    }//GEN-LAST:event_txtMaSVActionPerformed

    private void txtMaSVInputMethodTextChanged(java.awt.event.InputMethodEvent evt) {//GEN-FIRST:event_txtMaSVInputMethodTextChanged
        // TODO add your handling code here:

    }//GEN-LAST:event_txtMaSVInputMethodTextChanged

    private void tblHocphiMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_tblHocphiMouseClicked
        // TODO add your handling code here:
        int i = tblHocphi.getSelectedRow();
        DefaultTableModel tb = (DefaultTableModel) tblHocphi.getModel();
        txtMahocphi.setText(tb.getValueAt(i, 0).toString());
        txtMaSV.setText(tb.getValueAt(i, 1).toString());
        txtTenSV.setText(tb.getValueAt(i, 2).toString());
        txtHPdadong.setText(tb.getValueAt(i, 5).toString());
        String handong = tb.getValueAt(i, 7).toString();
        java.util.Date ngh;
        try {
            ngh = new SimpleDateFormat("yyyy-MM-dd").parse(handong);
            txtNghandong.setDate(ngh);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_tblHocphiMouseClicked

    private void btSuaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btSuaActionPerformed
        // TODO add your handling code here:
        String idhp = txtMahocphi.getText().trim();
        String msv = txtMaSV.getText().trim();
        String idsv = masv.get(msv);
        String tensv = txtTenSV.getText().trim();
        
        Double hp = Double.parseDouble(txtHPdadong.getText().trim());
        Date ngbatdau = new Date(txtNghandong.getDate().getTime());
        
        try {
            conn = ConnectDB.KetnoiDB();
//            String sql = "UPDATE tacgia SET tentacgia=N'" + ttg + "',ngaysinh='" + ngs + "',gioitinh=N'" + gt + "',dienthoai=" + "'" + dt + "',email='" + email + "',diachi=N'" + dc + "' WHERE matacgia='" + mtg + "'";
            String sqli = "UPDATE hocphi SET idsinhvien=?, hocphidadong=?, hanchot=? WHERE idhocphi=?";
            PreparedStatement st = conn.prepareStatement(sqli);

            st.setString(1, idsv);
            st.setDouble(2, hp);
            st.setDate(3, ngbatdau);
            st.setString(4, idhp);

            st.execute();

            conn.close();
            JOptionPane.showMessageDialog(this, "Sửa thành công");
            load_hocphi();
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btSuaActionPerformed

    private void btXoaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btXoaActionPerformed
        // TODO add your handling code here:
        String idhp = txtMahocphi.getText().trim();
        
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "DELETE FROM hocphi WHERE idhocphi=?";
            PreparedStatement st = conn.prepareStatement(sql);
            st.setString(1, idhp);
            int reply = JOptionPane.showConfirmDialog(null, "Bạn có chắc chắn xóa không?", null, JOptionPane.YES_NO_OPTION);
            if (reply == JOptionPane.YES_OPTION) {
                st.executeUpdate();
                JOptionPane.showMessageDialog(this, "Xóa thành công");
            } else {

            }
            load_hocphi();
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btXoaActionPerformed

    private void btXuatExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btXuatExcelActionPerformed
        // TODO add your handling code here:
        try {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet spreadsheet = workbook.createSheet("lophoc");
            // register the columns you wish to track and compute the column width

            CreationHelper createHelper = workbook.getCreationHelper();

            XSSFRow row = null;
            Cell cell = null;

            row = spreadsheet.createRow((short) 2);
            row.setHeight((short) 500);
            cell = row.createCell(0, CellType.STRING);
            cell.setCellValue("DANH SÁCH HỌC PHÍ SINH VIÊN");

            //Tạo dòng tiêu đều của bảng
            // create CellStyle
            CellStyle cellStyle_Head = DinhdangHeader(spreadsheet);
            row = spreadsheet.createRow((short) 3);
            row.setHeight((short) 500);
            cell = row.createCell(0, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("STT");

            cell = row.createCell(1, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Mã khoản phí");

            cell = row.createCell(2, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Mã sinh viên");

            cell = row.createCell(3, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Tên sinh viên");

            cell = row.createCell(4, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Tín chỉ đăng ký");

            cell = row.createCell(5, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Tổng học phí");

            cell = row.createCell(6, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Học phí đã đóng");
            
            cell = row.createCell(7, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Nợ học phí");
            
            cell = row.createCell(8, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Hạn đóng");
            
            cell = row.createCell(9, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Trạng thái");

            //Kết nối DB
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT \n"
                    + "    hocphi.idhocphi,\n"
                    + "    sinhvien.masinhvien AS masinhvien,\n"
                    + "    sinhvien.hoten AS tensinhvien,\n"
                    + "    hocphi.tinchi,\n"
                    + "    hocphi.tonghocphi,\n"
                    + "    hocphi.hocphidadong,\n"
                    + "    hocphi.nohocphi,\n"
                    + "    hocphi.hanchot,\n"
                    + "    hocphi.trangthai\n"
                    + "FROM \n"
                    + "    hocphi\n"
                    + "JOIN \n"
                    + "    sinhvien ON hocphi.idsinhvien = sinhvien.idsinhvien";
            PreparedStatement st = conn.prepareStatement(sql);
            ResultSet rs = st.executeQuery();
            //Đổ dữ liệu từ rs vào các ô trong excel
            ResultSetMetaData rsmd = rs.getMetaData();
            int tongsocot = rsmd.getColumnCount();

            //Đinh dạng Tạo đường kẻ cho ô chứa dữ liệu
            CellStyle cellStyle_data = spreadsheet.getWorkbook().createCellStyle();
            cellStyle_data.setBorderLeft(BorderStyle.THIN);
            cellStyle_data.setBorderRight(BorderStyle.THIN);
            cellStyle_data.setBorderBottom(BorderStyle.THIN);

            int i = 0;
            while (rs.next()) {
                row = spreadsheet.createRow((short) 4 + i);
                row.setHeight((short) 400);

                cell = row.createCell(0);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(i + 1);

                cell = row.createCell(1);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("idhocphi"));

                cell = row.createCell(2);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("masinhvien"));

                //Định dạng ngày tháng trong excel
                cell = row.createCell(3);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("tensinhvien"));

                cell = row.createCell(4);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("tinchi"));
                
                cell = row.createCell(5);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("tonghocphi"));
                
                cell = row.createCell(6);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("hocphidadong"));
                
                cell = row.createCell(7);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("nohocphi"));
                
                java.util.Date hanchot = new java.util.Date(rs.getDate("hanchot").getTime());
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cell = row.createCell(8);
                cell.setCellValue(hanchot);
                cell.setCellStyle(cellStyle);

                cell = row.createCell(9);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("trangthai"));

                i++;
            }
            //Hiệu chỉnh độ rộng của cột
            for (int col = 0; col < tongsocot; col++) {
                spreadsheet.autoSizeColumn(col);
            }

            File f = new File("C:\\TestJava\\danhsachdonghocphi.xlsx");
            FileOutputStream out = new FileOutputStream(f);
            workbook.write(out);
            JOptionPane.showMessageDialog(this, "Xuất file excel thành công");
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_btXuatExcelActionPerformed

    private void btNhapExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btNhapExcelActionPerformed
        // TODO add your handling code here:
        try {
            JFileChooser fc = new JFileChooser();
            int lc = fc.showOpenDialog(this);
            if (lc == JFileChooser.APPROVE_OPTION) {
                File file = fc.getSelectedFile();
                String tenfile = file.getName();
                if (tenfile.endsWith(".xlsx")) {    //endsWith chọn file có phần kết thúc ...
                    ReadExcel(file.getPath());
                    JOptionPane.showMessageDialog(this, "Nhập file thành công");
                    load_hocphi();
                } else {
                    JOptionPane.showMessageDialog(this, "Phải chọn file excel");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_btNhapExcelActionPerformed

    private void txtMaSVCaretPositionChanged(java.awt.event.InputMethodEvent evt) {//GEN-FIRST:event_txtMaSVCaretPositionChanged
        // TODO add your handling code here:
    }//GEN-LAST:event_txtMaSVCaretPositionChanged


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btNhapExcel;
    private javax.swing.JButton btSua;
    private javax.swing.JButton btThem;
    private javax.swing.JButton btXoa;
    private javax.swing.JButton btXuatExcel;
    private javax.swing.JComboBox<String> jComboBox1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTable tblHocphi;
    private javax.swing.JTextField txtHPdadong;
    private javax.swing.JTextField txtMaSV;
    private javax.swing.JTextField txtMahocphi;
    private com.toedter.calendar.JDateChooser txtNghandong;
    private javax.swing.JTextField txtTenSV;
    // End of variables declaration//GEN-END:variables

    private void Themhocphi(String masinhvien, String tensinhvien, String sotinchi, String hocphidadong, Date hanchot) {
        String msv = masv.get(masinhvien);
        
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "INSERT INTO hocphi (idsinhvien, tinchi, hocphidadong, hanchot) VALUES (?,?,?,?)";
            PreparedStatement st = conn.prepareStatement(sql);

            st.setString(1, msv);
            st.setString(2, sotinchi);
            st.setString(3, hocphidadong);
            st.setDate(4, hanchot);
            
            st.execute();

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
