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
public class qllophoc extends javax.swing.JPanel {

    /**
     * Creates new form qllophoc
     */
    public qllophoc() {
        initComponents();
        load_CboxGiangvien();
        load_lophoc();
    }

    Connection conn = null;
    Map<String, String> giaovien = new HashMap<>();

    private void load_CboxGiangvien() {
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT * FROM giaovien";
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(sql);

            while (rs.next()) {
                CboxGiangvien.addItem(rs.getString("hoten"));
                giaovien.put(rs.getString("hoten"), rs.getString("idgiaovien"));
            }

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void load_lophoc() {
        try {
            conn = ConnectDB.KetnoiDB();

            Statement statement = conn.createStatement();
            String query = "SELECT \n"
                    + "    lophoc.idlop,\n"
                    + "    lophoc.tenlop,\n"
                    + "    giaovien.hoten AS tengiaovien, \n"
                    + "    lophoc.sosinhvien, \n"
                    + "    lophoc.ngaybatdau,\n"
                    + "    lophoc.ngayketthuc\n"
                    + "FROM \n"
                    + "    lophoc\n"
                    + "LEFT JOIN \n"
                    + "    giaovien ON lophoc.idgiaovien = giaovien.idgiaovien;";
            ResultSet resultset = statement.executeQuery(query);

            tblLophoc.removeAll();
            String[] tdb = {"Mã lớp", "Tên lớp học", "Giảng viên", "Số lượng sinh viên", "Ngày bắt đầu", "Ngày kết thúc"};
            DefaultTableModel model = new DefaultTableModel(tdb, 0);

            int i = 0;
            while (resultset.next()) {

                Vector vector = new Vector();

                vector.add(resultset.getString("idlop"));
                vector.add(resultset.getString("tenlop"));
                vector.add(resultset.getString("tengiaovien"));
                vector.add(resultset.getString("sosinhvien"));
                vector.add(resultset.getString("ngaybatdau"));
                vector.add(resultset.getString("ngayketthuc"));
                model.addRow(vector);
            }
            tblLophoc.setModel(model);
            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private boolean checkTenlop() {
        boolean kq = false;
        try {
            conn = ConnectDB.KetnoiDB();
            String tenlop = txtTenlop.getText();
            String sql = "SELECT * FROM lophoc WHERE tenlop='" + tenlop + "'";
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(sql);
            if (!rs.next()) {
                kq = true;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return kq;
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
                double idl, sosv;
                String idlop, tenlop, tengiaovien, sosinhvien;
                Date ngaybatdau, ngayketthuc;
//                Date ngs;

                idl = row.getCell(0).getNumericCellValue();
                idlop = String.valueOf((int) idl);
                tenlop = row.getCell(1).getStringCellValue();
                tengiaovien = row.getCell(2).getStringCellValue();
                sosv = row.getCell(3).getNumericCellValue();
                sosinhvien = String.valueOf((int) sosv);
                ngaybatdau = new Date(row.getCell(4).getDateCellValue().getTime());
                ngayketthuc = new Date(row.getCell(5).getDateCellValue().getTime());

                Themlophoc(idlop, tenlop, tengiaovien, sosinhvien, ngaybatdau, ngayketthuc);
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

        jPanel2 = new javax.swing.JPanel();
        jLabel2 = new javax.swing.JLabel();
        jPanel3 = new javax.swing.JPanel();
        btThem = new javax.swing.JButton();
        btSua = new javax.swing.JButton();
        btXoa = new javax.swing.JButton();
        btXuatExcel = new javax.swing.JButton();
        btNhapExcel = new javax.swing.JButton();
        jPanel4 = new javax.swing.JPanel();
        jComboBox1 = new javax.swing.JComboBox<>();
        jTextField1 = new javax.swing.JTextField();
        jPanel5 = new javax.swing.JPanel();
        jScrollPane2 = new javax.swing.JScrollPane();
        tblLophoc = new javax.swing.JTable();
        jPanel1 = new javax.swing.JPanel();
        txtNgaybatdau = new com.toedter.calendar.JDateChooser();
        jLabel1 = new javax.swing.JLabel();
        txtTenlop = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        txtSosv = new javax.swing.JTextField();
        jLabel10 = new javax.swing.JLabel();
        txtNgayketthuc = new com.toedter.calendar.JDateChooser();
        jLabel12 = new javax.swing.JLabel();
        txtMalop = new javax.swing.JTextField();
        CboxGiangvien = new javax.swing.JComboBox<>();

        jLabel2.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        jLabel2.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel2.setText("QUẢN LÝ LỚP HỌC");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(426, 426, 426)
                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 243, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jPanel3.setBorder(javax.swing.BorderFactory.createTitledBorder("Thao tác"));

        btThem.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\OneDrive\\Hình ảnh\\Icon\\add.PNG")); // NOI18N
        btThem.setText("Thêm");
        btThem.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        btThem.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        btThem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btThemActionPerformed(evt);
            }
        });

        btSua.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\OneDrive\\Hình ảnh\\Icon\\edit.PNG")); // NOI18N
        btSua.setText("Sửa");
        btSua.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        btSua.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        btSua.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btSuaActionPerformed(evt);
            }
        });

        btXoa.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\OneDrive\\Hình ảnh\\Icon\\delete.png")); // NOI18N
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

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
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
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(btNhapExcel)
                    .addComponent(btXuatExcel)
                    .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                        .addComponent(btSua)
                        .addComponent(btThem)
                        .addComponent(btXoa)))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jPanel4.setBorder(javax.swing.BorderFactory.createTitledBorder("Tìm kiếm"));

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jTextField1.setText("jTextField1");

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addGap(28, 28, 28)
                .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(118, 118, 118)
                .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(0, 16, Short.MAX_VALUE))
        );

        tblLophoc.setModel(new javax.swing.table.DefaultTableModel(
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
        tblLophoc.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                tblLophocMouseClicked(evt);
            }
        });
        jScrollPane2.setViewportView(tblLophoc);

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane2, javax.swing.GroupLayout.Alignment.TRAILING)
        );
        jPanel5Layout.setVerticalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane2, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, 363, Short.MAX_VALUE)
        );

        jLabel1.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel1.setText("Tên lớp");

        jLabel3.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel3.setText("Ngày bắt đầu");

        jLabel5.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel5.setText("Giảng viên");

        jLabel7.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel7.setText("Tổng số sinh viên");

        jLabel10.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel10.setText("Ngày kết thúc");

        jLabel12.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel12.setText("Mã lớp");

        CboxGiangvien.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Giảng viên", " " }));

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel12)
                .addGap(34, 34, 34)
                .addComponent(txtMalop, javax.swing.GroupLayout.PREFERRED_SIZE, 179, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(30, 30, 30)
                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 61, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(txtTenlop, javax.swing.GroupLayout.PREFERRED_SIZE, 190, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(37, 37, 37)
                .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 80, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(CboxGiangvien, javax.swing.GroupLayout.PREFERRED_SIZE, 181, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 55, Short.MAX_VALUE)
                .addComponent(jLabel7)
                .addGap(18, 18, 18)
                .addComponent(txtSosv, javax.swing.GroupLayout.PREFERRED_SIZE, 81, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(15, 15, 15))
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(155, 155, 155)
                .addComponent(jLabel3)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(txtNgaybatdau, javax.swing.GroupLayout.PREFERRED_SIZE, 152, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(132, 132, 132)
                .addComponent(jLabel10, javax.swing.GroupLayout.PREFERRED_SIZE, 109, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(txtNgayketthuc, javax.swing.GroupLayout.PREFERRED_SIZE, 165, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addContainerGap()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(txtTenlop, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel1)
                            .addComponent(jLabel12, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtMalop, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(10, 10, 10)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtSosv, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(CboxGiangvien, javax.swing.GroupLayout.PREFERRED_SIZE, 34, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(txtNgayketthuc, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jLabel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(txtNgaybatdau, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jLabel10, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(18, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel5, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jPanel4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(jPanel5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );
    }// </editor-fold>//GEN-END:initComponents
    private void btThemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btThemActionPerformed
        // TODO add your handling code here:
        String tenlop = txtTenlop.getText().trim();
        String gv = CboxGiangvien.getSelectedItem().toString();
        String magv = giaovien.get(gv);
        Date ngbatdau = new Date(txtNgaybatdau.getDate().getTime());
        Date ngketthuc = new Date(txtNgayketthuc.getDate().getTime());

        if (!checkTenlop()) {
            JOptionPane.showMessageDialog(this, "Lớp học đã tồn tại");
            return;
        }
        try {
            conn = ConnectDB.KetnoiDB();
//            String sql = "Insert Tacgia values('" + mtg + "',N'" + ttg + "','" + ngs + "',N'" + gt + "',"
//                    + "'" + dt + "','" + email + "',N'" + dc + "')";
            String sqli = "INSERT INTO lophoc (tenlop, idgiaovien, ngaybatdau, ngayketthuc) VALUES (?,?,?,?)";
            PreparedStatement st = conn.prepareStatement(sqli);

            st.setString(1, tenlop);
            st.setString(2, magv);
            st.setDate(3, ngbatdau);
            st.setDate(4, ngketthuc);

            st.execute();

            conn.close();
            JOptionPane.showMessageDialog(this, "Thêm mới thành công");
            load_lophoc();
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btThemActionPerformed

    private void btSuaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btSuaActionPerformed
        // TODO add your handling code here:
        String idlop = txtMalop.getText().trim();
        String tenlop = txtTenlop.getText().trim();
        String gv = CboxGiangvien.getSelectedItem().toString();
        String magv = giaovien.get(gv);
        Date ngbatdau = new Date(txtNgaybatdau.getDate().getTime());
        Date ngketthuc = new Date(txtNgayketthuc.getDate().getTime());

        if (!checkTenlop()) {
            JOptionPane.showMessageDialog(this, "Lớp học đã tồn tại");
            return;
        }
        try {
            conn = ConnectDB.KetnoiDB();
//            String sql = "UPDATE tacgia SET tentacgia=N'" + ttg + "',ngaysinh='" + ngs + "',gioitinh=N'" + gt + "',dienthoai=" + "'" + dt + "',email='" + email + "',diachi=N'" + dc + "' WHERE matacgia='" + mtg + "'";
            String sqli = "UPDATE lophoc SET tenlop=?, idgiaovien=?, ngaybatdau=?, ngayketthuc=? WHERE idlop=?";
            PreparedStatement st = conn.prepareStatement(sqli);

            st.setString(1, tenlop);
            st.setString(2, magv);
            st.setDate(3, ngbatdau);
            st.setDate(4, ngketthuc);
            st.setString(5, idlop);

            st.execute();

            conn.close();
            JOptionPane.showMessageDialog(this, "Sửa thành công");
            load_lophoc();
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btSuaActionPerformed

    private void tblLophocMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_tblLophocMouseClicked
        // TODO add your handling code here:
        int i = tblLophoc.getSelectedRow();
        DefaultTableModel tb = (DefaultTableModel) tblLophoc.getModel();
        txtMalop.setText(tb.getValueAt(i, 0).toString());
        txtTenlop.setText(tb.getValueAt(i, 1).toString());
        CboxGiangvien.setSelectedItem(tb.getValueAt(i, 2).toString());
        txtSosv.setText(tb.getValueAt(i, 3).toString());
        String ngbatdau = tb.getValueAt(i, 4).toString();
        java.util.Date ngbd;
        try {
            ngbd = new SimpleDateFormat("yyyy-MM-dd").parse(ngbatdau);
            txtNgaybatdau.setDate(ngbd);
        } catch (Exception e) {
            e.printStackTrace();
        }

        String ngketthuc = tb.getValueAt(i, 5).toString();
        java.util.Date ngkt;
        try {
            ngkt = new SimpleDateFormat("yyyy-MM-dd").parse(ngketthuc);
            txtNgayketthuc.setDate(ngkt);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_tblLophocMouseClicked

    private void btXoaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btXoaActionPerformed
        // TODO add your handling code here:
        String malop = txtMalop.getText();
        try {
            conn = ConnectDB.KetnoiDB();
            Statement st = conn.createStatement();
            String sql = "DELETE FROM lophoc WHERE idlop='" + malop + "'";
            int reply = JOptionPane.showConfirmDialog(null, "Bạn có chắc chắn xóa không?", null, JOptionPane.YES_NO_OPTION);
            if (reply == JOptionPane.YES_OPTION) {
                st.executeUpdate(sql);
                JOptionPane.showMessageDialog(this, "Xóa thành công");
            } else {

            }
            load_lophoc();
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
            cell.setCellValue("DANH SÁCH LỚP HỌC");

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
            cell.setCellValue("Mã Lớp");

            cell = row.createCell(2, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Tên Lớp");

            cell = row.createCell(3, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Giảng viên");

            cell = row.createCell(4, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Số lượng sinh viên");

            cell = row.createCell(5, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Ngày bắt đầu");

            cell = row.createCell(6, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Ngày kết thúc");

            //Kết nối DB
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT \n"
                    + "    lophoc.idlop,\n"
                    + "    lophoc.tenlop,\n"
                    + "    giaovien.hoten AS tengiaovien, \n"
                    + "    lophoc.sosinhvien, \n"
                    + "    lophoc.ngaybatdau,\n"
                    + "    lophoc.ngayketthuc\n"
                    + "FROM \n"
                    + "    lophoc\n"
                    + "LEFT JOIN \n"
                    + "    giaovien ON lophoc.idgiaovien = giaovien.idgiaovien;";
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
                cell.setCellValue(rs.getString("idlop"));

                cell = row.createCell(2);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("tenlop"));

                //Định dạng ngày tháng trong excel
                cell = row.createCell(3);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("tengiaovien"));

                cell = row.createCell(4);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("sosinhvien"));

                java.util.Date ngaybd = new java.util.Date(rs.getDate("ngaybatdau").getTime());
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cell = row.createCell(5);
                cell.setCellValue(ngaybd);
                cell.setCellStyle(cellStyle);

                java.util.Date ngaykt = new java.util.Date(rs.getDate("ngayketthuc").getTime());
                cell = row.createCell(6);
                cell.setCellValue(ngaykt);
                cell.setCellStyle(cellStyle);

                i++;
            }
            //Hiệu chỉnh độ rộng của cột
            for (int col = 0; col < tongsocot; col++) {
                spreadsheet.autoSizeColumn(col);
            }

            File f = new File("C:\\TestJava\\danhsachlophoc.xlsx");
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
                    load_lophoc();
                } else {
                    JOptionPane.showMessageDialog(this, "Phải chọn file excel");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_btNhapExcelActionPerformed


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JComboBox<String> CboxGiangvien;
    private javax.swing.JButton btNhapExcel;
    private javax.swing.JButton btSua;
    private javax.swing.JButton btThem;
    private javax.swing.JButton btXoa;
    private javax.swing.JButton btXuatExcel;
    private javax.swing.JComboBox<String> jComboBox1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTable tblLophoc;
    private javax.swing.JTextField txtMalop;
    private com.toedter.calendar.JDateChooser txtNgaybatdau;
    private com.toedter.calendar.JDateChooser txtNgayketthuc;
    private javax.swing.JTextField txtSosv;
    private javax.swing.JTextField txtTenlop;
    // End of variables declaration//GEN-END:variables

    private void Themlophoc(String idlop, String tenlop, String tengiaovien, String sosinhvien, Date ngaybatdau, Date ngayketthuc) {
        String magv = giaovien.get(tengiaovien);
        
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "INSERT INTO lophoc (tenlop, idgiaovien, ngaybatdau, ngayketthuc) VALUES (?,?,?,?)";
            PreparedStatement st = conn.prepareStatement(sql);

            st.setString(1, tenlop);
            st.setString(2, magv);
            st.setDate(3, ngaybatdau);
            st.setDate(4, ngayketthuc);
            
            st.execute();

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
