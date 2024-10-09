/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JPanel.java to edit this template
 */
package QuanLyHoSoSinhVien;

import ConnectDatabase.ConnectDB;
import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
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
import javax.swing.JButton;
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
public class qldangkihoc extends javax.swing.JPanel {

    private int currentPage = 1; // Trang hiện tại
    private int rowsPerPage = 5; // Số bản ghi mỗi trang
    private int totalRows; // Tổng số bản ghi
    private int totalPage; // Tổng số trang

//    private qlhocphi qlhocphi;
    /**
     * Creates new form qldangkihoc
     */
    public qldangkihoc() {
        initComponents();
        load_dkhoc(currentPage);
        load_monhoc();
        load_sv();
//        qlhocphi = new qlhocphi();
    }

//    private void reload_hocphi() {
//        qlhocphi.load_hocphi();
//        qlhocphi.getTblHocphi().revalidate();
//        qlhocphi.getTblHocphi().repaint();
//    }
    Connection conn = null;
    Map<String, String> monhoc = new HashMap<>();

    private void load_monhoc() {
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT * FROM monhoc";
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(sql);

            while (rs.next()) {
                CboxMonhoc.addItem(rs.getString("tenmonhoc"));
                monhoc.put(rs.getString("tenmonhoc"), rs.getString("idmonhoc"));
            }

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

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

    private void load_dkhoc(int currentPage) {
        try {
            conn = ConnectDB.KetnoiDB();
            Statement statement = conn.createStatement();

            // Tính toán OFFSET dựa trên trang hiện tại
            int offset = (currentPage - 1) * rowsPerPage;

            // Đếm tổng số bản ghi để tính tổng số trang
            String countQuery = "SELECT COUNT(*) FROM dangkyhoc";
            ResultSet countResult = statement.executeQuery(countQuery);
            if (countResult.next()) {
                totalRows = countResult.getInt(1);
            }
            totalPage = (int) Math.ceil((double) totalRows / rowsPerPage);

            // Câu truy vấn với LIMIT và OFFSET cho phân trang
            String query = "SELECT \n"
                    + "    dangkyhoc.iddangky,\n"
                    + "    sinhvien.masinhvien,\n"
                    + "    sinhvien.hoten AS tensinhvien,\n"
                    + "    monhoc.tenmonhoc AS tenmonhoc,\n"
                    + "    dangkyhoc.ngaydangky\n"
                    + "FROM \n"
                    + "    dangkyhoc\n"
                    + "JOIN \n"
                    + "    sinhvien ON dangkyhoc.idsinhvien = sinhvien.idsinhvien\n"
                    + "JOIN \n"
                    + "    monhoc ON dangkyhoc.idmonhoc = monhoc.idmonhoc\n"
                    + "LIMIT " + rowsPerPage + " OFFSET " + offset;

            ResultSet resultset = statement.executeQuery(query);

            tblDKhoc.removeAll();
            String[] tdb = {"Mã đăng ký", "Mã sinh viên", "Tên sinh viên", "Môn học đã đăng ký", "Ngày đăng ký"};
            DefaultTableModel model = new DefaultTableModel(tdb, 0);

            while (resultset.next()) {
                Vector vector = new Vector();
                vector.add(resultset.getString("iddangky"));
                vector.add(resultset.getString("masinhvien"));
                vector.add(resultset.getString("tensinhvien"));
                vector.add(resultset.getString("tenmonhoc"));
                vector.add(resultset.getString("ngaydangky"));
                model.addRow(vector);
            }

            tblDKhoc.setModel(model);
            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        txtMadangky.setEnabled(false);
    }

    private void setupPaginationButtons() {
        // Nếu nút đã tồn tại, không cần xóa hoặc tạo lại từ đầu
//    PanelPageButton.removeAll();  // Xóa tất cả các nút cũ trước khi thêm mới

    for (int i = 1; i <= totalPage; i++) {
        JButton pageButton = new JButton(String.valueOf(i));
        final int page = i;

        // Đặt nút của trang hiện tại có màu khác biệt để phân biệt
        if (i == currentPage) {
            pageButton.setBackground(Color.GRAY);  // Màu nổi bật cho trang hiện tại
        } else {
            pageButton.setBackground(null);  // Màu bình thường cho các trang khác
        }

        pageButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                currentPage = page;
                load_dkhoc(currentPage);  // Tải dữ liệu của trang hiện tại
                setupPaginationButtons();  // Cập nhật lại các nút trang
            }
        });

        PanelPageButton.add(pageButton);  // Thêm nút vào JPanel
    }

    // Kiểm tra nếu nút "Previous" có cần bật hay không
    PanelPageButton.setEnabled(currentPage > 1);

    // Kiểm tra nếu nút "Next" có cần bật hay không
    btnNext.setEnabled(currentPage < totalPage);

    PanelPageButton.revalidate();  // Cập nhật lại giao diện
    PanelPageButton.repaint();  // Vẽ lại JPanel
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
                double iddk, hpdong;
                String masinhvien, tensinhvien, tenmonhoc;
                Date ngaydangky;
//                Date ngs;

                iddk = row.getCell(0).getNumericCellValue();
                String iddangkyhoc = String.valueOf((int) iddk);
                masinhvien = row.getCell(1).getStringCellValue();
                tensinhvien = row.getCell(2).getStringCellValue();
                tenmonhoc = row.getCell(3).getStringCellValue();
                ngaydangky = new Date(row.getCell(4).getDateCellValue().getTime());

                Dangkymonhoc(iddangkyhoc, masinhvien, tensinhvien, tenmonhoc, ngaydangky);
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
        jLabel1 = new javax.swing.JLabel();
        jPanel3 = new javax.swing.JPanel();
        btThem = new javax.swing.JButton();
        btSua = new javax.swing.JButton();
        btXoa = new javax.swing.JButton();
        btXuatExcel = new javax.swing.JButton();
        btNhapExcel = new javax.swing.JButton();
        jPanel4 = new javax.swing.JPanel();
        jComboBox1 = new javax.swing.JComboBox<>();
        jTextField1 = new javax.swing.JTextField();
        jPanel6 = new javax.swing.JPanel();
        txtNgdangky = new com.toedter.calendar.JDateChooser();
        jLabel2 = new javax.swing.JLabel();
        txtTenSV = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jLabel12 = new javax.swing.JLabel();
        txtMaSV = new javax.swing.JTextField();
        jLabel13 = new javax.swing.JLabel();
        txtMadangky = new javax.swing.JTextField();
        CboxMonhoc = new javax.swing.JComboBox<>();
        jPanel5 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        tblDKhoc = new javax.swing.JTable();
        PanelPageButton = new javax.swing.JPanel();
        btnPrevious = new javax.swing.JButton();
        btnNext = new javax.swing.JButton();

        jLabel1.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("Đăng ký học");

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(426, 426, 426)
                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 243, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, 43, Short.MAX_VALUE)
                .addContainerGap())
        );

        jPanel3.setBorder(javax.swing.BorderFactory.createTitledBorder("Thao tác"));

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

        btXuatExcel.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\Documents\\NetBeansProjects\\BTL_Nhom4\\src\\main\\resources\\image\\Fatcow-Farm-Fresh-Excel-exports.32.png")); // NOI18N
        btXuatExcel.setText("Xuất excel");
        btXuatExcel.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        btXuatExcel.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        btXuatExcel.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        btXuatExcel.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btXuatExcelActionPerformed(evt);
            }
        });

        btNhapExcel.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\Documents\\NetBeansProjects\\BTL_Nhom4\\src\\main\\resources\\image\\Fatcow-Farm-Fresh-Excel-imports.32.png")); // NOI18N
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
                .addComponent(btSua, javax.swing.GroupLayout.PREFERRED_SIZE, 72, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(btXoa)
                .addGap(18, 18, 18)
                .addComponent(btXuatExcel)
                .addGap(18, 18, 18)
                .addComponent(btNhapExcel, javax.swing.GroupLayout.PREFERRED_SIZE, 89, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(btNhapExcel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(btXuatExcel, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, jPanel3Layout.createSequentialGroup()
                        .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(btSua)
                            .addComponent(btThem)
                            .addComponent(btXoa))
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );

        jPanel4.setBorder(javax.swing.BorderFactory.createTitledBorder("Tìm kiếm"));

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Item 1", "Item 2", "Item 3", "Item 4" }));

        jTextField1.setText("jTextField1");

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addGap(27, 27, 27)
                .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(114, 114, 114)
                .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(372, Short.MAX_VALUE))
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(16, Short.MAX_VALUE))
        );

        jPanel6.setBorder(javax.swing.BorderFactory.createTitledBorder("Nhập dữ liệu"));

        jLabel2.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel2.setText("Tên sinh viên");

        jLabel3.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel3.setText("Ngày đăng ký");

        jLabel7.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel7.setText("Môn học");

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
        jLabel13.setText("Mã đăng ký");
        jLabel13.setToolTipText("");

        CboxMonhoc.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Môn học" }));

        javax.swing.GroupLayout jPanel6Layout = new javax.swing.GroupLayout(jPanel6);
        jPanel6.setLayout(jPanel6Layout);
        jPanel6Layout.setHorizontalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel6Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel13, javax.swing.GroupLayout.PREFERRED_SIZE, 91, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel7))
                .addGap(36, 36, 36)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(CboxMonhoc, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(txtMadangky, javax.swing.GroupLayout.DEFAULT_SIZE, 179, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 128, Short.MAX_VALUE)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel12)
                    .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 121, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(1, 1, 1)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(txtNgdangky, javax.swing.GroupLayout.PREFERRED_SIZE, 179, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtMaSV, javax.swing.GroupLayout.PREFERRED_SIZE, 179, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(61, 61, 61)
                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 86, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(txtTenSV, javax.swing.GroupLayout.PREFERRED_SIZE, 190, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(69, 69, 69))
        );
        jPanel6Layout.setVerticalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel6Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtTenSV, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 34, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel12, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtMaSV, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel13, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtMadangky, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(20, 20, 20)
                .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(jLabel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(txtNgdangky, javax.swing.GroupLayout.DEFAULT_SIZE, 34, Short.MAX_VALUE)
                    .addComponent(CboxMonhoc))
                .addContainerGap(16, Short.MAX_VALUE))
        );

        tblDKhoc.setModel(new javax.swing.table.DefaultTableModel(
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
        tblDKhoc.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                tblDKhocMouseClicked(evt);
            }
        });
        jScrollPane1.setViewportView(tblDKhoc);

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel5Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane1)
                .addContainerGap())
        );
        jPanel5Layout.setVerticalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 357, Short.MAX_VALUE)
        );

        btnPrevious.setText("<<");
        btnPrevious.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnPreviousActionPerformed(evt);
            }
        });

        btnNext.setText(">>");
        btnNext.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnNextActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout PanelPageButtonLayout = new javax.swing.GroupLayout(PanelPageButton);
        PanelPageButton.setLayout(PanelPageButtonLayout);
        PanelPageButtonLayout.setHorizontalGroup(
            PanelPageButtonLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(PanelPageButtonLayout.createSequentialGroup()
                .addContainerGap()
                .addComponent(btnPrevious, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(btnNext, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );
        PanelPageButtonLayout.setVerticalGroup(
            PanelPageButtonLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(PanelPageButtonLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                .addComponent(btnPrevious)
                .addComponent(btnNext))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel2, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGap(18, 18, 18))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jPanel6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addContainerGap())))
            .addComponent(jPanel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(PanelPageButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jPanel4, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addComponent(jPanel6, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(PanelPageButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );
    }// </editor-fold>//GEN-END:initComponents

    private void btThemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btThemActionPerformed
        // TODO add your handling code here:
        String msv = txtMaSV.getText().trim();
        String idsv = masv.get(msv);
        String tensv = txtTenSV.getText().trim();
        String mh = CboxMonhoc.getSelectedItem().toString();
        int mamh = Integer.parseInt(monhoc.get(mh));
        Date ngdangky = new Date(txtNgdangky.getDate().getTime());

        //        if (!checkTenlop()) {
        //            JOptionPane.showMessageDialog(this, "Lớp học đã tồn tại");
        //            return;
        //        }
        try {
            conn = ConnectDB.KetnoiDB();
            //            String sql = "Insert Tacgia values('" + mtg + "',N'" + ttg + "','" + ngs + "',N'" + gt + "',"
            //                    + "'" + dt + "','" + email + "',N'" + dc + "')";
            String sqli = "INSERT INTO dangkyhoc (idsinhvien, idmonhoc, ngaydangky) VALUES (?,?,?)";
            PreparedStatement st = conn.prepareStatement(sqli);

            st.setString(1, idsv);
            st.setInt(2, mamh);
            st.setDate(3, ngdangky);

            st.execute();

            conn.close();
            JOptionPane.showMessageDialog(this, "Thêm mới thành công");
            load_dkhoc(currentPage);
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btThemActionPerformed

    private void btSuaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btSuaActionPerformed
        // TODO add your handling code here:
        String madk = txtMadangky.getText().trim();
        String msv = txtMaSV.getText().trim();
        String idsv = masv.get(msv);
        String tensv = txtTenSV.getText().trim();
        String mh = CboxMonhoc.getSelectedItem().toString();
        int mamh = Integer.parseInt(monhoc.get(mh));
        Date ngdangky = new Date(txtNgdangky.getDate().getTime());

        try {
            conn = ConnectDB.KetnoiDB();
            //            String sql = "UPDATE tacgia SET tentacgia=N'" + ttg + "',ngaysinh='" + ngs + "',gioitinh=N'" + gt + "',dienthoai=" + "'" + dt + "',email='" + email + "',diachi=N'" + dc + "' WHERE matacgia='" + mtg + "'";
            String sqli = "UPDATE dangkyhoc SET idsinhvien=?, idmonhoc=?, ngaydangky=? WHERE iddangky=?";
            PreparedStatement st = conn.prepareStatement(sqli);

            st.setString(1, idsv);
            st.setInt(2, mamh);
            st.setDate(3, ngdangky);
            st.setString(4, madk);

            st.execute();

            conn.close();
            JOptionPane.showMessageDialog(this, "Sửa thành công");
            load_dkhoc(currentPage);
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btSuaActionPerformed

    private void btXoaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btXoaActionPerformed
        // TODO add your handling code here:
        String iddk = txtMadangky.getText().trim();

        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "DELETE FROM dangkyhoc WHERE iddangky=?";
            PreparedStatement st = conn.prepareStatement(sql);
            st.setString(1, iddk);
            int reply = JOptionPane.showConfirmDialog(null, "Bạn có chắc chắn xóa không?", null, JOptionPane.YES_NO_OPTION);
            if (reply == JOptionPane.YES_OPTION) {
                st.executeUpdate();
                JOptionPane.showMessageDialog(this, "Xóa thành công");
            } else {

            }
            load_dkhoc(currentPage);
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btXoaActionPerformed

    private void btXuatExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btXuatExcelActionPerformed
        // TODO add your handling code here:
        try {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet spreadsheet = workbook.createSheet("dangkyhoc");
            // register the columns you wish to track and compute the column width

            CreationHelper createHelper = workbook.getCreationHelper();

            XSSFRow row = null;
            Cell cell = null;

            row = spreadsheet.createRow((short) 2);
            row.setHeight((short) 500);
            cell = row.createCell(0, CellType.STRING);
            cell.setCellValue("DANH SÁCH SINH VIÊN ĐĂNG KÝ MÔN HỌC");

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
            cell.setCellValue("Mã đăng ký");

            cell = row.createCell(2, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Mã sinh viên");

            cell = row.createCell(3, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Tên sinh viên");

            cell = row.createCell(4, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Tên môn học");

            cell = row.createCell(5, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Ngày đăng ký");

            //Kết nối DB
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT \n"
                    + "    dangkyhoc.iddangky,\n"
                    + "    sinhvien.masinhvien,\n"
                    + "    sinhvien.hoten AS tensinhvien,\n"
                    + "    monhoc.tenmonhoc AS tenmonhoc,\n"
                    + "    dangkyhoc.ngaydangky\n"
                    + "FROM \n"
                    + "    dangkyhoc\n"
                    + "JOIN \n"
                    + "    sinhvien ON dangkyhoc.idsinhvien = sinhvien.idsinhvien\n"
                    + "JOIN \n"
                    + "    monhoc ON dangkyhoc.idmonhoc = monhoc.idmonhoc";
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
                cell.setCellValue(rs.getString("iddangky"));

                cell = row.createCell(2);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("masinhvien"));

                cell = row.createCell(3);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("tensinhvien"));

                cell = row.createCell(4);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("tenmonhoc"));

                //Định dạng ngày tháng trong excel
                java.util.Date hanchot = new java.util.Date(rs.getDate("ngaydangky").getTime());
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cell = row.createCell(5);
                cell.setCellValue(hanchot);
                cell.setCellStyle(cellStyle);

                i++;
            }
            //Hiệu chỉnh độ rộng của cột
            for (int col = 0; col < tongsocot; col++) {
                spreadsheet.autoSizeColumn(col);
            }

            File f = new File("C:\\TestJava\\danhsachdangkyhoc.xlsx");
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
                    load_dkhoc(currentPage);
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

    private void txtMaSVInputMethodTextChanged(java.awt.event.InputMethodEvent evt) {//GEN-FIRST:event_txtMaSVInputMethodTextChanged
        // TODO add your handling code here:
    }//GEN-LAST:event_txtMaSVInputMethodTextChanged

    private void txtMaSVActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtMaSVActionPerformed
        // TODO add your handling code here:
        //        String msv = evt.getActionCommand();
        String msv = txtMaSV.getText().trim();
        txtTenSV.setText(tensv.get(msv));
    }//GEN-LAST:event_txtMaSVActionPerformed

    private void tblDKhocMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_tblDKhocMouseClicked
        // TODO add your handling code here:
        int i = tblDKhoc.getSelectedRow();
        DefaultTableModel tb = (DefaultTableModel) tblDKhoc.getModel();
        txtMadangky.setText(tb.getValueAt(i, 0).toString());
        txtMaSV.setText(tb.getValueAt(i, 1).toString());
        txtTenSV.setText(tb.getValueAt(i, 2).toString());
        CboxMonhoc.setSelectedItem(tb.getValueAt(i, 3).toString());
        String ngaydk = tb.getValueAt(i, 4).toString();
        java.util.Date ngdk;
        try {
            ngdk = new SimpleDateFormat("yyyy-MM-dd").parse(ngaydk);
            txtNgdangky.setDate(ngdk);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_tblDKhocMouseClicked

    private void btnPreviousActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnPreviousActionPerformed
        // TODO add your handling code here:
        if (currentPage > 1) {
            currentPage--;
            load_dkhoc(currentPage);
            setupPaginationButtons(); // Cập nhật lại các nút trang
        }
    }//GEN-LAST:event_btnPreviousActionPerformed

    private void btnNextActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnNextActionPerformed
        // TODO add your handling code here:
        if (currentPage < totalPage) {
            currentPage++;
            load_dkhoc(currentPage);
            setupPaginationButtons(); // Cập nhật lại các nút trang
        }
    }//GEN-LAST:event_btnNextActionPerformed


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JComboBox<String> CboxMonhoc;
    private javax.swing.JPanel PanelPageButton;
    private javax.swing.JButton btNhapExcel;
    private javax.swing.JButton btSua;
    private javax.swing.JButton btThem;
    private javax.swing.JButton btXoa;
    private javax.swing.JButton btXuatExcel;
    private javax.swing.JButton btnNext;
    private javax.swing.JButton btnPrevious;
    private javax.swing.JComboBox<String> jComboBox1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JPanel jPanel6;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTable tblDKhoc;
    private javax.swing.JTextField txtMaSV;
    private javax.swing.JTextField txtMadangky;
    private com.toedter.calendar.JDateChooser txtNgdangky;
    private javax.swing.JTextField txtTenSV;
    // End of variables declaration//GEN-END:variables

    private void Dangkymonhoc(String iddangkyhoc, String masinhvien, String tensinhvien, String tenmonhoc, Date ngaydangky) {
        String msv = masv.get(masinhvien);
        String mh = monhoc.get(tenmonhoc);

        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "INSERT INTO dangkyhoc (idsinhvien, idmonhoc, ngaydangky) VALUES (?,?,?)";
            PreparedStatement st = conn.prepareStatement(sql);

            st.setString(1, msv);
            st.setString(2, mh);
            st.setDate(3, ngaydangky);

            st.execute();

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
