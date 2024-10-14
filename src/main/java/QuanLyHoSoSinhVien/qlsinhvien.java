/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JPanel.java to edit this template
 */
package QuanLyHoSoSinhVien;

import ConnectDatabase.ConnectDB;
import java.awt.CardLayout;
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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
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
public class qlsinhvien extends javax.swing.JPanel {

    
    /**
     * Creates new form qlsinhvien
     */
    public qlsinhvien() {
        initComponents();
        load_sinhvien();
        load_Cboxsv();
        load_CboxLophoc();
        load_CboxKhoa();
//        load_CboxNganh();
    }

    Connection conn = null;

    private void load_sinhvien() {
        try {
            conn = ConnectDB.KetnoiDB();

            Statement statement = conn.createStatement();
            String query = "SELECT \n"
                    + "    sinhvien.masinhvien,\n"
                    + "    sinhvien.hoten,\n"
                    + "    sinhvien.ngaysinh,\n"
                    + "    sinhvien.gioitinh,\n"
                    + "    sinhvien.diachi,\n"
                    + "    sinhvien.sodienthoai,\n"
                    + "    sinhvien.email,\n"
                    + "    lophoc.tenlop,\n"
                    + "    khoa.tenkhoa,\n"
                    + "    nganhhoc.tennganh,\n"
                    + "    sinhvien.ngaynhaphoc\n"
                    + "FROM sinhvien\n"
                    + "JOIN lophoc ON sinhvien.idlop = lophoc.idlop\n"
                    + "JOIN khoa ON sinhvien.idkhoa = khoa.idkhoa\n"
                    + "JOIN nganhhoc ON sinhvien.idnganh = nganhhoc.idnganh;";
            ResultSet resultset = statement.executeQuery(query);

            tblSinhvien.removeAll();
            String[] tdb = {"Mã sinh viên", "Tên sinh viên", "Ngày sinh", "Giới tính", "Địa chỉ", "Điện thoại", "Email", "Lớp", "Khoa", "Ngành", "Ngày nhập học"};
            DefaultTableModel model = new DefaultTableModel(tdb, 0);

            int i = 0;
            while (resultset.next()) {

                Vector vector = new Vector();

                vector.add(resultset.getString("masinhvien"));
                vector.add(resultset.getString("hoten"));
                vector.add(resultset.getString("ngaysinh"));
                vector.add(resultset.getString("gioitinh"));
                vector.add(resultset.getString("diachi"));
                vector.add(resultset.getString("sodienthoai"));
                vector.add(resultset.getString("email"));
                vector.add(resultset.getString("tenlop"));
                vector.add(resultset.getString("tenkhoa"));
                vector.add(resultset.getString("tennganh"));
                vector.add(resultset.getString("ngaynhaphoc"));
                model.addRow(vector);
            }
            tblSinhvien.setModel(model);
            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        txtMasinhvien.setEnabled(false);
    }

    Map<String, String> lophoc = new HashMap<>();
    Map<String, String> khoa = new HashMap<>();
    Map<String, String> nganh = new HashMap<>();

    private void load_CboxLophoc() {
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT * FROM lophoc";
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(sql);

            while (rs.next()) {
                CboxLop.addItem(rs.getString("tenlop"));
                lophoc.put(rs.getString("tenlop"), rs.getString("idlop"));
            }

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void load_CboxKhoa() {
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT * FROM khoa";
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(sql);

            while (rs.next()) {
                CboxKhoa.addItem(rs.getString("tenkhoa"));
                khoa.put(rs.getString("tenkhoa"), rs.getString("idkhoa"));
            }

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void load_CboxNganh() {
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT * FROM nganhhoc";
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(sql);

            while (rs.next()) {
                CboxNganh.addItem(rs.getString("tennganh"));
                nganh.put(rs.getString("tennganh"), rs.getString("idnganh"));
            }

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    Map<String, String> sv = new HashMap<>();

    private void load_Cboxsv() {
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT * FROM sinhvien";
            Statement st = conn.createStatement();
            ResultSet rs = st.executeQuery(sql);

            while (rs.next()) {
                sv.put(rs.getString("masinhvien"), rs.getString("idsinhvien"));
            }

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
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
                double sdt;
                String hoten, gioitinh, diachi, sodienthoai, email, tenlop, tenkhoa, tennganh;
                Date ngaysinh, ngaynhaphoc;
//                Date ngs;
                hoten = row.getCell(0).getStringCellValue();
                ngaysinh = new Date(row.getCell(1).getDateCellValue().getTime());
                gioitinh = row.getCell(2).getStringCellValue();
                diachi = row.getCell(3).getStringCellValue();
                sodienthoai = row.getCell(4).getStringCellValue();
//                sodienthoai = String.valueOf((long) sdt);
                email = row.getCell(5).getStringCellValue();
                tenlop = row.getCell(6).getStringCellValue();
                tenkhoa = row.getCell(7).getStringCellValue();
                tennganh = row.getCell(8).getStringCellValue();
                ngaynhaphoc = new Date(row.getCell(9).getDateCellValue().getTime());

                Themsinhvien(hoten, ngaysinh, gioitinh, diachi, sodienthoai, email, tenlop, tenkhoa, tennganh, ngaynhaphoc);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
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
        CboxTimkiem = new javax.swing.JComboBox<>();
        jPanel7 = new javax.swing.JPanel();
        jPanel8 = new javax.swing.JPanel();
        txtTimkiem = new javax.swing.JTextField();
        jPanel9 = new javax.swing.JPanel();
        dateTimkiem1 = new com.toedter.calendar.JDateChooser();
        dateTimkiem2 = new com.toedter.calendar.JDateChooser();
        btTimkiem = new javax.swing.JButton();
        jPanel6 = new javax.swing.JPanel();
        jScrollPane2 = new javax.swing.JScrollPane();
        tblSinhvien = new javax.swing.JTable();
        jPanel1 = new javax.swing.JPanel();
        txtNgaysinh = new com.toedter.calendar.JDateChooser();
        jLabel1 = new javax.swing.JLabel();
        txtHoten = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        CboxGioitinh = new javax.swing.JComboBox<>();
        jLabel5 = new javax.swing.JLabel();
        txtDiachi = new javax.swing.JTextField();
        txtSodienthoai = new javax.swing.JTextField();
        jLabel7 = new javax.swing.JLabel();
        txtEmail = new javax.swing.JTextField();
        jLabel6 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        CboxKhoa = new javax.swing.JComboBox<>();
        jLabel8 = new javax.swing.JLabel();
        CboxLop = new javax.swing.JComboBox<>();
        jLabel11 = new javax.swing.JLabel();
        CboxNganh = new javax.swing.JComboBox<>();
        jLabel10 = new javax.swing.JLabel();
        txtNgaynhaphoc = new com.toedter.calendar.JDateChooser();
        jLabel12 = new javax.swing.JLabel();
        txtMasinhvien = new javax.swing.JTextField();

        jLabel2.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        jLabel2.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel2.setText("QUẢN LÝ SINH VIÊN");

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

        jPanel3.setBackground(new java.awt.Color(255, 255, 255));
        jPanel3.setBorder(javax.swing.BorderFactory.createTitledBorder("Thao tác"));

        btThem.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\Documents\\NetBeansProjects\\BTL_Nhom4\\src\\main\\resources\\image\\icons8_plus_+_48px_1-removebg-preview.png")); // NOI18N
        btThem.setText("Thêm");
        btThem.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        btThem.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        btThem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btThemActionPerformed(evt);
            }
        });

        btSua.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\Documents\\NetBeansProjects\\BTL_Nhom4\\src\\main\\resources\\image\\icons8_edit_property_48px-removebg-preview.png")); // NOI18N
        btSua.setText("Sửa");
        btSua.setHorizontalTextPosition(javax.swing.SwingConstants.CENTER);
        btSua.setVerticalTextPosition(javax.swing.SwingConstants.BOTTOM);
        btSua.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btSuaActionPerformed(evt);
            }
        });

        btXoa.setIcon(new javax.swing.ImageIcon("C:\\Users\\PC\\Documents\\NetBeansProjects\\BTL_Nhom4\\src\\main\\resources\\image\\icons8_trash_can_48px-removebg-preview.png")); // NOI18N
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

        CboxTimkiem.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Tên lớp", "Giảng viên", "Số sinh viên", "Thời gian học" }));
        CboxTimkiem.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                CboxTimkiemItemStateChanged(evt);
            }
        });
        CboxTimkiem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CboxTimkiemActionPerformed(evt);
            }
        });

        jPanel7.setLayout(new java.awt.CardLayout());

        javax.swing.GroupLayout jPanel8Layout = new javax.swing.GroupLayout(jPanel8);
        jPanel8.setLayout(jPanel8Layout);
        jPanel8Layout.setHorizontalGroup(
            jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel8Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(txtTimkiem, javax.swing.GroupLayout.DEFAULT_SIZE, 290, Short.MAX_VALUE)
                .addContainerGap())
        );
        jPanel8Layout.setVerticalGroup(
            jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel8Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(txtTimkiem, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(24, Short.MAX_VALUE))
        );

        jPanel7.add(jPanel8, "card2");

        javax.swing.GroupLayout jPanel9Layout = new javax.swing.GroupLayout(jPanel9);
        jPanel9.setLayout(jPanel9Layout);
        jPanel9Layout.setHorizontalGroup(
            jPanel9Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel9Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(dateTimkiem1, javax.swing.GroupLayout.PREFERRED_SIZE, 132, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(dateTimkiem2, javax.swing.GroupLayout.DEFAULT_SIZE, 140, Short.MAX_VALUE)
                .addContainerGap())
        );
        jPanel9Layout.setVerticalGroup(
            jPanel9Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel9Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel9Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(dateTimkiem1, javax.swing.GroupLayout.DEFAULT_SIZE, 30, Short.MAX_VALUE)
                    .addComponent(dateTimkiem2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(24, Short.MAX_VALUE))
        );

        jPanel7.add(jPanel9, "card3");

        btTimkiem.setText("Tìm kiếm");
        btTimkiem.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btTimkiemActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(CboxTimkiem, javax.swing.GroupLayout.PREFERRED_SIZE, 125, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(40, 40, 40)
                .addComponent(jPanel7, javax.swing.GroupLayout.PREFERRED_SIZE, 302, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(btTimkiem)
                .addGap(29, 29, 29))
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel7, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(CboxTimkiem, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btTimkiem))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        tblSinhvien.setModel(new javax.swing.table.DefaultTableModel(
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
        tblSinhvien.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                tblSinhvienMouseClicked(evt);
            }
        });
        jScrollPane2.setViewportView(tblSinhvien);

        javax.swing.GroupLayout jPanel6Layout = new javax.swing.GroupLayout(jPanel6);
        jPanel6.setLayout(jPanel6Layout);
        jPanel6Layout.setHorizontalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel6Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane2)
                .addContainerGap())
        );
        jPanel6Layout.setVerticalGroup(
            jPanel6Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel6Layout.createSequentialGroup()
                .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 294, Short.MAX_VALUE)
                .addContainerGap())
        );

        jPanel1.setBorder(javax.swing.BorderFactory.createTitledBorder("Nhập dữ liệu"));

        jLabel1.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel1.setText("Họ tên");

        jLabel3.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel3.setText("Ngày sinh");

        jLabel4.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel4.setText("Giới tính");

        CboxGioitinh.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "-- Chọn giới tính", "Nam", "Nữ" }));

        jLabel5.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel5.setText("Địa chỉ");

        txtDiachi.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtDiachiActionPerformed(evt);
            }
        });

        txtSodienthoai.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtSodienthoaiActionPerformed(evt);
            }
        });

        jLabel7.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel7.setText("Email");

        jLabel6.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel6.setText("Số điện thoại");

        jLabel9.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel9.setText("Khoa");

        CboxKhoa.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Khoa" }));
        CboxKhoa.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                CboxKhoaItemStateChanged(evt);
            }
        });

        jLabel8.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel8.setText("Lớp");

        CboxLop.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Lớp" }));

        jLabel11.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel11.setText("Ngành");

        CboxNganh.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Ngành" }));

        jLabel10.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel10.setText("Ngày nhập học");

        jLabel12.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel12.setText("Mã sinh viên");

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel12)
                    .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 61, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(CboxLop, javax.swing.GroupLayout.PREFERRED_SIZE, 137, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(jLabel9, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(CboxKhoa, javax.swing.GroupLayout.PREFERRED_SIZE, 181, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(40, 40, 40)
                        .addComponent(jLabel11, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(CboxNganh, javax.swing.GroupLayout.PREFERRED_SIZE, 171, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(36, 36, 36)
                        .addComponent(jLabel10, javax.swing.GroupLayout.PREFERRED_SIZE, 109, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addComponent(txtNgaynhaphoc, javax.swing.GroupLayout.PREFERRED_SIZE, 165, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addContainerGap())
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(txtDiachi, javax.swing.GroupLayout.PREFERRED_SIZE, 213, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(103, 103, 103)
                                .addComponent(jLabel6, javax.swing.GroupLayout.PREFERRED_SIZE, 92, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtSodienthoai, javax.swing.GroupLayout.PREFERRED_SIZE, 165, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(txtMasinhvien, javax.swing.GroupLayout.PREFERRED_SIZE, 179, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(30, 30, 30)
                                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 61, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(txtHoten, javax.swing.GroupLayout.PREFERRED_SIZE, 258, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel3)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtNgaysinh, javax.swing.GroupLayout.PREFERRED_SIZE, 152, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(39, 39, 39)
                                .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtEmail, javax.swing.GroupLayout.PREFERRED_SIZE, 181, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGap(18, 18, 18)
                        .addComponent(CboxGioitinh, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(14, 14, 14))))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                        .addComponent(CboxGioitinh, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jLabel4, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(txtHoten, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jLabel1)
                        .addComponent(jLabel12, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(txtMasinhvien, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(txtNgaysinh, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(23, 23, 23)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtDiachi, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtSodienthoai, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel6, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtEmail, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addGap(18, 18, 18)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel10, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(jLabel11, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(CboxNganh, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addComponent(txtNgaynhaphoc, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(CboxKhoa, javax.swing.GroupLayout.PREFERRED_SIZE, 34, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jLabel9, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 34, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(CboxLop, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap(19, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGap(12, 12, 12))
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
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
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 23, Short.MAX_VALUE)
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanel6, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
        );
    }// </editor-fold>//GEN-END:initComponents

    private void btThemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btThemActionPerformed
        // TODO add your handling code here:
//        TrangChu trangchu = new TrangChu();
//        ThemSVDialog tsvdialog = new ThemSVDialog(trangchu, true);
//        tsvdialog.setVisible(true);

        String tensv = txtHoten.getText().trim();
        Date ngs = new Date(txtNgaysinh.getDate().getTime());
        String gt = CboxGioitinh.getSelectedItem().toString();
        String dc = txtDiachi.getText().trim();
        String dt = txtSodienthoai.getText().trim();
        String email = txtEmail.getText().trim();
        String tenlop = CboxLop.getSelectedItem().toString();
        String malop = lophoc.get(tenlop);
        String tenkhoa = CboxKhoa.getSelectedItem().toString();
        String makhoa = khoa.get(tenkhoa);
        String tennganh = CboxNganh.getSelectedItem().toString();
        String manganh = nganh.get(tennganh);
        Date ngaynhaphoc = new Date(txtNgaynhaphoc.getDate().getTime());
//        if(!checkMatacgia()){
//            JOptionPane.showMessageDialog(this, "Mã tác giả đã tồn tại");
//            return;
//        }
        try {
            conn = ConnectDB.KetnoiDB();
//            String sql = "Insert Tacgia values('" + mtg + "',N'" + ttg + "','" + ngs + "',N'" + gt + "',"
//                    + "'" + dt + "','" + email + "',N'" + dc + "')";
            String sqli = "INSERT INTO sinhvien (hoten, ngaysinh, gioitinh, diachi, sodienthoai, email, idlop, idkhoa, idnganh, ngaynhaphoc) VALUES (?,?,?,?,?,?,?,?,?,?)";
            PreparedStatement st = conn.prepareStatement(sqli);

            st.setString(1, tensv);
            st.setDate(2, ngs);
            st.setString(3, gt);
            st.setString(4, dc);
            st.setString(5, dt);
            st.setString(6, email);
            st.setString(7, malop);
            st.setString(8, makhoa);
            st.setString(9, manganh);
            st.setDate(10, ngaynhaphoc);
            st.execute();

            conn.close();
            JOptionPane.showMessageDialog(this, "Thêm mới thành công");
            load_sinhvien();
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btThemActionPerformed

    private void btXoaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btXoaActionPerformed
        // TODO add your handling code here:
        int i = tblSinhvien.getSelectedRow();
        DefaultTableModel tb = (DefaultTableModel) tblSinhvien.getModel();
        String masv = tb.getValueAt(i, 0).toString();
        int idmasv = Integer.parseInt(sv.get(masv));
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "DELETE FROM sinhvien WHERE idsinhvien=?";
            PreparedStatement st = conn.prepareStatement(sql);
            st.setInt(1, idmasv);
            int reply = JOptionPane.showConfirmDialog(null, "Bạn có chắc chắn xóa không?", null, JOptionPane.YES_NO_OPTION);
            if (reply == JOptionPane.YES_OPTION) {
                st.executeUpdate();
                JOptionPane.showMessageDialog(this, "Xóa thành công");
                load_sinhvien();
            } else {

            }
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btXoaActionPerformed

    private void txtDiachiActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtDiachiActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtDiachiActionPerformed

    private void txtSodienthoaiActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtSodienthoaiActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtSodienthoaiActionPerformed

    private void btSuaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btSuaActionPerformed
        // TODO add your handling code here:
        String masv = txtMasinhvien.getText().trim();
        String tensv = txtHoten.getText().trim();
        Date ngs = new Date(txtNgaysinh.getDate().getTime());
        String gt = CboxGioitinh.getSelectedItem().toString();
        String dc = txtDiachi.getText().trim();
        String dt = txtSodienthoai.getText().trim();
        String email = txtEmail.getText().trim();
        String tenlop = CboxLop.getSelectedItem().toString();
        String malop = lophoc.get(tenlop);
        String tenkhoa = CboxKhoa.getSelectedItem().toString();
        String makhoa = khoa.get(tenkhoa);
        String tennganh = CboxNganh.getSelectedItem().toString();
        String manganh = nganh.get(tennganh);
        Date ngaynhaphoc = new Date(txtNgaynhaphoc.getDate().getTime());

        try {
            conn = ConnectDB.KetnoiDB();
//            String sql = "UPDATE tacgia SET tentacgia=N'" + ttg + "',ngaysinh='" + ngs + "',gioitinh=N'" + gt + "',dienthoai=" + "'" + dt + "',email='" + email + "',diachi=N'" + dc + "' WHERE matacgia='" + mtg + "'";
            String sqli = "UPDATE sinhvien SET hoten=?, ngaysinh=?, gioitinh=?, diachi=?, sodienthoai=?, email=?, idlop=?, idkhoa=?, idnganh=?, ngaynhaphoc=? WHERE masinhvien=?";
            PreparedStatement st = conn.prepareStatement(sqli);

            st.setString(1, tensv);
            st.setDate(2, ngs);
            st.setString(3, gt);
            st.setString(4, dc);
            st.setString(5, dt);
            st.setString(6, email);
            st.setString(7, malop);
            st.setString(8, makhoa);
            st.setString(9, manganh);
            st.setDate(10, ngaynhaphoc);
            st.setString(11, masv);
            st.execute();

            conn.close();
            JOptionPane.showMessageDialog(this, "Sửa thành công");
            load_sinhvien();
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }//GEN-LAST:event_btSuaActionPerformed

    private void tblSinhvienMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_tblSinhvienMouseClicked
        // TODO add your handling code here:
        int i = tblSinhvien.getSelectedRow();
        DefaultTableModel tb = (DefaultTableModel) tblSinhvien.getModel();
        txtMasinhvien.setText(tb.getValueAt(i, 0).toString());
        txtHoten.setText(tb.getValueAt(i, 1).toString());
        String ngay = tb.getValueAt(i, 2).toString();
        java.util.Date ngs;
        try {
            ngs = new SimpleDateFormat("yyyy-MM-dd").parse(ngay);
            txtNgaysinh.setDate(ngs);
        } catch (Exception e) {
            e.printStackTrace();
        }

        CboxGioitinh.setSelectedItem(tb.getValueAt(i, 3).toString());
        txtDiachi.setText(tb.getValueAt(i, 4).toString());
        txtSodienthoai.setText(tb.getValueAt(i, 5).toString());
        txtEmail.setText(tb.getValueAt(i, 6).toString());
        CboxLop.setSelectedItem(tb.getValueAt(i, 7).toString());
        CboxKhoa.setSelectedItem(tb.getValueAt(i, 8).toString());
        CboxNganh.setSelectedItem(tb.getValueAt(i, 9).toString());
        String ngaynhaphoc = tb.getValueAt(i, 10).toString();
        java.util.Date ngnhaphoc;
        try {
            ngnhaphoc = new SimpleDateFormat("yyyy-MM-dd").parse(ngaynhaphoc);
            txtNgaynhaphoc.setDate(ngnhaphoc);
        } catch (Exception e) {
            e.printStackTrace();
        }
        txtMasinhvien.setEnabled(false);
    }//GEN-LAST:event_tblSinhvienMouseClicked

    private void btXuatExcelActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btXuatExcelActionPerformed
        // TODO add your handling code here:
        try {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet spreadsheet = workbook.createSheet("sinhvien");
            // register the columns you wish to track and compute the column width

            CreationHelper createHelper = workbook.getCreationHelper();

            XSSFRow row = null;
            Cell cell = null;

            row = spreadsheet.createRow((short) 2);
            row.setHeight((short) 500);
            cell = row.createCell(0, CellType.STRING);
            cell.setCellValue("DANH SÁCH SINH VIÊN");

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
            cell.setCellValue("Mã Sinh Viên");

            cell = row.createCell(2, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Tên Sinh Viên");

            cell = row.createCell(3, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Ngày Sinh");

            cell = row.createCell(4, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Giới Tính");

            cell = row.createCell(5, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Địa Chỉ");

            cell = row.createCell(6, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Số Điện Thoại");

            cell = row.createCell(7, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Email");

            cell = row.createCell(8, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Lớp");

            cell = row.createCell(9, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Khoa");

            cell = row.createCell(10, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Ngành");

            cell = row.createCell(11, CellType.STRING);
            cell.setCellStyle(cellStyle_Head);
            cell.setCellValue("Ngày nhập học");

            //Kết nối DB
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT \n"
                    + "    sinhvien.masinhvien,\n"
                    + "    sinhvien.hoten,\n"
                    + "    sinhvien.ngaysinh,\n"
                    + "    sinhvien.gioitinh,\n"
                    + "    sinhvien.diachi,\n"
                    + "    sinhvien.sodienthoai,\n"
                    + "    sinhvien.email,\n"
                    + "    lophoc.tenlop,\n"
                    + "    khoa.tenkhoa,\n"
                    + "    nganhhoc.tennganh,\n"
                    + "    sinhvien.ngaynhaphoc\n"
                    + "FROM sinhvien\n"
                    + "JOIN lophoc ON sinhvien.idlop = lophoc.idlop\n"
                    + "JOIN khoa ON sinhvien.idkhoa = khoa.idkhoa\n"
                    + "JOIN nganhhoc ON sinhvien.idnganh = nganhhoc.idnganh;";
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
                cell.setCellValue(rs.getString("masinhvien"));

                cell = row.createCell(2);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("hoten"));

                //Định dạng ngày tháng trong excel
                java.util.Date ngay = new java.util.Date(rs.getDate("ngaysinh").getTime());
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cell = row.createCell(3);
                cell.setCellValue(ngay);
                cell.setCellStyle(cellStyle);

                cell = row.createCell(4);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("gioitinh"));

                cell = row.createCell(5);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("diachi"));

                cell = row.createCell(6);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("sodienthoai"));

                cell = row.createCell(7);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("email"));

                cell = row.createCell(8);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("tenlop"));

                cell = row.createCell(9);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("tenkhoa"));

                cell = row.createCell(10);
                cell.setCellStyle(cellStyle_data);
                cell.setCellValue(rs.getString("tennganh"));

                java.util.Date ngaynhaphoc = new java.util.Date(rs.getDate("ngaynhaphoc").getTime());
                cell = row.createCell(11);
                cell.setCellValue(ngaynhaphoc);
                cell.setCellStyle(cellStyle);

                i++;
            }
            //Hiệu chỉnh độ rộng của cột
            for (int col = 0; col < tongsocot; col++) {
                spreadsheet.autoSizeColumn(col);
            }

            File f = new File("C:\\TestJava\\danhsachsinhvien.xlsx");
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
                    load_sinhvien();
                } else {
                    JOptionPane.showMessageDialog(this, "Phải chọn file excel");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_btNhapExcelActionPerformed

    private void CboxKhoaItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_CboxKhoaItemStateChanged
        // TODO add your handling code here:
        String tenkhoa = CboxKhoa.getSelectedItem().toString();
        String makhoa = khoa.get(tenkhoa);
        CboxKhoa.setSelectedItem("");
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "SELECT * FROM nganhhoc WHERE idkhoa=?";
            PreparedStatement st = conn.prepareStatement(sql);
            st.setString(1, makhoa);
            ResultSet rs = st.executeQuery();

            int count = CboxNganh.getItemCount();
            for (int i = count - 1; i > 0; i--) {
                CboxNganh.removeItemAt(i);
            }

            while (rs.next()) {
                CboxNganh.addItem(rs.getString("tennganh"));
                nganh.put(rs.getString("tennganh"), rs.getString("idnganh"));
            }

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_CboxKhoaItemStateChanged
    CardLayout cardLayout;
    private void CboxTimkiemItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_CboxTimkiemItemStateChanged
        // TODO add your handling code here:
        txtTimkiem.setText("");
        String name = CboxTimkiem.getSelectedItem().toString();
        cardLayout = (CardLayout) jPanel6.getLayout();
        if (name.equals("Thời gian học")) {
            cardLayout.show(jPanel6, "card3");
        } else {
            cardLayout.show(jPanel6, "card2");
        }
    }//GEN-LAST:event_CboxTimkiemItemStateChanged

    private void CboxTimkiemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CboxTimkiemActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_CboxTimkiemActionPerformed

    private void btTimkiemActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btTimkiemActionPerformed
        // TODO add your handling code here:

        try {
            HashMap<String, String> searchData = new HashMap<>();
            String dkTimkiem = CboxTimkiem.getSelectedItem().toString();
            String tenlop = "";
            String gv = "";
            String idgv = "";
            String sosv = "";

            if (dkTimkiem.equals("Tên lớp")) {
                tenlop = txtTimkiem.getText().trim();
            } else if (dkTimkiem.equals("Giảng viên")) {
                gv = txtTimkiem.getText().trim();
//                idgv = giaovien.get(gv);
            } else if (dkTimkiem.equals("Số sinh viên")) {
                sosv = txtTimkiem.getText().trim();
            }

            if (!tenlop.isEmpty()) {
                searchData.put("tenlop", tenlop);
            } else if (!idgv.isEmpty()) {
                searchData.put("idgiaovien", idgv);
            } else if (!sosv.isEmpty()) {
                searchData.put("sosinhvien", sosv);
            }

            conn = ConnectDB.KetnoiDB();

            // Tạo câu truy vấn động
            StringBuilder sql = new StringBuilder("SELECT * FROM lophoc WHERE 1=1");

            // Duyệt qua HashMap và thêm các điều kiện vào câu truy vấn
            List<Object> parameters = new ArrayList<>();  // Danh sách các tham số sẽ được thêm vào PreparedStatement
            for (Map.Entry<String, String> entry : searchData.entrySet()) {
                String key = entry.getKey();
                String value = entry.getValue();

                if (key.equals("tenlop")) {
                    sql.append(" AND tenlop LIKE ?");
                    parameters.add("%" + value + "%");  // Thêm giá trị tương ứng vào danh sách tham số
                } else if (key.equals("idgiaovien")) {
                    sql.append(" AND idgiaovien = ?");
                    parameters.add(value);
                } else if (key.equals("sosinhvien")) {
                    sql.append(" AND sosinhvien = ?");
                    parameters.add(value);  // Không cần wildcard (%) với giá trị số
                }
            }

            // Chuẩn bị PreparedStatement
            PreparedStatement pstmt = conn.prepareStatement(sql.toString());

            // Đặt các tham số vào PreparedStatement
            for (int i = 0; i < parameters.size(); i++) {
                pstmt.setObject(i + 1, parameters.get(i));  // Set các tham số, bắt đầu từ 1
            }

            // Thực thi câu truy vấn
            ResultSet resultset = pstmt.executeQuery();

            tblSinhvien.removeAll();
            String[] tdb = {"Mã lớp", "Tên lớp học", "Giảng viên", "Số lượng sinh viên", "Ngày bắt đầu", "Ngày kết thúc"};
            DefaultTableModel model = new DefaultTableModel(tdb, 0);

            int i = 0;
            while (resultset.next()) {

                Vector vector = new Vector();

                vector.add(resultset.getString("idlop"));
                vector.add(resultset.getString("tenlop"));
//                String id = magv.get(resultset.getString("idgiaovien"));
//                vector.add(id);
                vector.add(resultset.getString("sosinhvien"));
                vector.add(resultset.getString("ngaybatdau"));
                vector.add(resultset.getString("ngayketthuc"));
                model.addRow(vector);
            }
            tblSinhvien.setModel(model);
            conn.close();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_btTimkiemActionPerformed


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JComboBox<String> CboxGioitinh;
    private javax.swing.JComboBox<String> CboxKhoa;
    private javax.swing.JComboBox<String> CboxLop;
    private javax.swing.JComboBox<String> CboxNganh;
    private javax.swing.JComboBox<String> CboxTimkiem;
    private javax.swing.JButton btNhapExcel;
    private javax.swing.JButton btSua;
    private javax.swing.JButton btThem;
    private javax.swing.JButton btTimkiem;
    private javax.swing.JButton btXoa;
    private javax.swing.JButton btXuatExcel;
    private com.toedter.calendar.JDateChooser dateTimkiem1;
    private com.toedter.calendar.JDateChooser dateTimkiem2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel12;
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
    private javax.swing.JPanel jPanel6;
    private javax.swing.JPanel jPanel7;
    private javax.swing.JPanel jPanel8;
    private javax.swing.JPanel jPanel9;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JTable tblSinhvien;
    private javax.swing.JTextField txtDiachi;
    private javax.swing.JTextField txtEmail;
    private javax.swing.JTextField txtHoten;
    private javax.swing.JTextField txtMasinhvien;
    private com.toedter.calendar.JDateChooser txtNgaynhaphoc;
    private com.toedter.calendar.JDateChooser txtNgaysinh;
    private javax.swing.JTextField txtSodienthoai;
    private javax.swing.JTextField txtTimkiem;
    // End of variables declaration//GEN-END:variables

    private void Themsinhvien(String hoten, Date ngaysinh, String gioitinh, String diachi, String sodienthoai, String email, String tenlop, String tenkhoa, String tennganh, Date ngaynhaphoc) {
        String malop = lophoc.get(tenlop);
        String makhoa = khoa.get(tenkhoa);
        String manganh = nganh.get(tennganh);
        try {
            conn = ConnectDB.KetnoiDB();
            String sql = "INSERT INTO sinhvien (hoten, ngaysinh, gioitinh, diachi, sodienthoai, email, idlop, idkhoa, idnganh, ngaynhaphoc) VALUES (?,?,?,?,?,?,?,?,?,?)";
            PreparedStatement st = conn.prepareStatement(sql);

            st.setString(1, hoten);
            st.setDate(2, ngaysinh);
            st.setString(3, gioitinh);
            st.setString(4, diachi);
            st.setString(5, sodienthoai);
            st.setString(6, email);
            st.setString(7, malop);
            st.setString(8, makhoa);
            st.setString(9, manganh);
            st.setDate(10, ngaynhaphoc);
            st.execute();

            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
