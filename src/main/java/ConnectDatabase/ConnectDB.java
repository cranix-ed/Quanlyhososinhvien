/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package ConnectDatabase;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

/**
 *
 * @author Cranux
 */
public class ConnectDB {
    public static Connection KetnoiDB() {
        Connection con = null;
            try {
                String url="jdbc:mysql://localhost:3306/quanlyhososinhvien?zeroDateTimeBehavior=CONVERT_TO_NULL";
                String user="root";
                String pass="";
                con = DriverManager.getConnection(url, user, pass);
            } catch (SQLException ex) {
                ex.printStackTrace();
            }
            return con;
    }
}
