/**
 * Created by yilun on 2017/5/23.
 */

import javafx.embed.swing.JFXPanel;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import javax.swing.*;
import javax.swing.filechooser.FileFilter;

public class WriteRemark2MysqlAutomatically extends JFrame {
    String absolutePath = "";

    public static void main(String args[]) {
        try {
            JFrame.setDefaultLookAndFeelDecorated(true);
            UIManager.setLookAndFeel("ch.randelshofer.quaqua.QuaquaLookAndFeel");
            new WriteRemark2MysqlAutomatically();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    public WriteRemark2MysqlAutomatically() {
        final JFrame jf=this;
        final JPanel jp = new JPanel(null);
        jp.setBounds(0,0,800,600);
        Container contentPane = getContentPane();
        JLabel labelIP = new JLabel("数据库IP:");
        labelIP.setBounds(300, 100, 100, 30);
        final JTextField jtfIP = new JTextField("10.1.62.241");
        jtfIP.setBounds(400, 100, 200, 30);

        JLabel labelDatabaseName = new JLabel("数据库名称:");
        labelDatabaseName.setBounds(300, 200, 100, 30);
        final JTextField jtfDatabaseName = new JTextField("sk");
        jtfDatabaseName.setBounds(400, 200, 200, 30);

        JLabel labelUserName = new JLabel("数据库用户名:");
        labelUserName.setBounds(300, 300, 100, 30);
        final JTextField jtfUserName = new JTextField("root");
        jtfUserName.setBounds(400, 300, 200, 30);

        JLabel labelPassword = new JLabel("数据库密码:");
        labelPassword.setBounds(300, 400, 100, 30);
        final JTextField jtfPassword = new JTextField("root");
        jtfPassword.setBounds(400, 400, 200, 30);

        final JLabel importing = new JLabel("导入中...,请稍后");
        importing.setFont(new   java.awt.Font("Dialog",   1,   25));
        importing.setForeground(Color.RED);
        importing.setBounds(400, 230, 200, 70);
        importing.setVisible(false);

        JButton open = new JButton("点击选择要导入的excel文件");
        open.setBounds(300, 500, 100, 30);

        JButton submit = new JButton("确定");
        submit.setBounds(500, 500, 100, 30);

        this.setLayout(null);

        //设置尺寸
        jp.add(labelIP);
        jp.add(jtfIP);
        jp.add(labelDatabaseName);
        jp.add(jtfDatabaseName);
        jp.add(labelUserName);
        jp.add(jtfUserName);
        jp.add(labelPassword);
        jp.add(jtfPassword);
        jp.add(importing);
        jp.add(open);
        jp.add(submit);
        contentPane.add(jp);
        this.setBounds(0, 0, 800, 600);
        this.setVisible(true);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        open.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc = new JFileChooser();
                jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
                jfc.setFileFilter(new FileFilter() {
                    @Override
                    public boolean accept(File file) {
                        String name = file.getName();
                        return file.isDirectory() || name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".xlsx");  // 仅显示目录和xls、xlsx文件
                    }

                    @Override
                    public String getDescription() {
                        return "*.xls;*.xlsx";
                    }
                });
                jfc.showDialog(new JLabel(), "选择");
                absolutePath = jfc.getSelectedFile().getAbsolutePath();
            }
        });
        //确定按钮
        submit.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                Connection mysqlconn = null;
                try {

                    //先尝试下连接数据库 报错就不执行
                    String IP = jtfIP.getText();
                    String databaseName = jtfDatabaseName.getText();
                    String userName = jtfUserName.getText();
                    String password = jtfPassword.getText();
                    final String mysqlDriver = "com.mysql.jdbc.Driver";
                    String mysqlurl = "jdbc:mysql://" + IP + ":3306/" + databaseName + "?useUnicode=true&characterEncoding=UTF-8";
                    String mysqlName = userName;
                    String mysqlPassword = password;
                    Class.forName(mysqlDriver);
                    mysqlconn = DriverManager.getConnection(mysqlurl, mysqlName, mysqlPassword);

                    //判断是不是已经选择了文件夹
                    if (absolutePath.isEmpty()) {
                        JOptionPane.showMessageDialog(jp, "请先选择Excel文件", "错误", JOptionPane.WARNING_MESSAGE);
                    } else {
                        WriteRemark wr = new WriteRemark();
                        wr.setMysqlconn(mysqlconn);
                        wr.setExcelPath(absolutePath);
                        wr.setMysqlName(userName);
                        wr.setMysqlurl(mysqlurl);
                        //检查Excel的框架是不是正确
                        Message checkMessage = wr.checkAndSaveRightExcel();
                        if (checkMessage.isSuccess()) {
                            Message writeMessage = wr.write();
                            //判断是不是插入成功
                            if (writeMessage.isSuccess()) {
                                JOptionPane.showMessageDialog(jp, writeMessage.getMessageContent(), "成功", JOptionPane.PLAIN_MESSAGE);
                            }else{
                                JOptionPane.showMessageDialog(jp, writeMessage.getMessageContent(), "失败", JOptionPane.WARNING_MESSAGE);
                            }
                        } else {
                            JOptionPane.showMessageDialog(jp, checkMessage.getMessageContent(),"失败" , JOptionPane.WARNING_MESSAGE);
                        }
                    }
                } catch (SQLException e1) {
                    JOptionPane.showMessageDialog(jp, "数据库连接失败,检查前四项是否正确", "错误", JOptionPane.WARNING_MESSAGE);
                    try {
                        mysqlconn.close();
                    } catch (Exception e3) {
                        e3.printStackTrace();
                    }
                } catch (Exception e2) {
                    e2.printStackTrace();
                }
            }
        });

    }

}
