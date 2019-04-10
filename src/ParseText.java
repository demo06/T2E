import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 解析文本
 * 2019年4月8日
 */
public class ParseText implements CallBack {
    String data1[] = new String[11];
    String data2[] = new String[11];
    JFrame jFrame;
    JPanel jPanel;
    JButton jBtnInsert, jBtnClear, jBtnAllBook, jBtnDivisive;
    JTextArea jTextArea;
    JLabel jLabel, jLabel2;
    JProgressBar jProgressBar;
    JComboBox jComboBox;
    String[] title = {"宽带账号", "地址", "客户姓名", "客户电话", "收费", "员工姓名", "成交日期"};
    String[] title2 = {"办理日期", "姓名", "办卡人户名", "宽带账号", "办理号码", "收费金额", "备注", "邮箱"};
    String[] areaTitle = {"宽带账号", "地址", "客户姓名", "客户电话"};
    String[] keyWord = {"宽带账号：", "地址：", "电话：", "姓名：", "备注：", "收费：",};
    String[] areas = {"郑东新区", "金水区", "经开区", "管城区", "中原区", "二七区", "惠济区", "高新区"};

    private String workName, coustomerName, coustomerMobile, account, coustomerAddress, newNumber, status, price, date,
            str, fullText, eMail, lastStr;
    boolean isExist;

    private ExcelUtil excelUtil;

    public ParseText(ExcelUtil excelUtil) {
        this.excelUtil = excelUtil;
    }

    public void initView() {
        jFrame = new JFrame("文档分析转换"); // 创建窗口
        jFrame.setSize(440, 370); // 设置窗口大小
        jFrame.setLocationRelativeTo(null); // 把窗口位置设置到屏幕中心
        jFrame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE); // 当点击窗口的关闭按钮时退出程序（没有这一句，程序不会退出）
        jFrame.setResizable(false);
        // 2. 创建中间容器（面板容器）
        jPanel = new JPanel(); // 创建面板容器，使用默认的布局管理器
        jPanel.setLayout(null);
        // 3. 创建一个基本组件（按钮），并添加到 面板容器 中
        //添加插入成交按钮
        jBtnInsert = new JButton("插入成交表");
        jBtnInsert.setBounds(320, 10, 100, 30);
        jPanel.add(jBtnInsert);
        //添加插入汇总按钮
        jBtnAllBook = new JButton("插入汇总表");
        jBtnAllBook.setBounds(320, 50, 100, 30);
        jPanel.add(jBtnAllBook);
        //添加清除按钮
        jBtnClear = new JButton("清除");
        jBtnClear.setBounds(320, 90, 100, 30);
        jPanel.add(jBtnClear);
        //插入选择城区标签
        jLabel2 = new JLabel();
        jLabel2.setBounds(320, 130, 100, 30);
        jLabel2.setText("选择城区");
        jPanel.add(jLabel2);
        //插入下拉列表
        jComboBox = new JComboBox(areas);
        jComboBox.setBounds(320, 160, 100, 30);
        jPanel.add(jComboBox);
        //插入文档分区按钮
        jBtnDivisive = new JButton("文档分区");
        jBtnDivisive.setBounds(320, 200, 100, 30);
        jPanel.add(jBtnDivisive);
        //插入文字标签
        jLabel = new JLabel();
        jLabel.setBounds(320, 260, 100, 30);
        jLabel.setFont(new Font("微软雅黑", 1, 23));
        jLabel.setForeground(Color.green);
        jPanel.add(jLabel);
        //插入进度条
        jProgressBar = new JProgressBar();
        jProgressBar.setBounds(10, 300, 410, 30);
        jProgressBar.setStringPainted(true);
        jPanel.add(jProgressBar);
        //插入文本输入框
        jTextArea = new JTextArea();
        jTextArea.setBounds(10, 10, 300, 280);
        jTextArea.setLineWrap(true);
        jTextArea.setWrapStyleWord(true);
        jTextArea.setEditable(true);
        jPanel.add(jTextArea);

        // 4. 把 面板容器 作为窗口的内容面板 设置到 窗口
        jFrame.setContentPane(jPanel);
        // 5. 显示窗口，前面创建的信息都在内存中，通过 jFrame.setVisible(true) 把内存中的窗口显示在屏幕上。
        jFrame.setVisible(true);
        jBtnInsert.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (jTextArea.getText() != null & !jTextArea.getText().equals("")) {
                    fullText = jTextArea.getText();
                    File file = new File("E://联通宽带成交记录.xlsx");
                    if (!file.isFile() && !file.exists()) {// 判断文件是否存在
                        excelUtil.createExcel(title, "E://联通宽带成交记录.xlsx", "每日成交汇总表");
                        excelUtil.addExcel(ParseText.this, parseText(fullText, true), "E://联通宽带成交记录.xlsx");
                    } else {
                        excelUtil.addExcel(ParseText.this, parseText(fullText, true), "E://联通宽带成交记录.xlsx");
                    }
                } else {
                    jLabel.setText("无内容");
                    jLabel.setForeground(Color.RED);
                }
            }
        });
        jBtnAllBook.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (jTextArea.getText() != null & !jTextArea.getText().equals("")) {
                    fullText = jTextArea.getText();
                    File file = new File("E://主管每日业绩汇总表.xlsx");
                    if (!file.isFile() && !file.exists()) {// 判断文件是否存在
                        excelUtil.createExcel(title2, "E://主管每日业绩汇总表.xlsx", "每日成交汇总表");
                        excelUtil.addExcel(ParseText.this, parseText(fullText, false), "E://主管每日业绩汇总表.xlsx");

                    } else {
                        excelUtil.addExcel(ParseText.this, parseText(fullText, false), "E://主管每日业绩汇总表.xlsx");
                    }
                } else {
                    jLabel.setText("无内容");
                    jLabel.setForeground(Color.RED);
                }
            }
        });
        jBtnClear.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                jTextArea.setText("");
                jLabel.setText("已清空");
                jLabel.setForeground(Color.green);
            }
        });
        jBtnDivisive.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                //todo  区分列表
                new Thread() {
                    @Override
                    public void run() {
                        String item = jComboBox.getSelectedItem().toString();
                        excelUtil.readExcel("E://222.xlsx", "E://" + item + ".xlsx", item, areaTitle);
                        for (int i = 0; i <= 100; i++) {
                            try {
                                Thread.sleep(100);
                            } catch (InterruptedException e) {
                                e.printStackTrace();
                            }
                            jProgressBar.setValue(i);
                        }
                        jProgressBar.setString("分区完成！");
                    }
                }.start();
            }
        });
    }

    private String[] parseText(String text, boolean isFirst) {
        str = text.replaceAll("\\s*", "");
        if (!str.contains(keyWord[0])) {
            jLabel.setText("内容有误");
            throw new RuntimeException("输入内容错误");
        }
        account = getStr("宽带账号：", "地址：");
        data1[0] = account;
        coustomerAddress = getStr("地址：", "套餐：");
        data1[1] = coustomerAddress;
        coustomerName = getStr("姓名：", "电话：");
        data1[2] = coustomerName;
        coustomerMobile = getStr("电话：", "宽带账号：");
        data1[3] = coustomerMobile;
        price = getStr("收费：", "备注：");
        if (price.equals("0")) {
            data1[4] = "已缴费";
        } else {
            data1[4] = "未缴费";
        }
        workName = str.substring(0, str.indexOf("客户"));
        data1[5] = workName;
        status = str.substring(str.indexOf("备注：") + ("备注：").length(), str.length() - coustomerMobile.length());
        if (str.contains("--") || str.contains("==")) {
            lastStr = str.substring(str.indexOf("--"), str.length());
            newNumber = str.substring(str.length() - coustomerMobile.length() - lastStr.length(), str.indexOf("--"));
            status = str.substring(str.indexOf("备注：") + ("备注：").length(), str.length() - lastStr.length());
            eMail = getStr("--", "==");
        } else {
            eMail = "";
            lastStr = "";
            newNumber = str.substring(str.length() - coustomerMobile.length(), str.length());
            status = str.substring(str.indexOf("备注：") + ("备注：").length(), str.length());
        }
        date = new SimpleDateFormat("MM-dd").format(new Date());
        data1[6] = date;
        data2[0] = "2019-" + date;
        data2[1] = coustomerName;
        data2[2] = "";
        data2[3] = account;
        data2[4] = newNumber;
        data2[5] = price;
        data2[6] = status;
        data2[7] = eMail;
        if (isFirst) {
            return data1;
        } else {
            return data2;
        }

    }

    private String getStr(String arg1, String arg2) {
        return str.substring(str.indexOf(arg1) + arg1.length(), str.indexOf(arg2));
    }

    public static void main(String args[]) {
        ExcelUtil excelUtil = new ExcelUtil();
        ParseText readText = new ParseText(excelUtil);
        readText.initView();
    }

    @Override
    public void onSuccess() {
        jLabel.setForeground(Color.green);
        jLabel.setText("插入成功");
    }

    @Override
    public void onError() {
        jLabel.setForeground(Color.red);
        jLabel.setText("插入失败");
    }
}
