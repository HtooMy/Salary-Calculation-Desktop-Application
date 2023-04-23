import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Desktop;
import java.awt.EventQueue;
import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DecimalFormat;

import javax.sound.sampled.AudioInputStream;
import javax.sound.sampled.AudioSystem;
import javax.sound.sampled.Clip;
import javax.sound.sampled.LineUnavailableException;
import javax.sound.sampled.UnsupportedAudioFileException;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.ScrollPaneConstants;
import javax.swing.border.TitledBorder;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.awt.Font;

public class SalaryCalculation {

	private JFrame frame;
	private JTable table;
	private JTextField presenceDayText;
	private JTextField datesInAMonthText;
	private JTextField salaryText;
	private JTextField basicSalaryText;
	private JTextField companyNameText;
	private JTextField employeeNameText;
	private JTextField firstDayOfWorkText;
	private JTextField idText;
	private JTextField resultSalaryText;
	private JTextField actualSalaryText;
	private JTextField textField_12;
	private JTextField descriptionText;
	private JTextField confirmedSalaryText;
	private JTextField officerBonusText;
	private JTextField workRegurityBonusText;
	private JTextField foodProvidenceText;
	private JTextField anualBonusText;
	private JTextField workingMonthText;
	private JTextField absenceFineText;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					SalaryCalculation window = new SalaryCalculation();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Application Methods
	 */

	/**
	 * Create the application.
	 */
	public SalaryCalculation() {
		connect();
		initialize();
		loadTable();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 1307, 536);
		frame.setTitle("Employee Salary Calculation");
		frame.setExtendedState(JFrame.MAXIMIZED_BOTH);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		JPanel bottomMainPanel = new JPanel();
		frame.getContentPane().add(bottomMainPanel, BorderLayout.SOUTH);
		bottomMainPanel.setLayout(new GridLayout(0, 3, 0, 0));

		/**
		 * Employee Data Section (Bottom left Bar)
		 */
		JPanel bottomRightPanel = new JPanel();
		bottomRightPanel.setBackground(Color.PINK);
		bottomRightPanel
				.setBorder(new TitledBorder(null, "Employee Data", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		bottomMainPanel.add(bottomRightPanel);
		bottomRightPanel.setLayout(new GridLayout(0, 2, 0, 0));

		JLabel id = new JLabel("ID");
		bottomRightPanel.add(id);

		idText = new JTextField();
		bottomRightPanel.add(idText);
		idText.setColumns(10);

		JLabel companyName = new JLabel("公司名字");
		bottomRightPanel.add(companyName);

		companyNameText = new JTextField();
		bottomRightPanel.add(companyNameText);
		companyNameText.setColumns(10);

		JLabel employeeName = new JLabel("姓名");
		bottomRightPanel.add(employeeName);

		employeeNameText = new JTextField();
		bottomRightPanel.add(employeeNameText);
		employeeNameText.setColumns(10);

		JLabel burmeseName = new JLabel("缅名");
		bottomRightPanel.add(burmeseName);

		burmeseNameText = new JTextField();
		bottomRightPanel.add(burmeseNameText);
		burmeseNameText.setColumns(10);

		JLabel firstDayOfWork = new JLabel("入职时间");
		bottomRightPanel.add(firstDayOfWork);

		firstDayOfWorkText = new JTextField();
		bottomRightPanel.add(firstDayOfWorkText);
		firstDayOfWorkText.setColumns(10);

		JLabel workingMonth = new JLabel("转正时间");
		bottomRightPanel.add(workingMonth);

		workingMonthText = new JTextField();
		bottomRightPanel.add(workingMonthText);
		workingMonthText.setColumns(10);

		JLabel presenceDay = new JLabel("实际计薪天数");
		bottomRightPanel.add(presenceDay);

		presenceDayText = new JTextField();
		bottomRightPanel.add(presenceDayText);
		presenceDayText.setColumns(10);

		/**
		 * Salary Calculation Section (Central Panel)
		 */
		JPanel bottomLeftPanel = new JPanel();
		bottomLeftPanel.setBackground(Color.PINK);
		bottomLeftPanel.setBorder(
				new TitledBorder(null, "Salary Calculation", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		bottomMainPanel.add(bottomLeftPanel);
		bottomLeftPanel.setLayout(new GridLayout(0, 2, 0, 0));

		JLabel dates = new JLabel("本月计薪天数");
		bottomLeftPanel.add(dates);

		datesInAMonthText = new JTextField();
		bottomLeftPanel.add(datesInAMonthText);
		datesInAMonthText.setColumns(10);

		JLabel salary = new JLabel("工资");
		bottomLeftPanel.add(salary);

		salaryText = new JTextField();
		bottomLeftPanel.add(salaryText);
		salaryText.setColumns(10);

		JLabel basicSalary = new JLabel(" 基本工资");
		bottomLeftPanel.add(basicSalary);

		basicSalaryText = new JTextField();
		bottomLeftPanel.add(basicSalaryText);
		basicSalaryText.setColumns(10);

		JLabel officerBonus = new JLabel("岗位工资");
		bottomLeftPanel.add(officerBonus);

		officerBonusText = new JTextField();
		bottomLeftPanel.add(officerBonusText);
		officerBonusText.setColumns(10);

		JLabel workRegurityBonus = new JLabel("考核工资");
		bottomLeftPanel.add(workRegurityBonus);

		workRegurityBonusText = new JTextField();
		bottomLeftPanel.add(workRegurityBonusText);
		workRegurityBonusText.setColumns(10);

		JLabel foodProvidence = new JLabel("餐补费用");
		bottomLeftPanel.add(foodProvidence);

		foodProvidenceText = new JTextField();
		bottomLeftPanel.add(foodProvidenceText);
		foodProvidenceText.setColumns(10);

		JLabel anualBonus = new JLabel("工龄工资");
		bottomLeftPanel.add(anualBonus);

		anualBonusText = new JTextField();
		bottomLeftPanel.add(anualBonusText);
		anualBonusText.setColumns(10);

		JPanel topPanel = new JPanel();
		topPanel.setBackground(Color.PINK);
		frame.getContentPane().add(topPanel, BorderLayout.NORTH);

		/**
		 * Button Sections (Top Bar)
		 */
		topPanel.setLayout(new FlowLayout(FlowLayout.CENTER, 5, 5));

		JLabel lblNewLabel = new JLabel("Worker Salary:");
		lblNewLabel.setForeground(Color.BLACK);
		topPanel.add(lblNewLabel);

		JButton normalEmployeeBtn = new JButton("New worker");
		topPanel.add(normalEmployeeBtn);
		normalEmployeeBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				salaryText.setText("5000");
				basicSalaryText.setText("2500");
				officerBonusText.setText("1500");
				workRegurityBonusText.setText("1000");
				foodProvidenceText.setText("2500");
				actualSalaryText.setText("7500");
			}
		});

		JButton threeMonthEmployee = new JButton("3-month worker");
		topPanel.add(threeMonthEmployee);
		threeMonthEmployee.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				salaryText.setText("6500");
				basicSalaryText.setText("4000");
				officerBonusText.setText("1500");
				workRegurityBonusText.setText("1000");
				foodProvidenceText.setText("2500");
				actualSalaryText.setText("9000");
			}
		});

		JButton hardWorkEmployeeBtn = new JButton("Hard working worker");
		topPanel.add(hardWorkEmployeeBtn);
		hardWorkEmployeeBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				salaryText.setText("6500");
				basicSalaryText.setText("4000");
				officerBonusText.setText("1500");
				workRegurityBonusText.setText("1000");
				foodProvidenceText.setText("2500");
				actualSalaryText.setText("9000");
			}
		});

		JButton hardWorkThreeMonthsEmployee = new JButton("Hard working 3-month worker");
		topPanel.add(hardWorkThreeMonthsEmployee);
		hardWorkThreeMonthsEmployee.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				salaryText.setText("6500");
				basicSalaryText.setText("4000");
				officerBonusText.setText("2500");
				workRegurityBonusText.setText("1000");
				foodProvidenceText.setText("2500");
				actualSalaryText.setText("10000");
			}
		});

		JButton agencyBtn = new JButton("Agency worker");
		topPanel.add(agencyBtn);
		agencyBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				salaryText.setText("8000");
				basicSalaryText.setText("4000");
				officerBonusText.setText("2500");
				workRegurityBonusText.setText("1500");
				foodProvidenceText.setText("2500");
				actualSalaryText.setText("10500");
			}
		});

		JButton angencyWorkerTwoBtn = new JButton("Angency worker 2");
		angencyWorkerTwoBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				salaryText.setText("9500");
				basicSalaryText.setText("5500");
				officerBonusText.setText("2500");
				workRegurityBonusText.setText("1500");
				foodProvidenceText.setText("2500");
				actualSalaryText.setText("12000");
			}
		});
		topPanel.add(angencyWorkerTwoBtn);

		JButton actualSalaryCalculationBtn = new JButton("Result Salary Calculation");
		topPanel.add(actualSalaryCalculationBtn, BorderLayout.CENTER);
		actualSalaryCalculationBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				resultSalaryCalculation();
			}
		});

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_ALWAYS);
		frame.getContentPane().add(scrollPane, BorderLayout.CENTER);

		/**
		 * Table Section (Center Panel)
		 */
		table = new JTable();
		scrollPane.setViewportView(table);

		JPanel resultPanel = new JPanel();
		resultPanel.setBackground(Color.PINK);
		resultPanel.setBorder(new TitledBorder(null, "Result", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		bottomMainPanel.add(resultPanel);
		resultPanel.setLayout(new GridLayout(0, 2, 0, 0));
				
						JLabel absenceFine = new JLabel("事假扣款");
						resultPanel.add(absenceFine);
		
				absenceFineText = new JTextField();
				resultPanel.add(absenceFineText);
				absenceFineText.setColumns(10);
		JLabel resultSalary = new JLabel("应发小计");
		resultPanel.add(resultSalary);

		resultSalaryText = new JTextField();
		resultPanel.add(resultSalaryText);
		resultSalaryText.setColumns(10);

		JLabel actualSalary = new JLabel("应发放工资总额");
		resultPanel.add(actualSalary);

		actualSalaryText = new JTextField();
		resultPanel.add(actualSalaryText);
		actualSalaryText.setColumns(10);

		JLabel confirmedSalary = new JLabel("5号已发工资中额");
		resultPanel.add(confirmedSalary);

		confirmedSalaryText = new JTextField();
		resultPanel.add(confirmedSalaryText);
		confirmedSalaryText.setColumns(10);

		JLabel description = new JLabel("备注");
		resultPanel.add(description);

		descriptionText = new JTextField();
		resultPanel.add(descriptionText);
		descriptionText.setColumns(10);

		JPanel panel_1 = new JPanel();
		panel_1.setBackground(Color.YELLOW);
		resultPanel.add(panel_1);
		panel_1.setLayout(new BorderLayout(0, 0));

		JLabel idSearchLable = new JLabel("id ဖြင့်ရှာခြင်း");
		idSearchLable.setFont(new Font("Lucida Grande", Font.PLAIN, 15));
		panel_1.add(idSearchLable);

		JPanel panel_2 = new JPanel();
		panel_2.setBackground(Color.YELLOW);
		resultPanel.add(panel_2);
		panel_2.setLayout(new BorderLayout(0, 0));

		idSearch = new JTextField();
		panel_2.add(idSearch);
		idSearch.addKeyListener(new KeyAdapter() {
			@Override
			public void keyReleased(KeyEvent e) {
				try {
					String pID = idSearch.getText();
					pst = con.prepareStatement("select * from Employee_Salary where id = ?");
					pst.setString(1, pID);
					ResultSet rs = pst.executeQuery();

					if (rs.next() == true) {
						String iD = rs.getString(1);
						String CompanyName = rs.getString(2);
						String EmployeeName = rs.getString(3);
						String BurmeseName = rs.getString(4);
						String FirstDayOfWork = rs.getString(5);
						String WorkingMonth = rs.getString(6);
						String PresenceDay = rs.getString(7);
						String AbsenceFine = rs.getString(8);
						String DatesInAMonth = rs.getString(9);
						String Salary = rs.getString(10);
						String BasicSalary = rs.getString(11);
						String OfficierBonus = rs.getString(12);
						String WorkRegurityBonus = rs.getString(13);
						String FoodProvidence = rs.getString(14);
						String AnualBonus = rs.getString(15);
						String ResultSalary = rs.getString(16);
						String ActualSalary = rs.getString(17);
						String ConfirmedSalary = rs.getString(18);
						String Description = rs.getString(19);

						presenceDayText.setText(PresenceDay);
						datesInAMonthText.setText(DatesInAMonth);
						salaryText.setText(Salary);
						basicSalaryText.setText(BasicSalary);
						companyNameText.setText(CompanyName);
						employeeNameText.setText(EmployeeName);
						burmeseNameText.setText(BurmeseName);
						firstDayOfWorkText.setText(FirstDayOfWork);
						idText.setText(iD);
						resultSalaryText.setText(ResultSalary);
						actualSalaryText.setText(ActualSalary);
						descriptionText.setText(Description);
						confirmedSalaryText.setText(ConfirmedSalary);
						officerBonusText.setText(OfficierBonus);
						workRegurityBonusText.setText(WorkRegurityBonus);
						foodProvidenceText.setText(FoodProvidence);
						anualBonusText.setText(AnualBonus);
						workingMonthText.setText(WorkingMonth);
						absenceFineText.setText(AbsenceFine);
					}
				} catch (SQLException ex) {
					ex.printStackTrace();
				}
			}
		});
		idSearch.setColumns(10);

		JLabel label = new JLabel("");
		resultPanel.add(label);

		JPanel panel = new JPanel();
		panel.setBackground(Color.PINK);
		frame.getContentPane().add(panel, BorderLayout.WEST);
		panel.setLayout(new GridLayout(0, 1, 0, 0));

		/**
		 * Side bar Button section
		 */

		/*
		 * Adding button
		 */
		JButton saveBtn = new JButton("Save");
		saveBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					e1.printStackTrace();
				}
				insertData();
			}
		});
		panel.add(saveBtn);

		/*
		 * Updating a table
		 */
		JButton updateBtn = new JButton("Update");
		updateBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				deleteData(idText.getText());
				insertData();
			}
		});
		panel.add(updateBtn);

		JButton deleteBtn = new JButton("Delete");
		deleteBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				String iD = idText.getText();
				if (iD.equals("")) {
					try {
						alertClip();
					} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
					JOptionPane.showMessageDialog(null, "No product is selected");
				} else {
					deleteData(iD);
					showDeleteDialog();
				}
			}
		});
		panel.add(deleteBtn);

		JButton downloadBtn = new JButton("Download ");
		downloadBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

				try {
					exportarExcel(table);
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});

		JButton refreshBtn = new JButton("Refresh");
		refreshBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				table.setModel(new DefaultTableModel());
				loadTable();
				clearText();
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});
		panel.add(refreshBtn);
		panel.add(downloadBtn);

		JButton clearBtn = new JButton("Clear");
		clearBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					clickClip();
				} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				clearText();
			}
		});
		panel.add(clearBtn);
	}

	/**
	 * Clear Method
	 */
	public void clearText() {
		presenceDayText.setText("");
		datesInAMonthText.setText("");
		salaryText.setText("");
		basicSalaryText.setText("");
		companyNameText.setText("");
		employeeNameText.setText("");
		burmeseNameText.setText("");
		firstDayOfWorkText.setText("");
		idText.setText("");
		resultSalaryText.setText("");
		actualSalaryText.setText("");
		descriptionText.setText("");
		confirmedSalaryText.setText("");
		officerBonusText.setText("");
		workRegurityBonusText.setText("");
		foodProvidenceText.setText("");
		anualBonusText.setText("");
		workingMonthText.setText("");
		absenceFineText.setText("");
		idSearch.setText("");
	}

	/**
	 * Application Business Logic (Salary Calculation)
	 */
	public void resultSalaryCalculation() {
		int salaryGotFromCalculation = Integer.valueOf(salaryText.getText());
		int anualBonus = Integer.valueOf(anualBonusText.getText());
		int presenceDays = Integer.valueOf(presenceDayText.getText());
		int totalDate = Integer.valueOf(datesInAMonthText.getText());
		double preSalary = salaryGotFromCalculation + Integer.valueOf(foodProvidenceText.getText());
		double salaryPerDay = preSalary / totalDate;
		double resultSalary = (salaryPerDay * presenceDays) + anualBonus;
		double absenceFine = preSalary - resultSalary;

		// Formating decimal value to two decimal place
		DecimalFormat df = new DecimalFormat("#.##");
		absenceFine = Double.valueOf(df.format(absenceFine));
		resultSalary = Double.valueOf(df.format(resultSalary));

		// Showing the result in the text field
		resultSalaryText.setText(Double.toString(resultSalary));
		confirmedSalaryText.setText(Double.toString(resultSalary));
		absenceFineText.setText(Double.toString(absenceFine));
		try {
			successClip();
		} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		JOptionPane.showMessageDialog(null, "The result salary is " + Double.toString(resultSalary));
	}

	/**
	 * Sound effect functions
	 */
	// Setting URL for sound file
	URL clickSoundURL, successSoundURL, alertSoundURL;

	public void setURL(URL fileName) {
		try {
			AudioInputStream audioStream = AudioSystem.getAudioInputStream(fileName);
			Clip clip = AudioSystem.getClip();
			clip.open(audioStream);
			clip.start();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	// Mouse Click Sound effect
	public void clickClip() throws UnsupportedAudioFileException, IOException, LineUnavailableException {
		clickSoundURL = getClass().getResource("mouse.wav");
		setURL(clickSoundURL);
	}

	// Success Sound effect
	public void successClip() throws UnsupportedAudioFileException, IOException, LineUnavailableException {
		successSoundURL = getClass().getResource("win.wav");
		setURL(successSoundURL);
	}

	// Alert Sound effect
	public void alertClip() throws UnsupportedAudioFileException, IOException, LineUnavailableException {
		alertSoundURL = getClass().getResource("alert.wav");
		setURL(alertSoundURL);
	}

	// Show dialog
	public void showDeleteDialog() {
		// Sound effect
		try {
			successClip();
		} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		JOptionPane.showMessageDialog(null, "Record Delete!!!!!");
		clearText();
		loadTable();
	}

	/**
	 * Connecting database
	 */
	// Variable for connecting SQL
	private Connection con;
	private Statement st;
	private PreparedStatement pst;
	private JTextField burmeseNameText;
	private JTextField idSearch;

	private void connect() {
		// Driver
		try {
			Class.forName("com.mysql.cj.jdbc.Driver");
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// Connecting Database with JDBC
		try {
			con = DriverManager.getConnection(
					"jdbc:mysql://sql12.freesqldatabase.com:3306/sql12592806?useUnicode=true&characterEncoding=utf8",
					"sql12592806", "YeeShin527845");
			st = con.createStatement();

			// Success in connecting database
			try {
				successClip();
			} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			JOptionPane.showMessageDialog(null, "Connecting Database");
		} catch (SQLException e) {

			// Fain in connecting database
			try {
				alertClip();
			} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			JOptionPane.showMessageDialog(null, "Connection fail");
			e.printStackTrace();
		}
	}

	// Loading Table
	public void loadTable() {
		try {
			table.setModel(new DefaultTableModel());
			String query = "SELECT * FROM Employee_Salary ORDER BY `应发放工资总额` ASC;"; // " ORDER BY 应发放工资总额 ASC;"
			ResultSet rs = st.executeQuery(query);
			ResultSetMetaData rsmd = rs.getMetaData();
			DefaultTableModel model = (DefaultTableModel) table.getModel();

			int cols = rsmd.getColumnCount();
			String[] colName = new String[cols];
			for (int i = 0; i < cols; i++) {
				colName[i] = rsmd.getColumnName(i + 1);
				model.setColumnIdentifiers(colName);

				String iD, CompanyName, EmployeeName, BurmeseName, FirstDayOfWork, WorkingMonth, PresenceDay,
						AbsenceFine, DatesInAMonth, Salary, BasicSalary, OfficierBonus, WorkRegurityBonus,
						FoodProvidence, AnualBonus, ResultSalary, ActualSalary, ConfirmedSalary, Description;

				while (rs.next()) {
					iD = rs.getString(1);
					CompanyName = rs.getString(2);
					EmployeeName = rs.getString(3);
					BurmeseName = rs.getString(4);
					FirstDayOfWork = rs.getString(5);
					WorkingMonth = rs.getString(6);
					PresenceDay = rs.getString(7);
					AbsenceFine = rs.getString(8);
					DatesInAMonth = rs.getString(9);
					Salary = rs.getString(10);
					BasicSalary = rs.getString(11);
					OfficierBonus = rs.getString(12);
					WorkRegurityBonus = rs.getString(13);
					FoodProvidence = rs.getString(14);
					AnualBonus = rs.getString(15);
					ResultSalary = rs.getString(16);
					ActualSalary = rs.getString(17);
					ConfirmedSalary = rs.getString(18);
					Description = rs.getString(19);

					// Adding rows in table
					String[] row = { iD, CompanyName, EmployeeName, BurmeseName, FirstDayOfWork, WorkingMonth,
							PresenceDay, AbsenceFine, DatesInAMonth, Salary, BasicSalary, OfficierBonus,
							WorkRegurityBonus, FoodProvidence, AnualBonus, ResultSalary, ActualSalary, ConfirmedSalary,
							Description };
					model.addRow(row);
				}
			}
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();

		}
	}

	/*
	 * Inserting data
	 */
	public void insertData() {
		String iD, CompanyName, EmployeeName, BurmeseName, FirstDayOfWork, WorkingMonth, PresenceDay, AbsenceFine,
				DatesInAMonth, Salary, BasicSalary, OfficierBonus, WorkRegurityBonus, FoodProvidence, AnualBonus,
				ResultSalary, ActualSalary, ConfirmedSalary, Description;

		iD = idText.getText();
		CompanyName = companyNameText.getText();
		EmployeeName = employeeNameText.getText();
		BurmeseName = burmeseNameText.getText();
		FirstDayOfWork = firstDayOfWorkText.getText();
		WorkingMonth = workingMonthText.getText();
		PresenceDay = presenceDayText.getText();
		AbsenceFine = absenceFineText.getText();
		DatesInAMonth = datesInAMonthText.getText();
		Salary = salaryText.getText();
		BasicSalary = basicSalaryText.getText();
		OfficierBonus = officerBonusText.getText();
		WorkRegurityBonus = workRegurityBonusText.getText();
		FoodProvidence = foodProvidenceText.getText();
		AnualBonus = anualBonusText.getText();
		ResultSalary = resultSalaryText.getText();
		ActualSalary = actualSalaryText.getText();
		ConfirmedSalary = confirmedSalaryText.getText();
		Description = descriptionText.getText();

		try {
			pst = con.prepareStatement("INSERT INTO `Employee_Salary` VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
			pst.setString(1, iD);
			pst.setString(2, CompanyName);
			pst.setString(3, EmployeeName);
			pst.setString(4, BurmeseName);
			pst.setString(5, FirstDayOfWork);
			pst.setString(6, WorkingMonth);
			pst.setString(7, PresenceDay);
			pst.setString(8, AbsenceFine);
			pst.setString(9, DatesInAMonth);
			pst.setString(10, Salary);
			pst.setString(11, BasicSalary);
			pst.setString(12, OfficierBonus);
			pst.setString(13, WorkRegurityBonus);
			pst.setString(14, FoodProvidence);
			pst.setString(15, AnualBonus);
			pst.setString(16, ResultSalary);
			pst.setString(17, ActualSalary);
			pst.setString(18, ConfirmedSalary);
			pst.setString(19, Description);
			pst.executeUpdate();

			try {
				successClip();
			} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			JOptionPane.showMessageDialog(null, "Successfully save data!");
			clearText();
			loadTable();

		} catch (SQLException e1) {
			try {
				alertClip();
			} catch (UnsupportedAudioFileException | IOException | LineUnavailableException e4) {
				// TODO Auto-generated catch block
				e4.printStackTrace();
			}
			JOptionPane.showMessageDialog(null, "Something went wrong");
			e1.printStackTrace();
		}
	}

	/*
	 * Deleting data
	 */
	public void deleteData(String id) {
		try {
			pst = con.prepareStatement("delete from Employee_Salary where id = ?");
			pst.setString(1, id);
			pst.executeUpdate();
		} catch (SQLException e1) {
			e1.printStackTrace();
		}
	}

	/**
	 * Download the table as an excel file
	 */
	public void openFile(String file) {
		try {
			File path = new File(file);
			Desktop.getDesktop().open(path);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void exportarExcel(JTable jtable) throws IOException {
		try {
			JFileChooser jFilerChooser = new JFileChooser();
			jFilerChooser.showSaveDialog(jtable);
			File saveFile = jFilerChooser.getSelectedFile();
			if (saveFile != null) {
				saveFile = new File(saveFile.toString() + ".xlsx");
				Workbook wb = new XSSFWorkbook();
				Sheet sheet = wb.createSheet("Salary_Calculation");
				Row rowCol = sheet.createRow(0);

				for (int i = 0; i < jtable.getColumnCount(); i++) {
					Cell cell = rowCol.createCell(i);
					cell.setCellValue(jtable.getColumnName(i));
				}

				for (int j = 0; j < jtable.getRowCount(); j++) {
					Row row = sheet.createRow(j + 1);
					for (int k = 0; k < jtable.getColumnCount(); k++) {
						Cell cell = row.createCell(k);
						if (jtable.getValueAt(j, k) != null) {
							cell.setCellValue(jtable.getValueAt(j, k).toString());
						}
					}
				}

				FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
				wb.write(out);
				wb.close();
				openFile(saveFile.toString());

			}
		} catch (Exception e) {
			JOptionPane.showMessageDialog(null, "Something went wrong! You cannot download it");
		}

	}
}
