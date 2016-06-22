package App;

/* File: MembershipReportGenerator.java
 * Programmer: Garrett Erven (Student at Clover High School)
 * Date: May 27th, 2016
 * Purpose: The purpose of this program is to maintain current member data and to output reports of
 * member statistics for the national office of Future Business Leaders of America Phi-Beta-Lambda. */

/* Rather than a dedicated help menu, tool tip messages appear when the user hovers the cursor over a GUI component momentarily. These messages
 * explain to the user what each component does and aids them in completing their intended task. This alternative is much easier to access. */

/* Reference the readMe.rtf accompanied with the program files for extensive program documentation. */

import java.awt.CardLayout;
import java.awt.Desktop;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.URISyntaxException;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.apache.commons.io.FileUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import GUI.Button;
import GUI.Frame;
import GUI.Label;
import GUI.Panel;
import GUI.TextField;

public class MembershipReportGenerator
{
	public static void main(String[] args)
	{
		// Establishes a directory on the user's computer and creates a blank master file if one does not currently exist
		setResources();
		
		// Array list of Member objects with an unlimited capacity
		ArrayList<Member> memberList = new ArrayList<Member>();
		
		// Reads the master file for current member data and adds it to the array list of Member objects
		readMembers(memberList);
		
		// Frame for use as the application's window
		Frame frame = new Frame("logos/FBLALogo.png");
		
		// Layout that allows for a graphical user interface with multiple screens
		CardLayout cardLayout = new CardLayout();
		
		// Layout that allows for the organization of text fields on panel2
		GridBagLayout gridBagLayout = new GridBagLayout();
		GridBagConstraints constraints = new GridBagConstraints();
		
		// A container and three panels for use as multiple screens
		Panel container = new Panel();
		Panel panel1 = new Panel();
		Panel panel2 = new Panel();
		Panel panel3 = new Panel();
		
		// Buttons for navigating the three panels
		Button editCurrentMembersButton = new Button("buttonIcons/editCurrentMembersButtonIcon.png", "Edit the data of current members");
		Button addNewMemberButton = new Button("buttonIcons/addNewMemberButtonIcon.png", "Add a new member to the database");
		Button memberDuesReportButton = new Button("buttonIcons/memberDuesReportButtonIcon.png", "Generate a report of members that owe dues");
		Button seniorEmailReportButton = new Button("buttonIcons/seniorEmailReportButtonIcon.png", "Generate a report of senior emails");
		Button fblaWebsiteButton = new Button("buttonIcons/fblaWebPageButtonIcon.png", "Open the FBLA-PBL website in your internet browser");
		Button addMemberButton = new Button("buttonIcons/addMemberButtonIcon.png", "Add the member to the database");
		Button backButton = new Button("buttonIcons/backButtonIcon.png", "Return to the previous screen");
		Button viewButton = new Button("buttonIcons/viewButtonIcon.png", "View the report in Microsoft Excel");
		Button printButton = new Button("buttonIcons/printButtonIcon.png", "Print the contents of the report on paper");
		Button exportButton = new Button("buttonIcons/exportButtonIcon.png", "Export the contents of the report as an .xls file");
		Button backButton2 = new Button("buttonIcons/backButtonIcon.png", "Return to the previous screen");
		
		// Text fields for collecting the data of a new member
		TextField memberNumberTextField = new TextField();
		TextField firstNameTextField = new TextField();
		TextField lastNameTextField = new TextField();
		TextField gradeTextField = new TextField();
		TextField schoolTextField = new TextField();
		TextField stateTextField = new TextField();
		TextField emailTextField = new TextField();
		TextField yearJoinedTextField = new TextField();
		TextField statusTextField = new TextField();
		TextField amountOwedTextField = new TextField();
		
		// Labels for decorational purposes and for distinguishing the text fields on panel2
		Label logo = new Label("logos/FBLA-PBLLogo.png");
		Label logo2 = new Label("logos/FBLA-PBLLogo.png");
		Label cityscape = new Label("logos/cityscape.png");
		Label cityscape2 = new Label("logos/cityscape.png");
		Label memberNumberLabel = new Label("Member Number:", "Example: 123456");
		Label firstNameLabel = new Label("First Name:", "Example: John");
		Label lastNameLabel = new Label("Last Name:", "Example: Doe");
		Label gradeLabel = new Label("Grade:", "Example: 9");
		Label schoolLabel = new Label("School:", "Example: Wood Crest High School");
		Label stateLabel = new Label("State:", "Example: Florida");
		Label emailLabel = new Label("Email:", "Example: jdoe@email.com");
		Label yearJoinedLabel = new Label("Year Joined:", "Example: 2010");
		Label statusLabel = new Label("Status:", "Example: Active");
		Label amountOwedLabel = new Label("Amount Owed:", "Example: 15.00");
		
		/* Adds buttons, creates action listeners, and adds labels to panel1 */
		panel1.add(editCurrentMembersButton);
		editCurrentMembersButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{
				if(memberList.size() == 0)
					JOptionPane.showMessageDialog(null, "Please add new members first!", "No Members", JOptionPane.WARNING_MESSAGE);
				else
					/* Generates an error message in the command prompt and closes the application if
					 * unable to open the master file */
					try
					{
						if(System.getProperty("os.name").startsWith("Mac OS X"))
							Desktop.getDesktop().open(new File(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/masterFile.xls"));
						else
							Desktop.getDesktop().open(new File("c:\\FBLA-PBL Membership Report Generator Program Files\\masterFile.xls"));
					}
					catch(IOException exception)
					{
						System.err.println("Error: Unable to open the master file!");
						System.exit(0);
					}
			}
		});
		panel1.add(addNewMemberButton);
		addNewMemberButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{
				cardLayout.show(container, "2");
			}
		});
		panel1.add(memberDuesReportButton);
		memberDuesReportButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{
				cardLayout.show(container, "3");
				memberList.clear();
				generateMemberDuesReport(sortMembers(readMembers(memberList)));
			}
		});
		panel1.add(seniorEmailReportButton);
		seniorEmailReportButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{
				cardLayout.show(container, "3");
				memberList.clear();
				generateSeniorEmailReport(sortMembers(readMembers(memberList)));
			}
		});
		panel1.add(fblaWebsiteButton);
		fblaWebsiteButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{
				/* Generates an error message in the command prompt and closes the application if
				 * the FBLA-PBL website cannot be opened in the user's default web browser */
				try
				{
					Desktop.getDesktop().browse(new URI("http://www.fbla-pbl.org/"));
				}
				catch(IOException | URISyntaxException exception)
				{
					System.err.println("Error: Unable to open the FBLA-PBL web page in the user's default web browser!");
					System.exit(0);
				}
			}
		});
		panel1.add(logo);
		panel1.add(cityscape);
		
		/* Sets layout, adds labels, adds text fields, adds buttons, and creates action listeners on panel2 */
		panel2.setLayout(gridBagLayout);
		constraints.fill = GridBagConstraints.HORIZONTAL;
		constraints.fill = GridBagConstraints.EAST;
		constraints.weighty = 1;
		constraints.gridx = 0;
		constraints.gridy = 0;
		panel2.add(memberNumberLabel, constraints);
		constraints.gridx = 1;
		constraints.gridy = 0;
		panel2.add(memberNumberTextField, constraints);
		constraints.gridx = 2;
		constraints.gridy = 0;
		panel2.add(firstNameLabel, constraints);
		constraints.gridx = 3;
		constraints.gridy = 0;
		panel2.add(firstNameTextField, constraints);
		constraints.gridx = 0;
		constraints.gridy = 2;
		panel2.add(lastNameLabel, constraints);
		constraints.gridx = 1;
		constraints.gridy = 2;
		panel2.add(lastNameTextField, constraints);
		constraints.gridx = 2;
		constraints.gridy = 2;
		panel2.add(gradeLabel, constraints);
		constraints.gridx = 3;
		constraints.gridy = 2;
		panel2.add(gradeTextField, constraints);
		constraints.gridx = 0;
		constraints.gridy = 4;
		panel2.add(schoolLabel, constraints);
		constraints.gridx = 1;
		constraints.gridy = 4;
		panel2.add(schoolTextField, constraints);
		constraints.gridx = 2;
		constraints.gridy = 4;
		panel2.add(stateLabel, constraints);
		constraints.gridx = 3;
		constraints.gridy = 4;
		panel2.add(stateTextField, constraints);
		constraints.gridx = 0;
		constraints.gridy = 6;
		panel2.add(emailLabel, constraints);
		constraints.gridx = 1;
		constraints.gridy = 6;
		panel2.add(emailTextField, constraints);
		constraints.gridx = 2;
		constraints.gridy = 6;
		panel2.add(yearJoinedLabel, constraints);
		constraints.gridx = 3;
		constraints.gridy = 6;
		panel2.add(yearJoinedTextField, constraints);
		constraints.gridx = 0;
		constraints.gridy = 8;
		panel2.add(statusLabel, constraints);
		constraints.gridx = 1;
		constraints.gridy = 8;
		panel2.add(statusTextField, constraints);
		constraints.gridx = 2;
		constraints.gridy = 8;
		panel2.add(amountOwedLabel, constraints);
		constraints.gridx = 3;
		constraints.gridy = 8;
		panel2.add(amountOwedTextField, constraints);
		constraints.gridx = 1;
		constraints.gridy = 10;
		panel2.add(addMemberButton, constraints);
		addMemberButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{
				if(memberNumberTextField.getText().equals("") || firstNameTextField.getText().equals("") || lastNameTextField.getText().equals("") || gradeTextField.getText().equals("") || schoolTextField.getText().equals("") || stateTextField.getText().equals("") || emailTextField.getText().equals("") || yearJoinedTextField.getText().equals("") || statusTextField.getText().equals("") || amountOwedTextField.getText().equals(""))
					JOptionPane.showMessageDialog(null, "Please make sure that you have filled in all text fields!", "Invalid Data", JOptionPane.WARNING_MESSAGE);
				
				String memNumstr = memberNumberTextField.getText();
				double memberNumber = Double.parseDouble(memNumstr);
				String fNamestr = firstNameTextField.getText();
				String lNamestr = lastNameTextField.getText();
				String gradstr = gradeTextField.getText();
				double grade = Double.parseDouble(gradstr);
				String schstr = schoolTextField.getText();
				String stastr = stateTextField.getText();
				String emastr = emailTextField.getText();
				String yJoinedstr = yearJoinedTextField.getText();
				double yearJoined = Double.parseDouble(yJoinedstr);
				String statstr = statusTextField.getText();
				String aOwedstr = amountOwedTextField.getText();
				double amountOwed = Double.parseDouble(aOwedstr);
				
				addNewMember(memberList, memberNumber, fNamestr, lNamestr, grade, schstr, stastr, emastr, yearJoined, statstr, amountOwed);
				
				if(JOptionPane.showConfirmDialog(null, "Member sucessfully added! Add another?", "Success", JOptionPane.INFORMATION_MESSAGE) == JOptionPane.YES_NO_OPTION)
				{
					memberNumberTextField.setText(null);
					firstNameTextField.setText(null);
					lastNameTextField.setText(null);
					gradeTextField.setText(null);
					schoolTextField.setText(null);
					stateTextField.setText(null);
					emailTextField.setText(null);
					yearJoinedTextField.setText(null);
					statusTextField.setText(null);
					amountOwedTextField.setText(null);
				}
				else
				{
					memberNumberTextField.setText(null);
					firstNameTextField.setText(null);
					lastNameTextField.setText(null);
					gradeTextField.setText(null);
					schoolTextField.setText(null);
					stateTextField.setText(null);
					emailTextField.setText(null);
					yearJoinedTextField.setText(null);
					statusTextField.setText(null);
					amountOwedTextField.setText(null);
					cardLayout.first(container);
				}
			}
		});
		constraints.gridx = 2;
		constraints.gridy = 10;
		panel2.add(backButton, constraints);
		backButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{
				cardLayout.first(container);
			}
		});
		
		/* Adds buttons, creates action listeners, and adds labels to panel3 */
		panel3.add(viewButton);
		viewButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{
				viewReport();
			}
		});
		panel3.add(printButton);
		printButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{
				printReport();
			}
		});
		panel3.add(exportButton);
		exportButton.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{	
				exportReport();
			}
		});
		panel3.add(backButton2);
		backButton2.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent buttonClicked)
			{
				cardLayout.first(container);
			}
		});
		panel3.add(logo2);
		panel3.add(cityscape2);
		
		// Adds the panels to the container
		container.setLayout(cardLayout);
		container.add(panel1, "1");
		container.add(panel2, "2");
		container.add(panel3, "3");
		cardLayout.show(container, "1");
		
		// Adds the container of panels to the frame
		frame.add(container);
		frame.pack();
		frame.setLocationRelativeTo(null);
		frame.setVisible(true);
	}
	
	public static ArrayList<Member> readMembers(ArrayList<Member> memberList)
	{
		// Temporary variables for the member attributes
		double memNum;
		String fName;
		String lName;
		double grad;
		String sch;
		String sta;
		String ema;
		double yJoined;
		String stat;
		double aOwed;
		
		InputStream input = null;
		
		/* Generates an error message in the command prompt and closes the application if
		 * unable to find the master file */
		try
		{
			if(System.getProperty("os.name").startsWith("Mac OS X"))
				input = new FileInputStream(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/masterFile.xls");
			else
				input = new FileInputStream("c:\\FBLA-PBL Membership Report Generator Program Files\\masterFile.xls");
		}
		catch(FileNotFoundException exception)
		{
			System.err.println("Error: Unable to find the master file!");
			System.exit(0);
		}
		
		Workbook workbook = null;
		
		/* Generates an error message in the command prompt and closes the application if
		 * unable to read the master file */
		try
		{
			workbook = WorkbookFactory.create(input);
		}
		catch(EncryptedDocumentException | InvalidFormatException | IOException exception)
		{
			System.err.println("Error: Unable to read the master file!");
			System.exit(0);
		}
		
		Sheet sheet = workbook.getSheetAt(0);
		
		// Transfers data from the master file to the array list of Member objects
		for(int count = 0; count < sheet.getPhysicalNumberOfRows(); count++)
		{
			int count2 = 0;
			Row row = sheet.getRow(count);
			Cell cell = row.getCell(count2);
			
			memNum = (cell.getNumericCellValue());
			Cell cell1 = row.getCell(count2 + 1);
			fName = (cell1.getStringCellValue());
			Cell cell2 = row.getCell(count2 + 2);
			lName = (cell2.getStringCellValue());
			Cell cell3 = row.getCell(count2 + 3);
			grad = (cell3.getNumericCellValue());
			Cell cell4 = row.getCell(count2 + 4);
			sch = (cell4.getStringCellValue());
			Cell cell5 = row.getCell(count2 + 5);
			sta = (cell5.getStringCellValue());
			Cell cell6 = row.getCell(count2 + 6);
			ema = (cell6.getStringCellValue());
			Cell cell7 = row.getCell(count2 + 7);
			yJoined = (cell7.getNumericCellValue());
			Cell cell8 = row.getCell(count2 + 8);
			stat = (cell8.getStringCellValue());
			Cell cell9 = row.getCell(count2 + 9);
			aOwed = (cell9.getNumericCellValue());
			
			// Constructs member objects that hold the attribute data and adds them to the array list
			Member member = new Member(memNum, fName, lName, grad, sch, sta, ema, yJoined, stat, aOwed);
			memberList.add(member);
		}
		
		return memberList;
	}
	
	public static void editCurrentMember(int indexOfMemberBeingEdited, ArrayList<Member> memberList, double memberNumber, String firstName, String lastName, double grade, String school, String state, String email, double yearJoined, String status, double amountOwed)
	{
		InputStream input = null;
		
		/* Generates an error message in the command prompt and closes the application if
		 * unable to find the master file */
		try
		{
			if(System.getProperty("os.name").startsWith("Mac OS X"))
				input = new FileInputStream(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/masterFile.xls");
			else
				input = new FileInputStream("c:\\FBLA-PBL Membership Report Generator Program Files\\masterFile.xls");
		}
		catch(FileNotFoundException exception)
		{
			System.err.println("Error: Unable to find the master file!");
			System.exit(0);
		}
		
		Workbook workbook = null;
		
		/* Generates an error message in the command prompt and closes the application if
		 * unable to read the master file */
		try
		{
			workbook = WorkbookFactory.create(input);
		}
		catch(EncryptedDocumentException | InvalidFormatException | IOException exception)
		{
			System.err.println("Error: Unable to read the master file!");
			System.exit(0);
		}
		
		Sheet sheet = workbook.getSheetAt(0);
		
		// Makes the desired changes to the master file
		Row row = sheet.createRow(indexOfMemberBeingEdited);
		Cell cell = row.createCell(0);
			
		cell.setCellValue(memberNumber);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue(firstName);
		Cell cell2 = row.createCell(2);
		cell2.setCellValue(lastName);
		Cell cell3 = row.createCell(3);
		cell3.setCellValue(grade);
		Cell cell4 = row.createCell(4);
		cell4.setCellValue(school);
		Cell cell5 = row.createCell(5);
		cell5.setCellValue(state);
		Cell cell6 = row.createCell(6);
		cell6.setCellValue(email);
		Cell cell7 = row.createCell(7);
		cell7.setCellValue(yearJoined);
		Cell cell8 = row.createCell(8);
		cell8.setCellValue(status);
		Cell cell9 = row.createCell(9);
		cell9.setCellValue(amountOwed);
		
		/* Generates an error message in the command prompt and closes the application if
		 * unable to rewrite the master file */
		try
		{
			FileOutputStream output = null;
			
			if(System.getProperty("os.name").startsWith("Mac OS X"))
				output = new FileOutputStream(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/masterFile.xls");
			else
				output = new FileOutputStream("c:\\FBLA-PBL Membership Report Generator Program Files\\masterFile.xls");
			
			workbook.write(output);
			output.close();
		}
		catch(Exception exception)
		{
			System.err.println("Error: Unable to rewrite the master file!");
			System.exit(0);
		}
	}
	
	public static void addNewMember(ArrayList<Member> memberList, double memberNumber, String firstName, String lastName, double grade, String school, String state, String email, double yearJoined, String status, double amountOwed)
	{
		InputStream input = null;
		
		/* Generates an error message in the command prompt and closes the application if
		 * unable to find the master file */
		try
		{
			if(System.getProperty("os.name").startsWith("Mac OS X"))
				input = new FileInputStream(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/masterFile.xls");
			else
				input = new FileInputStream("c:\\FBLA-PBL Membership Report Generator Program Files\\masterFile.xls");
		}
		catch(FileNotFoundException exception)
		{
			System.err.println("Error: Unable to find the master file!");
			System.exit(0);
		}
		
		Workbook workbook = null;
		
		/* Generates an error message in the command prompt and closes the application if
		 * unable to read the master file */
		try
		{
			workbook = WorkbookFactory.create(input);
		}
		catch(EncryptedDocumentException | InvalidFormatException | IOException exception)
		{
			System.err.println("Error: Unable to read the master file!");
			System.exit(0);
		}
		
		Sheet sheet = workbook.getSheetAt(0);
		
		// Transfers data from the text fields to the master file
		Row row = sheet.createRow(sheet.getPhysicalNumberOfRows());
		Cell cell = row.createCell(0);
			
		cell.setCellValue(memberNumber);
		Cell cell1 = row.createCell(1);
		cell1.setCellValue(firstName);
		Cell cell2 = row.createCell(2);
		cell2.setCellValue(lastName);
		Cell cell3 = row.createCell(3);
		cell3.setCellValue(grade);
		Cell cell4 = row.createCell(4);
		cell4.setCellValue(school);
		Cell cell5 = row.createCell(5);
		cell5.setCellValue(state);
		Cell cell6 = row.createCell(6);
		cell6.setCellValue(email);
		Cell cell7 = row.createCell(7);
		cell7.setCellValue(yearJoined);
		Cell cell8 = row.createCell(8);
		cell8.setCellValue(status);
		Cell cell9 = row.createCell(9);
		cell9.setCellValue(amountOwed);
		
		/* Generates an error message in the command prompt and closes the application if
		 * unable to rewrite the master file */
		try
		{
			FileOutputStream output = null;
			
			if(System.getProperty("os.name").startsWith("Mac OS X"))
				output = new FileOutputStream(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/masterFile.xls");
			else
				output = new FileOutputStream("c:\\FBLA-PBL Membership Report Generator Program Files\\masterFile.xls");
			
			workbook.write(output);
			output.close();
		}
		catch(Exception exception)
		{
			System.err.println("Error: Unable to rewrite the master file!");
			System.exit(0);
		}
	}
	
	public static void generateMemberDuesReport(ArrayList<Member> memberList)
	{
		// Variables to record data for the report footer
		int numOfNonactiveMembers = 0;
		int numOfActiveMembers = 0;
		int numOfMembersOwing = 0;
		double totalAmountOwed = 0.0;
		
		// Currency formatter for outputting monetary values
		NumberFormat formatter = NumberFormat.getCurrencyInstance();
		
		// Creates Microsoft Excel workbook
		@SuppressWarnings("resource")
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("Member Dues Report");
		
		// Aligns content to the center of each cell
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		
		/* Sets the page orientation to landscape, limits each page
		 * to 50 lines, and sets the default width of each cell */
		sheet.getPrintSetup().setLandscape(true);
		sheet.setRowBreak(49);
		sheet.setColumnBreak(6);
		sheet.setDefaultColumnWidth(13);
		
		// Creates header cells
		Row headerRow = sheet.createRow(0);
		Cell headerCell = headerRow.createCell(0);
		headerCell.setCellValue("State");
		headerCell.setCellStyle(cellStyle);
		Cell headerCell1 = headerRow.createCell(1);
		headerCell1.setCellValue("Member Number");
		headerCell1.setCellStyle(cellStyle);
		Cell headerCell2 = headerRow.createCell(2);
		headerCell2.setCellValue("First Name");
		headerCell2.setCellStyle(cellStyle);
		Cell headerCell3 = headerRow.createCell(3);
		headerCell3.setCellValue("Last Name");
		headerCell3.setCellStyle(cellStyle);
		Cell headerCell4 = headerRow.createCell(4);
		headerCell4.setCellValue("Year Joined");
		headerCell4.setCellStyle(cellStyle);
		Cell headerCell5 = headerRow.createCell(5);
		headerCell5.setCellValue("Grade");
		headerCell5.setCellStyle(cellStyle);
		Cell headerCell6 = headerRow.createCell(6);
		headerCell6.setCellValue("Amount Owed");
		headerCell6.setCellStyle(cellStyle);
		
		/* Tallys the number of active and inactive members, tallys the number of members that owe dues,
		 * adds the member's information to the report if the member owes dues, and accumulates the total amount owed by all members */
		int row = 1;
		
		for(int count = 0; count < memberList.size(); count++)
		{
			if(memberList.get(count).getstatus().equalsIgnoreCase("Nonactive"))
				numOfNonactiveMembers++;
			
			if(memberList.get(count).getamountOwed() != 0.0)
			{
				numOfMembersOwing++;
				totalAmountOwed += memberList.get(count).getamountOwed();
				
				int cell = 0;
				Row row1 = sheet.createRow(row);
				
				Cell cell1 = row1.createCell(cell);
				cell1.setCellValue(memberList.get(count).getstate());
				cell1.setCellStyle(cellStyle);
				Cell cell2 = row1.createCell(cell + 1);
				cell2.setCellValue(memberList.get(count).getmembershipNumber());
				cell2.setCellStyle(cellStyle);
				Cell cell3 = row1.createCell(cell + 2);
				cell3.setCellValue(memberList.get(count).getfirstName());
				cell3.setCellStyle(cellStyle);
				Cell cell4 = row1.createCell(cell + 3);
				cell4.setCellValue(memberList.get(count).getlastName());
				cell4.setCellStyle(cellStyle);
				Cell cell5 = row1.createCell(cell + 4);
				cell5.setCellValue(memberList.get(count).getyearJoined());
				cell5.setCellStyle(cellStyle);
				Cell cell6 = row1.createCell(cell + 5);
				cell6.setCellValue(memberList.get(count).getgrade());
				cell6.setCellStyle(cellStyle);
				Cell cell7 = row1.createCell(cell + 6);
				cell7.setCellValue(formatter.format(memberList.get(count).getamountOwed()));
				cell7.setCellStyle(cellStyle);
				
				row++;
			}
		}
		
		numOfActiveMembers = (memberList.size()) - numOfNonactiveMembers;
		
		// Creates footer
		Footer footer = sheet.getFooter();
		footer.setCenter("Nonactive Members: " + numOfNonactiveMembers + ", Active Members: " + numOfActiveMembers + ", Members owing dues: " + numOfMembersOwing + ", Total Amount Owed: " + formatter.format(totalAmountOwed));
		
		/* Generates an error message in the command prompt and closes the application if
		 * unable to create the report */
		try
		{
			FileOutputStream output;
			
			if(System.getProperty("os.name").startsWith("Mac OS X"))
				output = new FileOutputStream(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/report.xls");
			else
				output = new FileOutputStream("c:\\FBLA-PBL Membership Report Generator Program Files\\report.xls");
			
			workbook.write(output);
			output.close();
		}
		catch(Exception exception)
		{
			System.err.println("Error: Unable to create the report!");
			System.exit(0);
		}
		
		JOptionPane.showMessageDialog(null, "Report generated! What would you like to do now?", "Success", JOptionPane.INFORMATION_MESSAGE);
	}
	
	public static void generateSeniorEmailReport(ArrayList<Member> memberList)
	{
		// Creates Microsoft Excel workbook
		@SuppressWarnings("resource")
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("Senior Email Report");
		
		// Aligns content to the center of each cell
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
				
		/* Sets the page orientation to landscape, limits each page
		 * to 50 lines, and sets the default width of each cell */
		sheet.getPrintSetup().setLandscape(true);
		sheet.setRowBreak(49);
		sheet.setColumnBreak(3);
		sheet.setDefaultColumnWidth(20);
		
		// Creates header cells
		Row headerRow = sheet.createRow(0);
		Cell headerCell = headerRow.createCell(0);
		headerCell.setCellValue("State");
		headerCell.setCellStyle(cellStyle);
		Cell headerCell1 = headerRow.createCell(1);
		headerCell1.setCellValue("First Name");
		headerCell1.setCellStyle(cellStyle);
		Cell headerCell2 = headerRow.createCell(2);
		headerCell2.setCellValue("Last Name");
		headerCell2.setCellStyle(cellStyle);
		Cell headerCell3 = headerRow.createCell(3);
		headerCell3.setCellValue("Email");
		headerCell3.setCellStyle(cellStyle);
		
		/* Tallys the number of senior members, and adds the member's
		 * information to the report if they are a senior */
		int row = 1;
		
		for(int count = 0; count < memberList.size(); count++)
		{
			if(memberList.get(count).getgrade() == 12)
			{
				int cell = 0;
				Row row1 = sheet.createRow(row);
				
				Cell cell1 = row1.createCell(cell);
				cell1.setCellValue(memberList.get(count).getstate());
				cell1.setCellStyle(cellStyle);
				Cell cell2 = row1.createCell(cell + 1);
				cell2.setCellValue(memberList.get(count).getfirstName());
				cell2.setCellStyle(cellStyle);
				Cell cell3 = row1.createCell(cell + 2);
				cell3.setCellValue(memberList.get(count).getlastName());
				cell3.setCellStyle(cellStyle);
				Cell cell4 = row1.createCell(cell + 3);
				cell4.setCellValue(memberList.get(count).getemail());
				cell4.setCellStyle(cellStyle);
				
				row++;
			}
		}
		
		/* Generates an error message in the command prompt and closes the application if
		 * unable to create the report */
		try
		{
			FileOutputStream output = null;
			
			if(System.getProperty("os.name").startsWith("Mac OS X"))
				output = new FileOutputStream(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/report.xls");
			else
				output = new FileOutputStream("c:\\FBLA-PBL Membership Report Generator Program Files\\report.xls");
			
			workbook.write(output);
			output.close();
		}
		catch(Exception exception)
		{
			System.err.println("Error: Unable to create the report!");
			System.exit(0);
		}
		
		JOptionPane.showMessageDialog(null, "Report generated! What would you like to do now?", "Success", JOptionPane.INFORMATION_MESSAGE);
	}
	
	public static ArrayList<Member> sortMembers(ArrayList<Member> memberList)
	{
		// Sorts the array list of member objects by state and alphabetically ignoring case
		Collections.sort(memberList, new Comparator<Member>()
		{
			public int compare(Member member1, Member member2)
			{
				return member1.getstate().compareToIgnoreCase(member2.getstate());
			}
		});
		
		return memberList;
	}
	
	public static void setResources()
	{
		// Creates a directory on the user's system to contain editable program files
		File directory;
				
		if(System.getProperty("os.name").startsWith("Mac OS X"))
			directory = new File(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files");
		else
			directory = new File("c:\\FBLA-PBL Membership Report Generator Program Files");
				
		directory.mkdir();
				
		// Describes the location of the master file
		File masterFile;
				
		if(System.getProperty("os.name").startsWith("Mac OS X"))
			masterFile = new File(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/masterFile.xls");
		else
			masterFile = new File("c:\\FBLA-PBL Membership Report Generator Program Files\\masterFile.xls");
				
		// Creates a new, empty master file if one does not currently exist
		if(!masterFile.exists())
		{
			// Creates Microsoft Excel workbook
			@SuppressWarnings("resource")
			Workbook workbook = new HSSFWorkbook();
			@SuppressWarnings("unused")
			Sheet sheet = workbook.createSheet("Master File");
					
			/* Generates an error message in the command prompt and closes the application if
			 * unable to create the master file */
			try
			{
				FileOutputStream output = null;
						
				if(System.getProperty("os.name").startsWith("Mac OS X"))
					output = new FileOutputStream(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/masterFile.xls");
				else
					output = new FileOutputStream("c:\\FBLA-PBL Membership Report Generator Program Files\\masterFile.xls");
						
				workbook.write(output);
				output.close();
			}
			catch(Exception exception)
			{
				System.err.println("Error: Unable to create the master file!");
				System.exit(0);
			}
		}
	}
	
	public static void viewReport()
	{
		/* Generates an error message in the command prompt and closes the application if
		 * unable to view the report */
		try
		{
			if(System.getProperty("os.name").startsWith("Mac OS X"))
				Desktop.getDesktop().open(new File(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/report.xls"));
			else
				Desktop.getDesktop().open(new File("c:\\FBLA-PBL Membership Report Generator Program Files\\report.xls"));
		}
		catch(IOException exception)
		{
			System.err.println("Error: Unable to view the report!");
			System.exit(0);
		}
	}
	
	public static void printReport()
	{
		/* Generates an error message in the command prompt and closes the application if
		 * unable to print the report */
		try
        {
			if(System.getProperty("os.name").startsWith("Mac OS X"))
				Desktop.getDesktop().print(new File(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/report.xls"));
			else
				Desktop.getDesktop().print(new File("c:\\FBLA-PBL Membership Report Generator Program Files\\report.xls"));
		}
		catch(IOException exception)
		{
			System.err.println("Error: Unable to print the report!");
			System.exit(0);
		}
	}
	
	public static void exportReport()
	{
		// Allows the user to choose a destination for the exported report
		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setDialogTitle("Choose an export destination");
		int userSelection = fileChooser.showSaveDialog(null);
		if(userSelection == JFileChooser.APPROVE_OPTION)
		{
			File source;
			
			if(System.getProperty("os.name").startsWith("Mac OS X"))
				source = new File(System.getProperty("user.home") + "/FBLA-PBL Membership Report Generator Program Files/report.xls");
			else
				source = new File("c:\\FBLA-PBL Membership Report Generator Program Files\\report.xls");
			
			File destination = new File(fileChooser.getCurrentDirectory() + "/" + fileChooser.getSelectedFile().getName() + ".xls");
			
			/* Generates an error message in the command prompt and closes the application if
			 * unable to export the report */
			try
			{
				FileUtils.copyFile(source, destination);
			}
			catch(IOException exception)
			{
				System.err.println("Error: Unable to export the report!");
			    System.exit(0);
			}
		}
	}
}