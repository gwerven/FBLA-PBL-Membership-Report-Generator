package App;

/* File: Member.java
 * Programmer: Garrett Erven (Student at Clover High School)
 * Date: May 27th, 2016
 * Purpose: The purpose of this program is to maintain current member data and to output reports of
 * member statistics for the national office of Future Business Leaders of America Phi-Beta-Lambda. */

public class Member
{
	// Instance data for the attributes of each member
	private double membershipNumber;
	private String firstName;
	private String lastName;
	private double grade;
	private String school;
	private String state;
	private String email;
	private double yearJoined;
	private String status;
	private double amountOwed;
	
	// Constructor for member objects
	public Member(double memNum, String fName, String lName, double grad, String sch, String sta, String ema, double yJoined, String stat, double aOwed)
	{
		// Initializes the private instance data of each member object
		membershipNumber = memNum;
		firstName = fName;
		lastName = lName;
		grade = grad;
		school = sch;
		state = sta;
		email = ema;
		yearJoined = yJoined;
		status = stat;
		amountOwed = aOwed;
	}
	
	// Accessor methods for all member data
	public double getmembershipNumber()
	{
		return membershipNumber;
	}
	
	public String getfirstName()
	{
		return firstName;
	}
	
	public String getlastName()
	{
		return lastName;
	}
	
	public double getgrade()
	{
		return grade;
	}
	
	public String getschool()
	{
		return school;
	}
	
	public String getstate()
	{
		return state;
	}
	
	public String getemail()
	{
		return email;
	}
	
	public double getyearJoined()
	{
		return yearJoined;
	}
	
	public String getstatus()
	{
		return status;
	}
	
	public double getamountOwed()
	{
		return amountOwed;
	}
}