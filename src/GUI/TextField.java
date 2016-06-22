package GUI;

//File: TextField.java
//Programmer: Garrett Erven
//Date: May 27th, 2016
//Purpose: This class serves as a constructor for text fields.

import javax.swing.JTextField;

@SuppressWarnings("serial")
public class TextField extends JTextField
{
	// Constructor for text fields that sets the number of columns
	public TextField()
	{
		setColumns(10);
	}
}