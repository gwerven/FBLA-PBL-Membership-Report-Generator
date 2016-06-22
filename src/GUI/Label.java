package GUI;

//File: Label.java
//Programmer: Garrett Erven
//Date: May 27th, 2016
//Purpose: This class serves as a constructor for labels.

import javax.swing.JLabel;
import java.awt.Color;
import java.awt.Font;
import javax.swing.ImageIcon;

@SuppressWarnings("serial")
public class Label extends JLabel
{
	// Constructor for labels that sets the icon image
	public Label(String filePath)
	{
		super(new ImageIcon(filePath));
	}
	
	// Constructor for labels that sets the text, font, color, and tool tip text
	public Label(String text, String toolTipText)
	{
		setText(text);
		setFont(new Font("Arial", Font.PLAIN, 18));
		setForeground(Color.blue);
		setToolTipText(toolTipText);
	}
}