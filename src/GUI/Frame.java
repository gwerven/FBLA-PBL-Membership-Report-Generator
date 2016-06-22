package GUI;

// File: Frame.java
// Programmer: Garrett Erven
// Date: May 27th, 2016
// Purpose: This class serves as a constructor for frames.

import javax.swing.JFrame;
import java.awt.Dimension;
import javax.swing.ImageIcon;

@SuppressWarnings("serial")
public class Frame extends JFrame
{
	// Constructor for the application frame that sets the title, icon image, size, location, close operation, and visibility
	public Frame(String iconFilePath)
	{
		setTitle("FBLA-PBL Membership Report Generator");
		setIconImage(new ImageIcon(iconFilePath).getImage());
		setPreferredSize(new Dimension(950, 650));
		setResizable(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}
}