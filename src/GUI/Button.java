package GUI;

// File: Button.java
// Programmer: Garrett Erven
// Date: May 27th, 2016
// Purpose: This class serves as a constructor for buttons.

import javax.swing.JButton;
import java.awt.Dimension;
import javax.swing.ImageIcon;

@SuppressWarnings("serial")
public class Button extends JButton
{
	// Constructor for buttons that sets the size, icon image, and tool tip text
	public Button(String filePath, String toolTipText)
	{
		setPreferredSize(new Dimension(180, 150));
		setIcon(new ImageIcon(filePath));
		setToolTipText(toolTipText);
	}
}