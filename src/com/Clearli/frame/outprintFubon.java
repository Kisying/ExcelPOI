package com.Clearli.frame;
import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JSplitPane;
import javax.swing.JTextField;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.JButton;
import javax.swing.JFileChooser;

import java.awt.event.ActionListener;
import java.io.File;
import java.awt.event.ActionEvent;
import javax.swing.JTextPane;
import com.Clearli.model.*;

import javax.swing.JOptionPane;
public class outprintFubon extends JFrame {

	private JPanel contentPane;
	private JTextField InputAddress;
	private JTextField outputAddress;
	
	static OutPrintExcle OutPrintExcle = new OutPrintExcle();
	private JButton InputAddressBtn;
	private JButton outPutAddressBtn;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					outprintFubon frame = new outprintFubon();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public outprintFubon() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		
		InputAddress = new JTextField();
		InputAddress.setColumns(10);
		
		outputAddress = new JTextField();
		outputAddress.setColumns(10);
		
		JButton StartAction = new JButton("\u958B\u59CB");
		StartAction.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String InputAdd = InputAddress.getText();
				String outputAdd = outputAddress.getText();
				
				try {
					OutPrintExcle.main(InputAdd, outputAdd);
					JOptionPane.showMessageDialog(null,"成功");
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
					JOptionPane.showMessageDialog(null,"失敗");
				}
			}
		});
		
		InputAddressBtn = new JButton("\u4E0A\u50B3\u5730\u5740");
		InputAddressBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				 JFileChooser jfc = new JFileChooser();
				 jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );
				 jfc.showDialog(new JLabel(), "選擇");
				 File file=jfc.getSelectedFile();
				    if(file.isDirectory()){
				     // System.out.println("文件夾:"+file.getAbsolutePath());
				      InputAddress.setText(file.getAbsolutePath());
				    }else if(file.isFile()){
				     // System.out.println("文件:"+file.getAbsolutePath());
				    }
			}
		});
		
		outPutAddressBtn = new JButton("\u8F38\u51FA\u5730\u5740");
		outPutAddressBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser jfc = new JFileChooser();
				 jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );
				 jfc.showDialog(new JLabel(), "選擇");
				 File file=jfc.getSelectedFile();
				    if(file.isDirectory()){
				      //System.out.println("文件夾:"+file.getAbsolutePath());
				      outputAddress.setText(file.getAbsolutePath());
				    }else if(file.isFile()){
				      //System.out.println("文件:"+file.getAbsolutePath());
				    }
			}
		});
		GroupLayout gl_contentPane = new GroupLayout(contentPane);
		gl_contentPane.setHorizontalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addGap(26)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
						.addGroup(Alignment.TRAILING, gl_contentPane.createSequentialGroup()
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
								.addComponent(InputAddressBtn)
								.addComponent(outPutAddressBtn))
							.addPreferredGap(ComponentPlacement.RELATED, 30, Short.MAX_VALUE)
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING, false)
								.addComponent(outputAddress, Alignment.TRAILING)
								.addComponent(InputAddress, Alignment.TRAILING, GroupLayout.DEFAULT_SIZE, 253, Short.MAX_VALUE)))
						.addComponent(StartAction, GroupLayout.DEFAULT_SIZE, 382, Short.MAX_VALUE))
					.addContainerGap())
		);
		gl_contentPane.setVerticalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addContainerGap()
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(InputAddress, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(InputAddressBtn))
					.addPreferredGap(ComponentPlacement.UNRELATED)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(outputAddress, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(outPutAddressBtn))
					.addGap(18)
					.addComponent(StartAction, GroupLayout.PREFERRED_SIZE, 108, GroupLayout.PREFERRED_SIZE)
					.addContainerGap(41, Short.MAX_VALUE))
		);
		contentPane.setLayout(gl_contentPane);
	}
}
