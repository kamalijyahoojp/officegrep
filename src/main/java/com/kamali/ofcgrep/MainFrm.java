package com.kamali.ofcgrep;

import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.KeyStroke;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;

public class MainFrm extends JFrame {

	/** auto generated serialVUID */
	private static final long serialVersionUID = -18368528029303945L;
	private int tabNo = 1;
	private JPanel contentPane;
	private JTabbedPane tabbedPane;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					MainFrm frame = new MainFrm();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}
	private void initComponent() {
		tabbedPane.addTab("Find1", new GrepUI());
	}
	/**
	 * Create the frame.
	 */
	public MainFrm() {
		addWindowListener(new WindowAdapter() {
			@Override
			public void windowOpened(WindowEvent e) {
				initComponent();
			}
		});
		setTitle("Office Grep");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 580, 500);

		JMenuBar menuBar = new JMenuBar();
		setJMenuBar(menuBar);

		JMenu mnFile = new JMenu("File");
		mnFile.setMnemonic('F');
		menuBar.add(mnFile);

		JMenuItem mntmExit = new JMenuItem("eXit");
		mntmExit.setMnemonic('X');
		mntmExit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				onExit();
			}
		});
		mnFile.add(mntmExit);

		JMenuItem mitAdd = new JMenuItem("Add Tab");
		mitAdd.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {
				tabbedPane.addTab("Find" + String.valueOf(++tabNo), new GrepUI());
			}
		});
		mitAdd.setHorizontalAlignment(SwingConstants.LEFT);
		mitAdd.setHorizontalTextPosition(SwingConstants.LEFT);
		mitAdd.setMaximumSize(new Dimension(90, 32767));
		mitAdd.setContentAreaFilled(false);
		mitAdd.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_F2, 0));
		menuBar.add(mitAdd);

		JMenuItem mitDel = new JMenuItem("Delete Tab");
		mitDel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {
				int idx = tabbedPane.getSelectedIndex();
				if (0 <= idx)
					tabbedPane.remove(idx);
			}
		});
		mitDel.setHorizontalTextPosition(SwingConstants.LEFT);
		mitDel.setHorizontalAlignment(SwingConstants.LEFT);
		mitDel.setMaximumSize(new Dimension(100, 32767));
		mitDel.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_F3, 0));
		menuBar.add(mitDel);

		JMenuItem mitAbout = new JMenuItem("About");
		mitAbout.setMaximumSize(new Dimension(70, 32767));
		mitAbout.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				onAbout();
			}
		});
		mitAbout.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_F1, 0));
		menuBar.add(mitAbout);

		JMenu mnDummy = new JMenu(" ");
		mnDummy.setAlignmentY(Component.BOTTOM_ALIGNMENT);
		mnDummy.setAlignmentX(Component.LEFT_ALIGNMENT);
		mnDummy.setMinimumSize(new Dimension(100, 17));
		mnDummy.setPreferredSize(new Dimension(100, 17));
		mnDummy.setEnabled(false);
		menuBar.add(mnDummy);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		contentPane.setLayout(new BorderLayout(0, 0));
		setContentPane(contentPane);

		tabbedPane = new JTabbedPane(JTabbedPane.TOP);
		tabbedPane.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent arg0) {
			}
		});
		contentPane.add(tabbedPane, BorderLayout.CENTER);
	}

	protected JTabbedPane getTabbedPane() {
		return tabbedPane;
	}

	protected void onExit() {
		this.dispose();
	}

	protected void onAbout() {
		JOptionPane.showMessageDialog(this,
				  "Grep tool for MS-Office\n(C)2020 Kamali(J)\n"
				+ "This application uses \"Apache POI\""
				, "About this application", JOptionPane.OK_OPTION | JOptionPane.PLAIN_MESSAGE);
	}
}
