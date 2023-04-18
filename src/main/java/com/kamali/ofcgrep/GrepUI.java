package com.kamali.ofcgrep;

import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.MessageFormat;
import java.util.List;
import java.util.Properties;
import java.util.regex.Pattern;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JMenuItem;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileFilter;
import javax.swing.table.DefaultTableModel;

public class GrepUI extends JPanel implements PropertyChangeListener,GrepTask.ProcessNotify {
	private class ExcelFilter extends FileFilter implements java.io.FileFilter {
		private String pattern = null;
		public void setPattern(String regex) {
			if (regex == null || regex.isEmpty()) {
				pattern = null;
			} else {
				pattern = regex;
			}
		}
		public String getPattern() { return pattern; }
		@Override
		public boolean accept(File file) {
			if (file.isDirectory()) {
				if (file.getName().equals(".")) return false;
				if (file.getName().equals("..")) return false;
				if (file.getName().startsWith("::")) return false;
				return true;
			}
			if (Pattern.matches(".+\\.[xX][lL][sS].*$", file.getName())) {
				if (file.getName().startsWith("~$")) return false;
				if (pattern == null) return true;
				if (file.getName().matches(pattern)) return true;
			}
			return false;
		}

		@Override
		public String getDescription() {
			return "Excel Files (*.xls*)";
		}

	}
	private final ExcelFilter excelFilter = new ExcelFilter();
	private final ExcelFilter fileFilter = new ExcelFilter();
	private final JFileChooser fc = new JFileChooser();
	private static final MessageFormat msgFmt = new MessageFormat("{0,number} / {1,number} {2}");
	private static final Properties props = new Properties();
	static {
		InputStream is = null;
		try {
			is = new FileInputStream("./ofcgrep.cfg");
			props.load(is);
		} catch (FileNotFoundException e) {
			;
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {}
			}
		}
	}
	private int oldNew = 0;
	private File lastDir = new File("./");
	private File selectedDir = null;
	private String strPattern = null;
	private String shPattern = null;
	private JTextField txtFolder;
	private JTextField txtFilePattern;
	private JTextField txtStrPattern;
	private JTextField txtShPattern;
	private JTable table;
	private JProgressBar progressBar;
	private JCheckBox chkCase;
	private JCheckBox chkMultiline;
	private JPopupMenu popupMenu;

	/**
	 * Create the panel.
	 */
	public GrepUI() {
		if (props.containsKey("lastDir")) {
			lastDir = new File(props.getProperty("lastDir"));
		}
		setLayout(new BorderLayout(0, 0));

		JPanel pnlInput = new JPanel();
		add(pnlInput, BorderLayout.NORTH);
		GridBagLayout gbl_pnlInput = new GridBagLayout();
		gbl_pnlInput.columnWidths = new int[] {10, 500, 1};
		gbl_pnlInput.rowHeights = new int[] {10, 10, 10, 10, 10};
		gbl_pnlInput.columnWeights = new double[]{0.0, 200.0, 1.0};
		gbl_pnlInput.rowWeights = new double[]{0.0, 0.0, 0.0, 0.0, 1.0};
		pnlInput.setLayout(gbl_pnlInput);

		JLabel lblFolder = new JLabel("Folder");
		GridBagConstraints gbc_lblFolder = new GridBagConstraints();
		gbc_lblFolder.insets = new Insets(0, 0, 5, 5);
		gbc_lblFolder.anchor = GridBagConstraints.EAST;
		gbc_lblFolder.gridx = 0;
		gbc_lblFolder.gridy = 0;
		pnlInput.add(lblFolder, gbc_lblFolder);

		txtFolder = new JTextField();
		lblFolder.setLabelFor(txtFolder);
		GridBagConstraints gbc_txtFolder = new GridBagConstraints();
		gbc_txtFolder.fill = GridBagConstraints.BOTH;
		gbc_txtFolder.insets = new Insets(0, 0, 5, 5);
		gbc_txtFolder.gridx = 1;
		gbc_txtFolder.gridy = 0;
		pnlInput.add(txtFolder, gbc_txtFolder);
		txtFolder.setColumns(10);

		JButton btnRef = new JButton("...");
		btnRef.setPreferredSize(new Dimension(36, 20));
		btnRef.setMinimumSize(new Dimension(36, 20));
		btnRef.setMaximumSize(new Dimension(36, 20));
		btnRef.setMargin(new Insets(0, 12, 0, 12));
		btnRef.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {
				selectFolder();
			}
		});
		GridBagConstraints gbc_btnRef = new GridBagConstraints();
		gbc_btnRef.insets = new Insets(0, 0, 5, 1);
		gbc_btnRef.gridx = 2;
		gbc_btnRef.gridy = 0;
		pnlInput.add(btnRef, gbc_btnRef);

		JLabel lblFilePattern = new JLabel("File Pattern");
		GridBagConstraints gbc_lblFilePattern = new GridBagConstraints();
		gbc_lblFilePattern.insets = new Insets(0, 0, 5, 5);
		gbc_lblFilePattern.anchor = GridBagConstraints.EAST;
		gbc_lblFilePattern.gridx = 0;
		gbc_lblFilePattern.gridy = 1;
		pnlInput.add(lblFilePattern, gbc_lblFilePattern);

		txtFilePattern = new JTextField();
		lblFilePattern.setLabelFor(txtFilePattern);
		GridBagConstraints gbc_txtFilePattern = new GridBagConstraints();
		gbc_txtFilePattern.gridwidth = 2;
		gbc_txtFilePattern.insets = new Insets(0, 0, 5, 1);
		gbc_txtFilePattern.fill = GridBagConstraints.HORIZONTAL;
		gbc_txtFilePattern.gridx = 1;
		gbc_txtFilePattern.gridy = 1;
		pnlInput.add(txtFilePattern, gbc_txtFilePattern);
		txtFilePattern.setColumns(10);

		txtShPattern = new JTextField();
		GridBagConstraints gbc_txtShPattern = new GridBagConstraints();
		gbc_txtShPattern.gridwidth = 2;
		gbc_txtShPattern.insets = new Insets(0, 0, 5, 1);
		gbc_txtShPattern.fill = GridBagConstraints.HORIZONTAL;
		gbc_txtShPattern.gridx = 1;
		gbc_txtShPattern.gridy = 2;
		pnlInput.add(txtShPattern, gbc_txtShPattern);
		txtShPattern.setColumns(10);

		JLabel lblShPattern = new JLabel("Sheet Pattern");
		GridBagConstraints gbc_lblShPattern = new GridBagConstraints();
		gbc_lblShPattern.insets = new Insets(0, 0, 5, 5);
		gbc_lblShPattern.gridx = 0;
		gbc_lblShPattern.gridy = 2;
		pnlInput.add(lblShPattern, gbc_lblShPattern);

		JLabel lblStrPattern = new JLabel("String Pattern");
		GridBagConstraints gbc_lblStrPattern = new GridBagConstraints();
		gbc_lblStrPattern.insets = new Insets(0, 0, 5, 5);
		gbc_lblStrPattern.anchor = GridBagConstraints.EAST;
		gbc_lblStrPattern.gridx = 0;
		gbc_lblStrPattern.gridy = 3;
		pnlInput.add(lblStrPattern, gbc_lblStrPattern);

		txtStrPattern = new JTextField();
		GridBagConstraints gbc_txtStrPattern = new GridBagConstraints();
		gbc_txtStrPattern.insets = new Insets(0, 0, 5, 1);
		gbc_txtStrPattern.gridwidth = 2;
		gbc_txtStrPattern.fill = GridBagConstraints.HORIZONTAL;
		gbc_txtStrPattern.gridx = 1;
		gbc_txtStrPattern.gridy = 3;
		pnlInput.add(txtStrPattern, gbc_txtStrPattern);
		txtStrPattern.setColumns(10);

		JPanel panel = new JPanel();
		FlowLayout flowLayout = (FlowLayout) panel.getLayout();
		flowLayout.setAlignment(FlowLayout.LEFT);
		flowLayout.setHgap(1);
		flowLayout.setVgap(0);
		GridBagConstraints gbc_panel = new GridBagConstraints();
		gbc_panel.insets = new Insets(0, 0, 0, 5);
		gbc_panel.fill = GridBagConstraints.BOTH;
		gbc_panel.gridx = 1;
		gbc_panel.gridy = 4;
		pnlInput.add(panel, gbc_panel);

		JButton btnFind = new JButton("Find");
		btnFind.setMnemonic('F');
		btnFind.setMnemonic(KeyEvent.VK_F5);
		panel.add(btnFind);
		btnFind.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {
				findText();
			}
		});
		btnFind.setHorizontalAlignment(SwingConstants.LEFT);

		chkCase = new JCheckBox("Case insesitive");
		chkCase.setHorizontalAlignment(SwingConstants.LEFT);
		panel.add(chkCase);

		chkMultiline = new JCheckBox("Multiline");
		panel.add(chkMultiline);

		popupMenu = new JPopupMenu();
		panel.add(popupMenu);
//		popupMenu.addMouseListener(new MouseAdapter() {
//			@Override
//			public void mouseClicked(MouseEvent e) {
//				openDocument(e.getSource());
//			}
//		});

		JMenuItem item1 = new JMenuItem("Open file");
		popupMenu.add(item1);

		JPanel pnlResult = new JPanel();
		add(pnlResult, BorderLayout.CENTER);
		pnlResult.setLayout(new BorderLayout(0, 0));

		JScrollPane scrollPane = new JScrollPane();
		pnlResult.add(scrollPane);

		table = new JTable();
		final Component me = this;
		table.addMouseListener(new MouseAdapter() {
			/* (非 Javadoc)
			 * @see java.awt.event.MouseAdapter#mouseClicked(java.awt.event.MouseEvent)
			 */
			@Override
			public void mouseClicked(MouseEvent e) {
				if (e.getButton() == MouseEvent.BUTTON1 && 1 < e.getClickCount()) {
					int row = table.rowAtPoint(e.getPoint());
					if (0 <= row) {
						DefaultTableModel model = (DefaultTableModel)table.getModel();
						String file = (String)model.getValueAt(row, 0);
						String dir = (String)model.getValueAt(row, 4);
						openDocument(dir, file);
					}
				}
			}

			/* (non Javadoc)
			 * @see java.awt.event.MouseAdapter#mouseReleased(java.awt.event.MouseEvent)
			 */
			@Override
			public void mouseReleased(MouseEvent e) {
				if (e.isPopupTrigger()) {
					int row = table.rowAtPoint(e.getPoint());
					if (0 <= row) {
						table.changeSelection(row, 0, false, false);
						if (e.getButton() == MouseEvent.BUTTON3) {
							showContextMenu(e.getComponent(), e.getX(), e.getY());
						}
					}
				}
			}
		});
		table.setModel(new DefaultTableModel(
			new Object[][] {
				{null, null, "", null, null},
			},
			new String[] {
				"File", "Sheet", "Address", "Value", "Folder"
			}
		) {
			Class[] columnTypes = new Class[] {
				String.class, String.class, String.class, String.class, String.class
			};
			public Class getColumnClass(int columnIndex) {
				return columnTypes[columnIndex];
			}
			boolean[] columnEditables = new boolean[] {
				false, false, false, false, false
			};
			public boolean isCellEditable(int row, int column) {
				return columnEditables[column];
			}
		});
		table.getColumnModel().getColumn(2).setPreferredWidth(55);
		//table.getColumnModel().getColumn(2).setMaxWidth(80);
		table.getColumnModel().getColumn(3).setPreferredWidth(100);
		scrollPane.setViewportView(table);

		progressBar = new JProgressBar();
		progressBar.setMaximumSize(new Dimension(32767, 14));
		progressBar.setMinimumSize(new Dimension(10, 14));
		progressBar.setPreferredSize(new Dimension(150, 14));
		progressBar.setStringPainted(true);
		add(progressBar, BorderLayout.SOUTH);

	}
	private void showContextMenu(Component c, int x, int y) {
		popupMenu.show(c, x, y);
	}
	private void selectFolder() {
		fc.setCurrentDirectory(lastDir);
		fc.setFileFilter(excelFilter);
		fc.setFileHidingEnabled(true);
		fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
		fc.setSelectedFile(lastDir);
		int ret = fc.showOpenDialog(this);
		if (ret != JFileChooser.APPROVE_OPTION) return;
		lastDir = fc.getSelectedFile();
		if (!lastDir.isDirectory()) {
			lastDir = lastDir.getParentFile();
		}
		selectedDir = lastDir;
		txtFolder.setText(selectedDir.getPath());
		props.setProperty("lastDir", lastDir.getPath());
		FileOutputStream os = null;
		try {
			os = new FileOutputStream("./ofcgrep.cfg");
			props.store(os, "");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (os != null) {
				try {
					os.close();
				} catch (IOException e) {}
			}
		}
	}
	private void findText() {
		int flag = chkCase.isSelected() ? (Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE) : 0;
		flag |= (chkMultiline.isSelected() ? (Pattern.MULTILINE | Pattern.DOTALL) : 0);
		String dir = txtFolder.getText();
		if (dir.isEmpty()) dir = "/";
		fileFilter.setPattern(txtFilePattern.getText());
		String ptn = txtStrPattern.getText();
		if (0 < ptn.length()) {
			strPattern = ptn;
		} else {
			strPattern = null;
			return;
		}
		ptn = txtShPattern.getText();
		if (0 < ptn.length()) {
			shPattern = ptn;
		} else {
			shPattern = null;
		}
		GrepTask task = new GrepTask(new File(dir), fileFilter, strPattern, shPattern);
		task.addPropertyChangeListener(this);
		task.setNotify(this);
		task.execute();
	}
	private void openDocument(String dir, String file) {
		Desktop desktop = Desktop.getDesktop();
		try {
			desktop.open(new File(dir, file));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	@Override
	public void propertyChange(PropertyChangeEvent ev) {
		String p = ev.getPropertyName();
		System.out.print(p);
		System.out.print(':');
		if ("progress".equals(p)) {
			int pg = ((GrepTask)ev.getSource()).getProgress();
			System.out.println(pg);
			progressBar.setValue(pg);
		} else if ("detail".equals(p)) {
			System.out.print(((Object[])ev.getNewValue())[0]);
			System.out.print('/');
			System.out.print(((Object[])ev.getNewValue())[1]);
			System.out.print(':');
			System.out.println(((Object[])ev.getNewValue())[2]);
			progressBar.setString(msgFmt.format((Object[])ev.getNewValue()));
		} else if ("state".equals(p)) {
			System.out.println(ev.getNewValue());
			if (SwingWorker.StateValue.STARTED.equals(ev.getNewValue())) {
				((DefaultTableModel)table.getModel()).setRowCount(0);
			} else if (SwingWorker.StateValue.DONE.equals(ev.getNewValue())) {
				progressBar.setString("done");
			}
		}
	}
	@Override
	public void process(List<Object[]> chunks) {
		Object[] rowdata = new Object[5];
		DefaultTableModel model = (DefaultTableModel)table.getModel();
		for (Object[] chunk : chunks) {
			File f = (File)chunk[0];
			rowdata[0] = f.getName();		//file name
			rowdata[1] = (String)chunk[1];	//sheet
			rowdata[2] = (String)chunk[2];	//cell
			rowdata[3] = (String)chunk[3];	//text
			rowdata[4] = f.getParent();		//directory
			model.addRow(rowdata);
		}
	}
	public JProgressBar getProgressBar() {
		return progressBar;
	}
}
