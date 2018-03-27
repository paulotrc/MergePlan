package br.com.paulotrc.mergeplan;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.KeyStroke;

public class Principal {
	
	private static ActionListener sairListener;
	private static ActionListener carregarArquivoXlsListener;
	private static ActionListener executarMergeListener;
	
	private static JFrame janela;
	private static JPanel painel; 
	private static JButton botaoCarregarXls;
	private static JButton botaoExecutarMerge;
	private static JTextField caixaPathArquivo; 
	private static File arquivoXlsSelecionado;
	
	public static void main(String[] args) {

		carregaListenersDaAplicacao();
		
		caixaPathArquivo = new JTextField("Caminho do Arquivo...");
		carregaCaracteristicasCaixaDeTexto(caixaPathArquivo);
		
		botaoCarregarXls = new JButton("Carregar Arquivo");
		botaoCarregarXls.addActionListener(carregarArquivoXlsListener);
		
		botaoExecutarMerge = new JButton("Executar Merge");
		botaoExecutarMerge.addActionListener(executarMergeListener);
		
		painel = new JPanel();
		painel.add(caixaPathArquivo);
		painel.add(botaoCarregarXls);
		painel.add(botaoExecutarMerge);
		
		janela = new JFrame("Merge - Plan");
		janela.setJMenuBar(carregaMenu());
		janela.add(painel);
		janela.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		janela.pack();
		janela.setSize(500, 170);
		janela.setVisible(true);
	}

	private static JMenuBar carregaMenu() {
		JMenuItem itemSair = new JMenuItem("Sair", KeyEvent.VK_T);
		itemSair.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_1, ActionEvent.ALT_MASK));
		itemSair.getAccessibleContext().setAccessibleDescription("This doesn't really do anything");
		itemSair.addActionListener(sairListener);
		
		JMenu menuHome = new JMenu("Principal");
		menuHome.add(itemSair);
		
		JMenuBar menuApp = new JMenuBar();
		menuApp.add(menuHome);
		return menuApp;
	}

	private static void carregaCaracteristicasCaixaDeTexto(JTextField caixaPathArquivo2) {
		caixaPathArquivo2.setEnabled(false);
		caixaPathArquivo2.setBackground(new Color(205,205,205));
		caixaPathArquivo2.setPreferredSize(new Dimension(400, 25));
		caixaPathArquivo2.setDisabledTextColor(new Color(0,127,255));
	}

	private static void carregaListenersDaAplicacao() {
		sairListener = new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.exit(0);
			}
		};

		//Carregar arquivo em memória
		carregarArquivoXlsListener = new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				carregarArquivoSelecionado();
			}

			private void carregarArquivoSelecionado() {
				JFileChooser fileChooser = new JFileChooser();
				int retorno = fileChooser.showOpenDialog(null);

				if (retorno == JFileChooser.APPROVE_OPTION) {
				  arquivoXlsSelecionado = fileChooser.getSelectedFile();
				  caixaPathArquivo.setText(arquivoXlsSelecionado.getAbsolutePath());
				} else {
					arquivoXlsSelecionado = null;
					caixaPathArquivo.setText("Caminho do Arquivo...");
				}
				
			}
		};
		
		executarMergeListener = new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if(arquivoXlsSelecionado != null){
					String extensaoArquivo = arquivoXlsSelecionado.getAbsolutePath().split("\\.")[arquivoXlsSelecionado.getAbsolutePath().split("\\.").length - 1];
					ExecutarPlan executar = new ExecutarPlan();
					try {
						executar.ExecutarPlanilha(arquivoXlsSelecionado);
						botaoExecutarMerge.setEnabled(false);
						botaoCarregarXls.setEnabled(false);
						if(extensaoArquivo.toUpperCase().matches("XLS")){
							executar.readXLSFile();
							executar.writeXLSFile();
						}
						if(extensaoArquivo.toUpperCase().matches("XLSX")){
							executar.readXLSFile();
							executar.writeXLSFile();
						}
						JOptionPane.showMessageDialog(null, "Processamento realizado com sucesso!");
						botaoExecutarMerge.setEnabled(true);
						botaoCarregarXls.setEnabled(true);
					} catch (Exception ex) {
						JOptionPane.showMessageDialog(null, "Erro no Processamento!");
						botaoExecutarMerge.setEnabled(true);
						botaoCarregarXls.setEnabled(true);
						ex.printStackTrace();
					}finally {
						executar = null;
					}
				}
			}
		};
		
		
		
	}

}
