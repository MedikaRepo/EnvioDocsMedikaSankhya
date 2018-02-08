package br.com.sankhya;

import javax.mail.MessagingException;

import br.com.sankhya.extensions.actionbutton.AcaoRotinaJava;
import br.com.sankhya.extensions.actionbutton.ContextoAcao;
import br.com.sankhya.extensions.actionbutton.Registro;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
//testeGitbb
public class MainClass implements AcaoRotinaJava 
{

	static StringBuffer mensagem = new StringBuffer();
	static String nomeVend="", ramalVend="", emailBody="";

	public static void getPDFData(byte[] arquivo, String nome) 
	{

		byte[] fileBytes;
		try
		{
			fileBytes = arquivo;
			OutputStream targetFile=  new FileOutputStream(
					"/home/mgeweb/modelos/relatorios/propostadevenda/"+nome+".pdf");
			targetFile.write(fileBytes);
			targetFile.close();

		} catch (Exception e) {
			e.printStackTrace();
			mensagem.append("Erro ao gerar anexo PDF. "+e.getMessage()+"\n");
		}
	}	

	public static Object[] ExecutaComandoNoBanco(String sql, String op)
	{
		try
		{  Object[] objRetorno=new Object[10];
		int cont=0; 
		Statement smnt = ConnectMSSQLServer.conn.createStatement(); 

		if(op=="select")
		{
			smnt.execute(sql);
			ResultSet result = smnt.executeQuery(sql); 

			while(result.next())
			{
				objRetorno[cont]=result.getObject(1);
				cont++;
			}
			return objRetorno;
		}
		else if(op=="alter")
		{
			smnt.executeUpdate(sql);
			objRetorno[0]=(Object)1;
			return objRetorno;
		}
		else
		{
			return null;
		}
		}
		catch(SQLException ex)
		{
			System.err.println("SQLException: " + ex.getMessage());
			mensagem.append("Erro ao obter campo SQL("+ex.getMessage()+") \n");
			return null;
		}
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	//public static void main(String[] args) 
	@Override
	public void doAction(ContextoAcao contexto) throws Exception
	{
		MailJava mj = new MailJava();
		String emailVend="";
		java.sql.Timestamp dhEnvio = null;
		@SuppressWarnings("unused")
		BigDecimal codParc=BigDecimal.ZERO, codEnvioDocsMedika=BigDecimal.ZERO;

		mensagem.setLength(0);

		//Conecta no banco do Sankhya
		ConnectMSSQLServer.dbConnect("jdbc:sqlserver://192.168.0.5:1433;DatabaseName=SANKHYA_PROD;", "adriano","Compiles23");

		//recupera o numero da negociação
		try
		{
			Registro[] linha = contexto.getLinhas();

			for (int i = 0; i < linha.length; i++) 
			{
				codEnvioDocsMedika= (BigDecimal)linha[i].getCampo("CODENVIODOCSMEDIKA");
				codParc=(BigDecimal) linha[i].getCampo("CODPARC");
				emailBody=(String) linha[i].getCampo("EMAILBODY");
				dhEnvio=(java.sql.Timestamp)linha[i].getCampo("DHENVIO");
			}
		}catch(Exception e)
		{
			mensagem.append("Erro ao obter campos sankhya. "+e.getMessage());
			contexto.setMensagemRetorno(mensagem.toString());
		}

		try
		{
			//Recupera o ramal do vendedor
			if(ExecutaComandoNoBanco("SELECT AD_RAMAL FROM TSIUSU "
					+ "WHERE CODUSU="+contexto.getUsuarioLogado().toString(), "select")!=null)
			{
				ramalVend=ExecutaComandoNoBanco("SELECT AD_RAMAL FROM TSIUSU "
						+ "WHERE CODUSU="+contexto.getUsuarioLogado().toString(), "select")[0].toString();
			}
		}catch(Exception e)
		{
			mensagem.append("Erro ao obter ramal do vendedor. "+e.getMessage()+"\n");
		}

		try
		{
			//Recupera o email do vendedor
			if(ExecutaComandoNoBanco("SELECT EMAIL FROM TSIUSU WHERE CODUSU="
					+contexto.getUsuarioLogado().toString(), "select")!=null)
			{
				emailVend=ExecutaComandoNoBanco("SELECT EMAIL FROM TSIUSU WHERE CODUSU="
						+contexto.getUsuarioLogado().toString(), "select")[0].toString();
			}
		}catch(Exception e)
		{
			mensagem.append("Erro ao obter email do vendedor. "+e.getMessage());
		}

		try
		{
			//Recupera o nome do vendedor
			if(ExecutaComandoNoBanco("SELECT FUN.NOMEFUNC FROM TGFVEN VEN"+
					" INNER JOIN TSIUSU USU ON USU.CODVEND = VEN.CODVEND"+
					" INNER JOIN TFPFUN FUN ON FUN.CODFUNC = USU.CODFUNC AND FUN.CODEMP=3"+
					" WHERE  USU.CODUSU="+contexto.getUsuarioLogado().toString(), "select")!=null)
			{
				nomeVend=ExecutaComandoNoBanco("SELECT FUN.NOMEFUNC FROM TGFVEN VEN"+
						" INNER JOIN TSIUSU USU ON USU.CODVEND = VEN.CODVEND"+
						" INNER JOIN TFPFUN FUN ON FUN.CODFUNC = USU.CODFUNC AND FUN.CODEMP=3"+
						" WHERE  USU.CODUSU="+contexto.getUsuarioLogado().toString(), "select")[0].toString();
			}
		}catch(Exception e)
		{
			mensagem.append("Erro ao obter nome do vendedor. "+e.getMessage()+"\n");
		}

		try
		{
		
		  dhEnvio.toString();
		  mensagem.append("Envio já executado! Faça um novo cadastro de envio.");
		  contexto.setMensagemRetorno(mensagem.toString());
		  
		}catch (Exception e)
		{
			//configuracoes de envio
			mj.setSmtpHostMail("email-ssl.com.br");
			mj.setSmtpPortMail("465");
			mj.setSmtpAuth("true");
			mj.setSmtpStarttls("true");
			mj.setUserMail(emailVend);
			mj.setFromNameMail("Equipe de Vendas Medika");
			mj.setSmtpAuth("adrianomedika");
			mj.setPassMail("adrianomedika");
			mj.setCharsetMail("ISO-8859-1");
			mj.setSubjectMail("Documentação Medika.");
			mj.setBodyMail(htmlMessage());
			mj.setTypeTextMail(MailJava.TYPE_TEXT_HTML);

			//sete quantos destinatarios desejar
			Object[] destinatarios=new Object[10];
			try
			{
				if(ExecutaComandoNoBanco("SELECT CTT.EMAIL FROM AD_EMAILSENVIODOCSMEDIKA ADE "
						+ "INNER JOIN TGFCTT CTT ON (CTT.CODPARC=ADE.CODPARC AND CTT.CODCONTATO=ADE.CODCONTATO) "
						+ " WHERE CODENVIODOCSMEDIKA="+codEnvioDocsMedika+" AND ADE.CODPARC="+codParc, "select")!=null)
				{
					destinatarios=ExecutaComandoNoBanco("SELECT CTT.EMAIL FROM AD_EMAILSENVIODOCSMEDIKA ADE "
							+ "INNER JOIN TGFCTT CTT ON (CTT.CODPARC=ADE.CODPARC AND CTT.CODCONTATO=ADE.CODCONTATO) "
							+ " WHERE CODENVIODOCSMEDIKA="+codEnvioDocsMedika+" AND ADE.CODPARC="+codParc, "select");	
				}
			}catch(Exception e1)
			{
				mensagem.append("Erro ao obter destinatários. "+e1.getMessage().toString()+"\n");
			}

			String[] stDestinatarios=new String[10];
			for(int i=0;i<destinatarios.length;i++)
			{
				if(destinatarios[i]!=null)
				{
					stDestinatarios[i]=destinatarios[i].toString().trim();
				}
			}
		
			mj.setToMailsUsers(stDestinatarios);

			Object[] anexos = new Object[10];
			Object[] nomes=new Object[10];
			//seta quatos anexos desejar
			List files = new ArrayList();

			try
			{
				//Seleciona os arquivos para anexar
				if(ExecutaComandoNoBanco("SELECT ADM.ARQUIVO FROM AD_ENVIODOCSMEDIKA ADE"
						+ " INNER JOIN AD_ANEXOSENVIODOCSMEDIKA ADA "
						+ " ON ADA.CODENVIODOCSMEDIKA=ADE.CODENVIODOCSMEDIKA"
						+ " INNER JOIN AD_DOCSMEDIKA ADM ON ADM.CODDOCSMEDIKA=ADA.CODDOCSMEDIKA"
						+ " WHERE ADE.CODENVIODOCSMEDIKA="+codEnvioDocsMedika+" AND ADE.CODPARC="+codParc, "select")!=null)
				{
					anexos=ExecutaComandoNoBanco("SELECT ADM.ARQUIVO FROM AD_ENVIODOCSMEDIKA ADE"
							+ " INNER JOIN AD_ANEXOSENVIODOCSMEDIKA ADA "
							+ " ON ADA.CODENVIODOCSMEDIKA=ADE.CODENVIODOCSMEDIKA"
							+ " INNER JOIN AD_DOCSMEDIKA ADM ON ADM.CODDOCSMEDIKA=ADA.CODDOCSMEDIKA"
							+ " WHERE ADE.CODENVIODOCSMEDIKA="+codEnvioDocsMedika+" AND ADE.CODPARC="+codParc, "select");
				}

				//Seleciona os nomes dos arquivos para os anexos
				if(ExecutaComandoNoBanco("SELECT ADM.DESCRICAO FROM AD_ENVIODOCSMEDIKA ADE"
						+ " INNER JOIN AD_ANEXOSENVIODOCSMEDIKA ADA "
						+ " ON ADA.CODENVIODOCSMEDIKA=ADE.CODENVIODOCSMEDIKA"
						+ " INNER JOIN AD_DOCSMEDIKA ADM ON ADM.CODDOCSMEDIKA=ADA.CODDOCSMEDIKA"
						+ " WHERE ADE.CODENVIODOCSMEDIKA="+codEnvioDocsMedika+" AND ADE.CODPARC="+codParc, "select")!=null)
				{
					nomes=ExecutaComandoNoBanco("SELECT ADM.DESCRICAO FROM AD_ENVIODOCSMEDIKA ADE"
							+ " INNER JOIN AD_ANEXOSENVIODOCSMEDIKA ADA "
							+ " ON ADA.CODENVIODOCSMEDIKA=ADE.CODENVIODOCSMEDIKA"
							+ " INNER JOIN AD_DOCSMEDIKA ADM ON ADM.CODDOCSMEDIKA=ADA.CODDOCSMEDIKA"
							+ " WHERE ADE.CODENVIODOCSMEDIKA="+codEnvioDocsMedika+" AND ADE.CODPARC="+codParc, "select");	
				}


				for(int i=0; i<anexos.length-1;i++)
				{
					if (anexos[i]!=null)
					{
						try
						{
							getPDFData((byte[])anexos[i], nomes[i].toString());
							files.add("/home/mgeweb/modelos/relatorios/propostadevenda/"+nomes[i].toString()+".pdf");
						}catch(Exception e1)
						{
							mensagem.append("Erro ao gerar/adicionar anexos. "+e1.getMessage()+"\n");
						}
					}
				}

				mj.setFileMails(files);

			}catch(Exception e1)
			{
				mensagem.append("Erro ao obter anexos e nomes. "+e1.getMessage().toString()+"\n");
			}

			try {
				new MailJavaSender().senderMail(mj);

				mensagem.append("Email enviado com sucesso!");
				
				contexto.setMensagemRetorno(mensagem.toString());

				ExecutaComandoNoBanco("UPDATE AD_ENVIODOCSMEDIKA SET DHENVIO=GETDATE() WHERE CODENVIODOCSMEDIKA="+codEnvioDocsMedika+" AND CODPARC="+codParc, "alter");

			} catch (UnsupportedEncodingException e1)
			{
				e1.printStackTrace();
				mensagem.append("Erro Usupported Enconding ao tentar enviar mensagem. "+e1.getMessage()+"\n");
			} catch (MessagingException e1) 
			{
				e1.printStackTrace();
				mensagem.append("Erro MessagingException ao tentar enviar mensagem. "+e1.getMessage()+"\n");
			}
		}
	}

	private static String htmlMessage() {
		return
				"<html><head> <meta charset="+"\"UTF-8"+"\"></head>"
				+ "<body style="+"\"font-famaly: arial; font-size:12px;"+"\">Prezado(s),<br/><br/>"+		             
				"Segue a documentação solicitada.<br/><br/>"+

				emailBody+"<br/><br/>"+

				"Atenciosamente,"+
				"<br/><br/>"+nomeVend+
				" - Tel:(31) 3688-1901 Ramal:"+ramalVend+" - Equipe de Vendas"+
				"<br><br><HR WIDTH=100% style="+"\"border:1px solid #191970;"+
				"\"><img src="+"\"https://static.wixstatic.com/media/e2601a_be5e1a3b59244509bd59709b1d78733c~mv2.png/v1/fill/w_251,h_104,al_c,usm_0.66_1.00_0.01/e2601a_be5e1a3b59244509bd59709b1d78733c~mv2.png"+
				"\"><br><br>Medika, qualidade em saúde. - <a href="+"\"http://www.medika.com.br"+
				"\">www.medika.com.br</a><br>"+
				"<HR WIDTH=100% style="+"\"border:1px solid #191970;"+"\">"+
				"</body></html>";
	}
}
