package br.com.sankhya;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.PasswordAuthentication;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Multipart;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.UnsupportedEncodingException;
import java.util.Map;
import java.util.Properties;

public class MailJavaSender {

	//cria as propriedades necessarias para o envio de email
	public void senderMail(final MailJava mail) throws
	UnsupportedEncodingException, MessagingException {

		Properties props = new Properties();
		props.put("mail.smtp.host", mail.getSmtpHostMail());
		props.put("mail.smtp.socketFactory.port", "465");
		props.put("mail.smtp.socketFactory.class",
				"javax.net.ssl.SSLSocketFactory");
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.port", "465"); 

		//classe anonima que realiza a autenticacao
		//do usuario no servidor de email.
		Authenticator auth = new Authenticator() {
			public PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication("adriano.soares@medika.com.br", "adrianomedika");
			}
		};

		try
		{
			// Cria a sessao passando as propriedades e a autenticacao
			Session session = Session.getInstance(props, auth);
			// Gera um log no console referente ao processo de envio
			//session.setDebug(true);


			try
			{
				//cria a mensagem setando o remetente e seus destinatarios
				Message msg = new MimeMessage(session);
				//aqui seta o remetente
				msg.setFrom(new InternetAddress(
						mail.getUserMail(), mail.getFromNameMail())
						);
				
				//boolean first = true;
				for (int i=0; i<mail.getToMailsUsers().length; i++) 
				{
					if (i==0) {
						//setamos o 1° destinatario
						msg.addRecipient(Message.RecipientType.TO,
								new InternetAddress(mail.getToMailsUsers()[i]));
				
					} else if(mail.getToMailsUsers()[i]!=null)
					{
						//setamos os demais destinatarios
						msg.addRecipient(Message.RecipientType.CC,
								new InternetAddress(mail.getToMailsUsers()[i]));
					}
				}

				// Adiciona um Assunto a Mensagem
				msg.setSubject(mail.getSubjectMail());

				// Cria o objeto que recebe o texto do corpo do email
				MimeBodyPart textPart = new MimeBodyPart();
				textPart.setContent(mail.getBodyMail(), mail.getTypeTextMail());

				// Monta a mensagem SMTP  inserindo o conteudo, texto e anexos
				Multipart mps = new MimeMultipart();
				for (int index = 0; index < mail.getFileMails().size(); index++) {

					// Cria um novo objeto para cada arquivo, e anexa o arquivo
					MimeBodyPart attachFilePart = new MimeBodyPart();
					FileDataSource fds =   new FileDataSource(
							mail.getFileMails().get(index)
							);
					attachFilePart.setDataHandler(new DataHandler(fds));
					attachFilePart.setFileName(fds.getName());

					//adiciona os anexos da mensagem
					mps.addBodyPart(attachFilePart, index);

				}

				//adiciona o corpo texto da mensagem
				mps.addBodyPart(textPart);

				//adiciona a mensagem o conteudo texto e anexo
				msg.setContent(mps);



				try
				{
					// Envia a Mensagem
					Transport.send(msg);
				}catch(Exception e)
				{
					MainClass.mensagem.append("Erro no transport.send. "+e.getMessage());
				}
			}catch(Exception e)
			{
				MainClass.mensagem.append("Erro setando propriedades da mensagem. "+e.getMessage());
			}

		}catch(Exception e)
		{
			MainClass.mensagem.append("Erro obtendo a sessão. "+e.getMessage());
		}
	}
}