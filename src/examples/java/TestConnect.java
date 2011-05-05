import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import java.util.Properties;

import javax.mail.Authenticator;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.Transport;

import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

public class TestConnect {

    public static void main(String[] args) throws Exception {
        if (args == null || args.length < 1) {
            System.out.println("Usage:  TestConnect <propertyFile>");
            System.out.println();
            System.out.println("Where propertyFile is a properties file " +
                    "containing configuration options.");
            System.exit(0);
        }
        File propFile = new File(args[0]);
        if (!propFile.isFile()) {
            System.err.println("Property file \"" +
                    propFile.getCanonicalPath() + "\" is not available.");
            System.exit(1);
        }
        Properties props = new Properties();
        InputStream input = new FileInputStream(propFile);
        props.load(input);
        input.close();
        testSmtp(props);
        testPop3(props);
    }

    private static void testSmtp(Properties props) {
        Transport transport = null;
        try {
            String host = props.getProperty("mail.smtp.host");
            if (host == null) {
                host = props.getProperty("mail.host");
                if (host == null) {
                    System.err.println("No \"mail.smtp.host\" or " +
                            "\"mail.host\" specified, skipping SMTP test.");
                    return;
                }
            }
            String user = props.getProperty("mail.smtp.user");
            if (user == null) {
                user = props.getProperty("mail.user");
                if (user == null) {
                    System.err.println( "No \"mail.smtp.user\" or " +
                            "\"mail.user\" found, skipping SMTP test.");
                    return;
                }
            }
            String password = props.getProperty("mail.smtp.password");
            if (password == null) {
                password = props.getProperty("mail.password");
                if (user == null) {
                    System.err.println( "No \"mail.smtp.password\" or " +
                            "\"mail.password\" found, skipping SMTP test.");
                    return;
                }
            }
            Session session = Session.getDefaultInstance(props);
            transport = session.getTransport("smtp");
            String transportClass = transport.getClass().getName();
            if (!"org.exjello.mail.ExchangeTransport".equals(transportClass)) {
                System.err.println(
                        "The exJello SMTP provider is NOT installed.");
                System.err.println("Skipping SMTP test.");
                return;
            }
            String email = props.getProperty("mail.smtp.from");
            if (email == null) {
                System.err.println("No \"mail.smtp.from\" " +
                        "specified, skipping SMTP test.");
                return;
            }
            MimeMessage msg = new MimeMessage(session);
            msg.setSubject("Test Message");
            msg.setText("Testing.");
            InternetAddress address = new InternetAddress(email);
            msg.setFrom(address);
            msg.setRecipients(Message.RecipientType.TO, new InternetAddress[] {
                address
            });
            if (password == null) System.exit(3);
            transport.connect(host, user, password);
            transport.sendMessage(msg, new InternetAddress[] { address });
        } catch (Exception ex) {
            System.err.println("Error testing SMTP:");
            ex.printStackTrace();
        } finally {
            if (transport != null) {
                try {
                    transport.close();
                } catch (Exception ignore) { }
            }
        }
    }

    private static void testPop3(Properties props) {
        Store store = null;
        Folder inbox = null;
        try {
            String host = props.getProperty("mail.pop3.host");
            if (host == null) {
                host = props.getProperty("mail.host");
                if (host == null) {
                    System.err.println("No \"mail.pop3.host\" or " +
                            "\"mail.host\" specified, skipping POP3 test.");
                    return;
                }
            }
            String user = props.getProperty("mail.pop3.user");
            if (user == null) {
                user = props.getProperty("mail.user");
                if (user == null) {
                    System.err.println( "No \"mail.pop3.user\" or " +
                            "\"mail.user\" found, skipping POP3 test.");
                    return;
                }
            }
            String password = props.getProperty("mail.pop3.password");
            if (password == null) {
                password = props.getProperty("mail.password");
                if (user == null) {
                    System.err.println( "No \"mail.pop3.password\" or " +
                            "\"mail.password\" found, skipping POP3 test.");
                    return;
                }
            }
            Session session = Session.getDefaultInstance(props);
            store = session.getStore("pop3");
            String storeClass = store.getClass().getName();
            if (!"org.exjello.mail.ExchangeStore".equals(storeClass)) {
                System.err.println(
                        "The exJello POP3 provider is NOT installed.");
                System.err.println("Skipping POP3 test.");
                return;
            }
            store.connect(host, user, password);
            inbox = store.getFolder("INBOX");
            inbox.open(Folder.READ_ONLY);
            System.out.println("Connected to POP3, inbox has " +
                    inbox.getMessageCount() + " messages.");
        } catch (Exception ex) {
            System.err.println("Error testing POP3:");
            ex.printStackTrace();
        } finally {
            if (inbox != null) {
                try {
                    inbox.close(false);
                } catch (Exception ignore) { }
            }
            if (store != null) {
                try {
                    store.close();
                } catch (Exception ignore) { }
            }
        }
    }

}
