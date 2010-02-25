package org.exjello.mail;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import javax.mail.Address;
import javax.mail.AuthenticationFailedException;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.URLName;

import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

public class ExchangeTransport extends Transport {

    private static final String SMTP_PROTOCOL = "smtp";

    private ExchangeConnection connection;

	public ExchangeTransport(Session session, URLName urlname) {
		super(session, urlname);
	}

	public void sendMessage(Message message, Address[] addresses)
			throws MessagingException {
        if (!(message instanceof MimeMessage)) {
            throw new MessagingException("Invalid message; " +
                    "only RFC822 MIME messages are supported.");
        }
        MimeMessage mimeMessage = (MimeMessage) message;
        if (addresses == null || addresses.length == 0) {
            throw new MessagingException("No addresses specified.");
        }
        List<Address> targetedRecipients =
                Arrays.asList(addresses);
        Address[] addressList =
                mimeMessage.getRecipients(Message.RecipientType.TO);
        if (addressList == null) addressList = new Address[0];
        List<Address> toRecipients = new ArrayList<Address>(
                Arrays.asList(addressList));
        addressList = mimeMessage.getRecipients(Message.RecipientType.CC);
        if (addressList == null) addressList = new Address[0];
        List<Address> ccRecipients = new ArrayList<Address>(
                Arrays.asList(addressList));
        addressList = mimeMessage.getRecipients(Message.RecipientType.BCC);
        if (addressList == null) addressList = new Address[0];
        List<Address> bccRecipients = new ArrayList<Address>(
                Arrays.asList(addressList));
        Iterator<Address> recipients = toRecipients.iterator();
        while (recipients.hasNext()) {
            if (!targetedRecipients.contains(recipients.next())) {
                recipients.remove();
            }
        }
        recipients = ccRecipients.iterator();
        while (recipients.hasNext()) {
            if (!targetedRecipients.contains(recipients.next())) {
                recipients.remove();
            }
        }
        recipients = bccRecipients.iterator();
        while (recipients.hasNext()) {
            if (!targetedRecipients.contains(recipients.next())) {
                recipients.remove();
            }
        }
        addressList = mimeMessage.getAllRecipients();
        if (addressList == null) addressList = new Address[0];
        List<Address> messageRecipients = Arrays.asList(addressList);
        for (Address address : addresses) {
            if (!(address instanceof InternetAddress)) {
                throw new MessagingException("Invalid address: " + address);
            }
            if (!messageRecipients.contains(address)) {
                bccRecipients.add(address);
            }
        }
        mimeMessage.setRecipients(Message.RecipientType.TO,
                toRecipients.isEmpty() ? (Address[]) null :
                        toRecipients.toArray(new Address[toRecipients.size()]));
        mimeMessage.setRecipients(Message.RecipientType.CC,
                ccRecipients.isEmpty() ? (Address[]) null :
                        ccRecipients.toArray(new Address[ccRecipients.size()]));
        mimeMessage.setRecipients(Message.RecipientType.BCC,
                bccRecipients.isEmpty() ? (Address[]) null :
                        bccRecipients.toArray(
                                new Address[bccRecipients.size()]));
        synchronized (this) {
            checkConnection();
            try {
                connection.send(mimeMessage);
            } catch (Exception ex) {
                throw new MessagingException(ex.getMessage(), ex);
            }
        }
	}

    protected boolean protocolConnect(String host, int port, String username,
            String password) throws MessagingException {
        synchronized (this) {
            try {
                connection = ExchangeConnection.createConnection(SMTP_PROTOCOL,
                        session, host, port, username, password);
            } catch (Exception ex) {
                throw new MessagingException(ex.getMessage(), ex);
            }
            try {
                connection.connect();
            } catch (Exception ex) {
                throw new AuthenticationFailedException(ex.getMessage());
            }
        }
        return true;
    }

    public boolean isConnected() {
        synchronized (this) {
            return super.isConnected() && (connection != null);
        }
    }

    protected void setConnected(boolean connected) {
        synchronized (this) {
            super.setConnected(connected);
            if (!connected) connection = null;
        }
    }

    private void checkConnection() throws IllegalStateException {
        if (!isConnected()) throw new IllegalStateException("Not connected.");
    }

}
