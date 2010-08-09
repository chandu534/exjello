/*
Copyright (c) 2010 Eric Glass, Mirco Attocchi

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*/

package org.exjello.mail;

import java.util.Properties;

import javax.mail.AuthenticationFailedException;
import javax.mail.Folder;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.URLName;

public class ExchangeStore extends Store {

    private static final String POP3_PROTOCOL = "pop3";

    private ExchangeConnection connection;

	public ExchangeStore(Session session, URLName urlname) {
		super(session, urlname);
	}
	
	public void connect(String host,
            String user,
            String password,
            String mailbox)
     throws MessagingException {
		Properties props = System.getProperties();
		props.setProperty(ExchangeConstants.MAILBOX_PROPERTY, mailbox);
		super.connect(host, user, password);
	}


    protected boolean protocolConnect(String host, int port, String username,
            String password) throws MessagingException {
        synchronized (this) {
            try {
                connection = ExchangeConnection.createConnection(POP3_PROTOCOL,
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

    public Folder getDefaultFolder() throws MessagingException {
        synchronized (this) {
            checkConnection();
            return new ExchangeFolder(this, "", connection);
        }
    }

    public Folder getFolder(String name) throws MessagingException {
        synchronized (this) {
            checkConnection();
            return new ExchangeFolder(this, name, connection);
        }
    }

    public Folder getFolder(URLName url) throws MessagingException {
        return getFolder(url.getFile());
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
