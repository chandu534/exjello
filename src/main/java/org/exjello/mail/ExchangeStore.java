/*
Copyright (c) 2010 Eric Glass

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

import java.io.ByteArrayInputStream;

import java.net.InetAddress;
import java.net.MalformedURLException;
import java.net.URL;

import java.util.Properties;

import javax.mail.AuthenticationFailedException;
import javax.mail.Folder;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.URLName;

public class ExchangeStore extends Store {

    /**
     * Property specifying the mailbox to which the connection is made.
     * This is an e-mail address, e.g. "my.mailbox@example.com".
     */
    public static final String MAILBOX_PROPERTY =
            "org.exjello.mail.mailbox";

    /**
     * Specifies whether retrievals should be filtered to return only unread
     * messages, or all messages; "<code>true</code>" retrieves all messages,
     * "<code>false</code>" retrieves only unread messages.  Defaults to
     * "<code>false</code>" (only unread messages).
     */
    public static final String UNFILTERED_PROPERTY =
            "org.exjello.mail.unfiltered";

    /**
     * Specifies whether delete operations should really delete the message,
     * or just mark it as read. "<code>true</code>" performs a delete operation,
     * "<code>false</code>" just marks deleted messages as read.  Defaults to
     * "<code>false</code>" (don't delete, just mark read).
     */
    public static final String DELETE_PROPERTY = "org.exjello.mail.delete";

    /**
     * Limit on the number of messages that will be retrieved.
     */
    public static final String LIMIT_PROPERTY = "org.exjello.mail.limit";

    /**
     * Specifies whether to connect over HTTPS.  If the host parameter
     * to a connection specifies a protocol, that will take precedence over
     * this setting.
     */
    public static final String HTTPS_PROPERTY = "mail.pop3.ssl.enable";

    /**
     * Turns on debugging.
     */
    public static final String DEBUG_PROPERTY = "mail.debug";

    /**
     * Specifies the port for the connection.  Defaults to 80 for HTTP and
     * 443 for HTTPS.  If the host parameter to a connection specifies a
     * port, then that will take precedence over this setting.
     */
    public static final String PORT_PROPERTY = "mail.pop3.port";

    /**
     * Timeout in milliseconds for read operations on the socket.
     */
    public static final String TIMEOUT_PROPERTY = "mail.pop3.timeout";

    /**
     * Timeout in milliseconds for how long to wait for a connection to be
     * established.
     */
    public static final String CONNECTION_TIMEOUT_PROPERTY =
            "mail.pop3.connectiontimeout";

    /**
     * Local address to bind to, useful for a multi-homed host.
     */
    public static final String LOCAL_ADDRESS_PROPERTY =
            "mail.pop3.localaddress";

    private static final String DEBUG_PASSWORD_PROPERTY =
            "org.exjello.mail.debug.password";

    private static final int HTTP_PORT = 80;

    private static final int HTTPS_PORT = 443;

    private ExchangeConnection connection;

    private boolean unfiltered;

    private boolean delete;

    private int limit;

	public ExchangeStore(Session session, URLName urlname) {
		super(session, urlname);
	}

    protected boolean protocolConnect(String host, int port, String username,
            String password) throws MessagingException {
        boolean debug = Boolean.parseBoolean(
                session.getProperty(DEBUG_PROPERTY));
        boolean debugPassword = Boolean.parseBoolean(
                session.getProperty(DEBUG_PASSWORD_PROPERTY));
        String pwd = (password == null) ? null : debugPassword ? password :
                "<password>";
        if (host == null || username == null || password == null) {
            if (debug) {
                System.err.println("Missing parameter; host=\"" + host +
                        "\",username=\"" + username + "\",password=\"" +
                                pwd + "\"");
            }
            return false;
        }
        unfiltered = Boolean.parseBoolean(
                session.getProperty(UNFILTERED_PROPERTY));
        delete = Boolean.parseBoolean(session.getProperty(DELETE_PROPERTY));
        boolean secure = Boolean.parseBoolean(
                session.getProperty(HTTPS_PROPERTY));
        limit = -1;
        String limitString = session.getProperty(LIMIT_PROPERTY);
        if (limitString != null) {
            try {
                limit = Integer.parseInt(limitString);
            } catch (NumberFormatException ex) {
                throw new MessagingException("Invalid limit specified: " +
                        limitString);
            }
        }
        try {
            URL url = new URL(host);
            // if parsing succeeded, then strip out the components and use
            secure = "https".equalsIgnoreCase(url.getProtocol());
            host = url.getHost();
            int specifiedPort = url.getPort();
            if (specifiedPort != -1) port = specifiedPort;
        } catch (MalformedURLException ex) {
            if (debug) {
                System.err.println("Not parsing " + host +
                        " as a URL; using explicit options for " +
                                "secure, host, and port.");
            }
        }
        if (port == -1) {
            try {
                port = Integer.parseInt(session.getProperty(PORT_PROPERTY));
            } catch (Exception ignore) { }
            if (port == -1) port = secure ? HTTPS_PORT : HTTP_PORT;
        }
        String server = (secure ? "https://" : "http://") + host;
        if (secure ? (port != HTTPS_PORT) : (port != HTTP_PORT)) {
            server += ":" + port;
        }
        String mailbox = session.getProperty(MAILBOX_PROPERTY);
        int index = username.indexOf(':');
        if (index != -1) {
            mailbox = username.substring(index + 1);
            username = username.substring(0, index);
            String mailboxOptions = null;
            index = mailbox.indexOf('[');
            if (index != -1) {
                mailboxOptions = mailbox.substring(index + 1);
                mailboxOptions = mailboxOptions.substring(0,
                        mailboxOptions.indexOf(']'));
                mailbox = mailbox.substring(0, index);
            }
            if (mailboxOptions != null) {
                Properties props = null;
                try {
                    props = parseOptions(mailboxOptions);
                } catch (Exception ex) {
                    throw new MessagingException(
                            "Unable to parse mailbox options: " +
                                    ex.getMessage(), ex);
                }
                String value = props.getProperty("unfiltered");
                if (value != null) unfiltered = Boolean.parseBoolean(value);
                value = props.getProperty("delete");
                if (value != null) delete = Boolean.parseBoolean(value);
                value = props.getProperty("limit");
                if (value != null) {
                    try {
                        limit = Integer.parseInt(value);
                    } catch (NumberFormatException ex) {
                        throw new MessagingException(
                                "Invalid limit specified: " + value);
                    }
                }
            } else if (debug) {
                System.err.println("No mailbox options specified; " +
                        "using explicit limit, unfiltered, and delete.");
            }
        } else if (debug) {
            System.err.println("No mailbox specified in username; " +
                    "using explicit mailbox, limit, unfiltered, and delete.");
        }
        int timeout = -1;
        String timeoutString = session.getProperty(TIMEOUT_PROPERTY);
        if (timeoutString != null) {
            try {
                timeout = Integer.parseInt(timeoutString);
            } catch (NumberFormatException ex) {
                throw new MessagingException("Invalid timeout value: " +
                        timeoutString);
            }
        }
        int connectionTimeout = -1;
        timeoutString = session.getProperty(CONNECTION_TIMEOUT_PROPERTY);
        if (timeoutString != null) {
            try {
                connectionTimeout = Integer.parseInt(timeoutString);
            } catch (NumberFormatException ex) {
                throw new MessagingException(
                        "Invalid connection timeout value: " + timeoutString);
            }
        }
        InetAddress localAddress = null;
        String localAddressString = session.getProperty(LOCAL_ADDRESS_PROPERTY);
        if (localAddressString != null) {
            try {
                localAddress = InetAddress.getByName(localAddressString);
            } catch (Exception ex) {
                throw new MessagingException(
                        "Invalid local address specified: " +
                                localAddressString);
            }
        }
        if (mailbox == null) {
            throw new MessagingException("No mailbox specified.");
        }
        if (debug) {
            System.err.println("Server:\t" + server);
            System.err.println("Username:\t" + username);
            System.err.println("Password:\t" + pwd);
            System.err.println("Mailbox:\t" + mailbox);
            System.err.print("Options:\t");
            System.err.print((limit > 0) ? "Message Limit = " + limit :
                    "Unlimited Messages");
            System.err.print(unfiltered ? "; Unfiltered" :
                    "; Filtered to Unread");
            System.err.println(delete ? "; Delete Messages on Delete" :
                    "; Mark as Read on Delete");
            if (timeout > 0) {
                System.err.println("Read timeout:\t" + timeout + " ms");
            }
            if (connectionTimeout > 0) {
                System.err.println("Connection timeout:\t" + connectionTimeout +
                        " ms");
            }
        }
        connection = new ExchangeConnection(server, mailbox, username,
                password, timeout, connectionTimeout, localAddress);
        try {
            connection.connect();
        } catch (Exception ex) {
            if (debug) ex.printStackTrace();
            throw new AuthenticationFailedException(ex.getMessage());
        }
        return true;
    }

    public Folder getDefaultFolder() throws MessagingException {
        checkConnection();
        return new ExchangeFolder(this, "", connection);
    }

    public Folder getFolder(String name) throws MessagingException {
        checkConnection();
        return new ExchangeFolder(this, name, connection);
    }

    public Folder getFolder(URLName url) throws MessagingException {
        return getFolder(url.getFile());
    }

    public boolean isUnfiltered() {
        return unfiltered;
    }

    public boolean isDeleting() {
        return delete;
    }

    public int getLimit() {
        return limit;
    }

    private void checkConnection() throws IllegalStateException {
        if (!isConnected()) throw new IllegalStateException("Not connected.");
    }

    private static Properties parseOptions(String options) throws Exception {
        StringBuilder collector = new StringBuilder();
        String[] nvPairs = options.split("[,;]");
        for (String nvPair : nvPairs) collector.append(nvPair).append('\n');
        Properties properties = new Properties();
        properties.load(new ByteArrayInputStream(
                collector.toString().getBytes("ISO-8859-1")));
        return properties;
    }

}
