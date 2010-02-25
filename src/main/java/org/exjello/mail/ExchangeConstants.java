package org.exjello.mail;

public interface ExchangeConstants {

    /**
     * Property specifying the mailbox to which the connection is made.
     * This is an e-mail address, e.g. "my.mailbox@example.com"; if specified,
     * this will take precedence over the "mail.smtp.from" setting.
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
     * Property specifying the mailbox to which the connection is made
     * (used for both SMTP and POP3). This is an e-mail address,
     * e.g. "my.mailbox@example.com".
     */
    public static final String FROM_PROPERTY = "from";

    /**
     * Specifies whether to connect over HTTPS.
     * If the host parameter to a connection specifies a protocol,
     * that will take precedence over this setting.
     */
    public static final String SSL_PROPERTY = "ssl.enable";

    /**
     * Specifies the port for the connection.
     * Defaults to 80 for HTTP and 443 for HTTPS.  If the host parameter to
     * a connection specifies a port, then that will take precedence over
     * this setting.
     */
    public static final String PORT_PROPERTY = "port";

    /**
     * Timeout in milliseconds for read operations on the socket.
     */
    public static final String TIMEOUT_PROPERTY = "timeout";

    /**
     * Timeout in milliseconds for how long to wait for a connection to be
     * established.
     */
    public static final String CONNECTION_TIMEOUT_PROPERTY =
            "connectiontimeout";

    /**
     * Local address to bind to, useful for a multi-homed host.
     */
    public static final String LOCAL_ADDRESS_PROPERTY = "localaddress";

}
