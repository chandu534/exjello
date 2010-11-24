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
     * Specifies whether retrievals should be filtered to return only
     * messages received after certain date.
     * Condition is made on 
     * "urn:schemas:httpmail:datereceived"
     * Date is specified as a "<code>String</code>" in ISO_8601 format
     * Date pattern is "yyyy-MM-dd'T'HH:mm:ss'Z'"
     */
    public static final String FILTER_LAST_CHECK =
        "org.exjello.mail.filterlastcheck";

    /**
     * Specifies whether retrievals should be filtered to return only
     * messages contains certain text in from field.
     * Condition is made on 
     * "urn:schemas:httpmail:fromemail" LIKE '%your_text%' 
     * For Example you can specify "@yourdomain.com" for filter
     * all internal messages 
     */  
    public static final String FILTER_FROM_PROPERTY =
            "org.exjello.mail.filterfrom";
    
    public static final String FILTER_NOT_FROM_PROPERTY =
        "org.exjello.mail.filternotfrom";

    /**
     * Specifies whether retrievals should be filtered to return only
     * messages contains certain text in to field.
     * Condition is made on 
     * "urn:schemas:httpmail:to" LIKE '%your_text%'
     * For Example you can specify "help@yourdomain.com" for filter
     * all messages sent to an alias
     */    
    public static final String FILTER_TO_PROPERTY =
        "org.exjello.mail.filterto";

    
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
