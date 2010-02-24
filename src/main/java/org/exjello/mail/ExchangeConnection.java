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

import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.IOException;
import java.io.OutputStream;
import java.io.StringWriter;

import java.net.InetAddress;
import java.net.URL;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import javax.mail.internet.SharedInputStream;

import javax.mail.util.SharedFileInputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;

import javax.xml.transform.dom.DOMSource;

import javax.xml.transform.stream.StreamResult;

import org.apache.commons.httpclient.Cookie;
import org.apache.commons.httpclient.Header;
import org.apache.commons.httpclient.HttpClient;
import org.apache.commons.httpclient.UsernamePasswordCredentials;

import org.apache.commons.httpclient.auth.AuthScope;

import org.apache.commons.httpclient.methods.ByteArrayRequestEntity;
import org.apache.commons.httpclient.methods.GetMethod;
import org.apache.commons.httpclient.methods.OptionsMethod;
import org.apache.commons.httpclient.methods.PostMethod;
import org.apache.commons.httpclient.methods.RequestEntity;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;

import org.xml.sax.helpers.DefaultHandler;

class ExchangeConnection {

    private static final Map<String, byte[]> RESOURCES =
            new HashMap<String, byte[]>();

    private static final String GET_UNREAD_MESSAGES_SQL_RESOURCE =
            "get-unread-messages.sql";

    private static final String GET_ALL_MESSAGES_SQL_RESOURCE =
            "get-all-messages.sql";

    private static final String SIGN_ON_URI =
            "/exchweb/bin/auth/owaauth.dll";

    private static final String HTTPMAIL_NAMESPACE = "urn:schemas:httpmail:";

    private static final String DAV_NAMESPACE = "DAV:";

    private static final String PROPFIND_METHOD = "PROPFIND";

    private static final String SEARCH_METHOD = "SEARCH";

    private static final String BDELETE_METHOD = "BDELETE";

    private static final String BPROPPATCH_METHOD = "BPROPPATCH";

    private static final String XML_CONTENT_TYPE =
            "text/xml; charset=\"UTF-8\"";

    private static final boolean[] ALLOWED_CHARS = new boolean[128];

    private static final char[] HEXABET = new char[] {
        '0', '1', '2', '3', '4', '5', '6', '7',
        '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'
    };

    private static byte[] findInboxEntity;

    private static byte[] unreadInboxEntity;

    private static byte[] allInboxEntity;

    private final String server;

    private final String mailbox;

    private final String username;

    private final String password;

    private final int timeout;

    private final int connectionTimeout;

    private final InetAddress localAddress;

    private String inboxRequest;

    private HttpClient client;

    private String inbox;

    static {
        // a - z
        for (int i = 97; i < 123; i++) ALLOWED_CHARS[i] = true;
        // A - Z
        for (int i = 64; i < 91; i++) ALLOWED_CHARS[i] = true;
        // 0 - 9
        for (int i = 48; i < 58; i++) ALLOWED_CHARS[i] = true;
        ALLOWED_CHARS[(int) '-'] = true;
        ALLOWED_CHARS[(int) '_'] = true;
        ALLOWED_CHARS[(int) '.'] = true;
        ALLOWED_CHARS[(int) '!'] = true;
        ALLOWED_CHARS[(int) '~'] = true;
        ALLOWED_CHARS[(int) '*'] = true;
        ALLOWED_CHARS[(int) '\''] = true;
        ALLOWED_CHARS[(int) '('] = true;
        ALLOWED_CHARS[(int) ')'] = true;
        ALLOWED_CHARS[(int) '%'] = true;
        ALLOWED_CHARS[(int) ':'] = true;
        ALLOWED_CHARS[(int) '@'] = true;
        ALLOWED_CHARS[(int) '&'] = true;
        ALLOWED_CHARS[(int) '='] = true;
        ALLOWED_CHARS[(int) '+'] = true;
        ALLOWED_CHARS[(int) '$'] = true;
        ALLOWED_CHARS[(int) ','] = true;
        ALLOWED_CHARS[(int) ';'] = true;
        ALLOWED_CHARS[(int) '/'] = true;
    }

    public ExchangeConnection(String server, String mailbox, String username,
            String password, int timeout, int connectionTimeout,
                    InetAddress localAddress) {
        this.server = server;
        this.mailbox = mailbox;
        this.username = username;
        this.password = password;
        this.timeout = timeout;
        this.connectionTimeout = connectionTimeout;
        this.localAddress = localAddress;
    }

    public void connect() throws Exception {
        synchronized (this) {
            inbox = null;
            try {
                signOn();
            } catch (Exception ex) {
                inbox = null;
                throw ex;
            }
        }
    }

    public List<String> getMessages(boolean includeRead, int limit)
            throws Exception {
        final List<String> messages = new ArrayList<String>();
        listInbox(includeRead, limit, new DefaultHandler() {
            private final StringBuilder content = new StringBuilder();
            public void characters(char[] ch, int start, int length)
                    throws SAXException {
                content.append(ch, start, length);
            }
            public void startElement(String uri, String localName,
                    String qName, Attributes attributes) throws SAXException {
                content.setLength(0);
            }
            public void endElement(String uri, String localName,
                    String qName) throws SAXException {
                if (!DAV_NAMESPACE.equals(uri)) return;
                if (!"href".equals(localName)) return;
                messages.add(content.toString());
            }
        });
        return Collections.unmodifiableList(messages);
    }

    public void delete(List<ExchangeMessage> messages) throws Exception {
        synchronized (this) {
            if (!isConnected()) {
                throw new IllegalStateException("Not connected.");
            }
            HttpClient client = getClient();
            String path = inbox;
            if (!path.endsWith("/")) path += "/";
            ExchangeMethod op = new ExchangeMethod(BDELETE_METHOD, path);
            op.setHeader("Content-Type", XML_CONTENT_TYPE);
            op.addHeader("If-Match", "*");
            op.addHeader("Brief", "t");
            op.setRequestEntity(createDeleteEntity(messages));
            InputStream stream = null;
            try {
                int status = client.executeMethod(op);
                stream = op.getResponseBodyAsStream();
                if (status >= 300) {
                    throw new IllegalStateException(
                            "Unable to delete messages.");
                }
            } finally {
                try {
                    if (stream != null) {
                        byte[] buf = new byte[65536];
                        try {
                            while (stream.read(buf, 0, 65536) != -1);
                        } catch (Exception ignore) {
                        } finally {
                            try {
                                stream.close();
                            } catch (Exception ignore2) { }
                        }
                    }
                } finally {
                    op.releaseConnection();
                }
            }
        }
    }

    public void markRead(List<ExchangeMessage> messages) throws Exception {
        synchronized (this) {
            if (!isConnected()) {
                throw new IllegalStateException("Not connected.");
            }
            HttpClient client = getClient();
            String path = inbox;
            if (!path.endsWith("/")) path += "/";
            ExchangeMethod op = new ExchangeMethod(BPROPPATCH_METHOD, path);
            op.setHeader("Content-Type", XML_CONTENT_TYPE);
            op.addHeader("If-Match", "*");
            op.addHeader("Brief", "t");
            op.setRequestEntity(createMarkReadEntity(messages));
            InputStream stream = null;
            try {
                int status = client.executeMethod(op);
                stream = op.getResponseBodyAsStream();
                if (status >= 300) {
                    throw new IllegalStateException(
                            "Unable to mark messages read.");
                }
            } finally {
                try {
                    if (stream != null) {
                        byte[] buf = new byte[65536];
                        try {
                            while (stream.read(buf, 0, 65536) != -1);
                        } catch (Exception ignore) {
                        } finally {
                            try {
                                stream.close();
                            } catch (Exception ignore2) { }
                        }
                    }
                } finally {
                    op.releaseConnection();
                }
            }
        }
    }

    public InputStream getInputStream(ExchangeMessage message)
            throws Exception {
        synchronized (this) {
            if (!isConnected()) {
                throw new IllegalStateException("Not connected.");
            }
            HttpClient client = getClient();
            GetMethod op = new GetMethod(escape(message.getUrl()));
            op.setRequestHeader("Translate", "F");
            InputStream stream = null;
            try {
                int status = client.executeMethod(op);
                stream = op.getResponseBodyAsStream();
                if (status >= 300) {
                    throw new IllegalStateException("Unable to obtain inbox: " +
                            status);
                }
                final File tempFile = File.createTempFile("exmail", null, null);
                tempFile.deleteOnExit();
                OutputStream output = new FileOutputStream(tempFile);
                byte[] buf = new byte[65536];
                int count;
                while ((count = stream.read(buf, 0, 65536)) != -1) {
                    output.write(buf, 0, count);
                }
                output.flush();
                output.close();
                stream.close();
                stream = null;
                return new CachedMessageStream(tempFile, (ExchangeFolder)
                        message.getFolder());
            } finally {
                try {
                    if (stream != null) {
                        byte[] buf = new byte[65536];
                        try {
                            while (stream.read(buf, 0, 65536) != -1);
                        } catch (Exception ignore) { 
                        } finally {
                            try {
                                stream.close();
                            } catch (Exception ignore2) { }
                        }
                    }
                } finally {
                    op.releaseConnection();
                }
            }
        }
    }

    private boolean isConnected() {
        return (inbox != null);
    }

    private void listInbox(boolean includeRead, int limit,
            DefaultHandler handler) throws Exception {
        synchronized (this) {
            if (!isConnected()) {
                throw new IllegalStateException("Not connected.");
            }
            HttpClient client = getClient();
            ExchangeMethod op = new ExchangeMethod(SEARCH_METHOD, inbox);
            op.setHeader("Content-Type", XML_CONTENT_TYPE);
            if (limit > 0) op.setHeader("Range", "rows=0-" + limit);
            op.setHeader("Brief", "t");
            op.setRequestEntity(includeRead ? createAllInboxEntity() :
                    createUnreadInboxEntity());
            InputStream stream = null;
            try {
                int status = client.executeMethod(op);
                stream = op.getResponseBodyAsStream();
                if (status >= 300) {
                    throw new IllegalStateException("Unable to obtain inbox.");
                }
                SAXParserFactory spf = SAXParserFactory.newInstance();
                spf.setNamespaceAware(true);
                SAXParser parser = spf.newSAXParser();
                parser.parse(stream, handler);
                stream.close();
                stream = null;
            } finally {
                try {
                    if (stream != null) {
                        byte[] buf = new byte[65536];
                        try {
                            while (stream.read(buf, 0, 65536) != -1);
                        } catch (Exception ignore) {
                        } finally {
                            try {
                                stream.close();
                            } catch (Exception ignore2) { }
                        }
                    }
                } finally {
                    op.releaseConnection();
                }
            }
        }
    }

    private void findInbox() throws Exception {
        inbox = null;
        HttpClient client = getClient();
        ExchangeMethod op = new ExchangeMethod(PROPFIND_METHOD,
                server + "/exchange/" + mailbox);
        op.setHeader("Content-Type", XML_CONTENT_TYPE);
        op.setHeader("Depth", "0");
        op.setHeader("Brief", "t");
        op.setRequestEntity(createFindInboxEntity());
        InputStream stream = null;
        try {
            int status = client.executeMethod(op);
            stream = op.getResponseBodyAsStream();
            if (status >= 300) {
                throw new IllegalStateException("Unable to obtain inbox.");
            }
            SAXParserFactory spf = SAXParserFactory.newInstance();
            spf.setNamespaceAware(true);
            SAXParser parser = spf.newSAXParser();
            parser.parse(stream, new DefaultHandler() {
                private final StringBuilder content = new StringBuilder();
                public void startElement(String uri, String localName,
                        String qName, Attributes attributes)
                                throws SAXException {
                    content.setLength(0);
                }
                public void characters(char[] ch, int start, int length)
                        throws SAXException {
                    content.append(ch, start, length);
                }
                public void endElement(String uri, String localName,
                        String qName) throws SAXException {
                    if (!HTTPMAIL_NAMESPACE.equals(uri)) return;
                    if (!"inbox".equals(localName)) return;
                    ExchangeConnection.this.inbox = content.toString();
                }
            });
            stream.close();
            stream = null;
        } finally {
            try {
                if (stream != null) {
                    byte[] buf = new byte[65536];
                    try {
                        while (stream.read(buf, 0, 65536) != -1);
                    } catch (Exception ignore) {
                    } finally {
                        try {
                            stream.close();
                        } catch (Exception ignore2) { }
                    }
                }
            } finally {
                op.releaseConnection();
            }
        }
    }

    private HttpClient getClient() {
        synchronized (this) {
            if (client == null) {
                client = new HttpClient();
                if (timeout > 0) client.getParams().setSoTimeout(timeout);
                if (connectionTimeout > 0) {
                    client.getHttpConnectionManager().getParams(
                            ).setConnectionTimeout(connectionTimeout);
                }
                if (localAddress != null) {
                    client.getHostConfiguration().setLocalAddress(localAddress);
                }
            }
            return client;
        }
    }

    private void signOn() throws Exception {
        HttpClient client = getClient();
        URL serverUrl = new URL(server);
        String host = serverUrl.getHost();
        int port = serverUrl.getPort();
        if (port == -1) port = serverUrl.getDefaultPort();
        AuthScope authScope = new AuthScope(host, port);
        client.getState().setCredentials(authScope,
                new UsernamePasswordCredentials(username, password));
        boolean authenticated = false;
        OptionsMethod authTest = new OptionsMethod(server + "/exchange");
        try {
            authenticated = (client.executeMethod(authTest) < 400);
        } finally {
            try {
                InputStream stream = authTest.getResponseBodyAsStream();
                byte[] buf = new byte[65536];
                try {
                    while (stream.read(buf, 0, 65536) != -1);
                } catch (Exception ignore) {
                } finally {
                    try {
                        stream.close();
                    } catch (Exception ignore2) { }
                }
            } finally {
                authTest.releaseConnection();
            }
        }
        if (!authenticated) {
            PostMethod op = new PostMethod(server + SIGN_ON_URI);
            op.setRequestHeader("Content-Type",
                    "application/x-www-form-urlencoded");
            op.addParameter("destination", server + "/exchange");
            op.addParameter("flags", "0");
            op.addParameter("username", username);
            op.addParameter("password", password);
            try {
                int status = client.executeMethod(op);
                if (status >= 400) {
                    throw new IllegalStateException("Sign-on failed: " +
                            status);
                }
            } finally {
                try {
                    InputStream stream = op.getResponseBodyAsStream();
                    byte[] buf = new byte[65536];
                    try {
                        while (stream.read(buf, 0, 65536) != -1);
                    } catch (Exception ignore) {
                    } finally {
                        try {
                            stream.close();
                        } catch (Exception ignore2) { }
                    }
                } finally {
                    op.releaseConnection();
                }
            }
        }
        findInbox();
    }

    private static RequestEntity createFindInboxEntity() throws Exception {
        synchronized (ExchangeConnection.class) {
            if (findInboxEntity == null) {
                DocumentBuilderFactory dbf =
                        DocumentBuilderFactory.newInstance();
                dbf.setNamespaceAware(true);
                Document doc = dbf.newDocumentBuilder().newDocument();
                Element propfind = doc.createElementNS(DAV_NAMESPACE,
                        "propfind");
                doc.appendChild(propfind);
                Element prop = doc.createElementNS(DAV_NAMESPACE, "prop");
                propfind.appendChild(prop);
                Element inbox = doc.createElementNS(HTTPMAIL_NAMESPACE,
                        "inbox");
                prop.appendChild(inbox);
                ByteArrayOutputStream collector = new ByteArrayOutputStream();
                Transformer transformer =
                        TransformerFactory.newInstance().newTransformer();
                transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
                transformer.transform(new DOMSource(doc),
                        new StreamResult(collector));
                findInboxEntity = collector.toByteArray();
            }
            return new ByteArrayRequestEntity(findInboxEntity,
                    XML_CONTENT_TYPE);
        }
    }

    private static RequestEntity createUnreadInboxEntity() throws Exception {
        synchronized (ExchangeConnection.class) {
            if (unreadInboxEntity == null) {
                unreadInboxEntity = createSearchEntity(new String(getResource(
                        GET_UNREAD_MESSAGES_SQL_RESOURCE), "UTF-8"));
            }
            return new ByteArrayRequestEntity(unreadInboxEntity,
                    XML_CONTENT_TYPE);
        }
    }

    private static RequestEntity createAllInboxEntity() throws Exception {
        synchronized (ExchangeConnection.class) {
            if (allInboxEntity == null) {
                allInboxEntity = createSearchEntity(new String(getResource(
                        GET_ALL_MESSAGES_SQL_RESOURCE), "UTF-8"));
            }
            return new ByteArrayRequestEntity(allInboxEntity, XML_CONTENT_TYPE);
        }
    }

    private static byte[] createSearchEntity(String sqlString)
            throws Exception {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        dbf.setNamespaceAware(true);
        Document doc = dbf.newDocumentBuilder().newDocument();
        Element searchRequest = doc.createElementNS(DAV_NAMESPACE,
                "searchrequest");
        doc.appendChild(searchRequest);
        Element sql = doc.createElementNS(DAV_NAMESPACE, "sql");
        searchRequest.appendChild(sql);
        sql.appendChild(doc.createTextNode(sqlString));
        ByteArrayOutputStream collector = new ByteArrayOutputStream();
        Transformer transformer =
                TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        transformer.transform(new DOMSource(doc), new StreamResult(collector));
        return collector.toByteArray();
    }

    private static RequestEntity createDeleteEntity(
            List<ExchangeMessage> messages) throws Exception {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        dbf.setNamespaceAware(true);
        Document doc = dbf.newDocumentBuilder().newDocument();
        Element delete = doc.createElementNS(DAV_NAMESPACE, "delete");
        doc.appendChild(delete);
        Element target = doc.createElementNS(DAV_NAMESPACE, "target");
        delete.appendChild(target);
        for (ExchangeMessage message : messages) {
            String url = message.getUrl();
            Element href = doc.createElementNS(DAV_NAMESPACE, "href");
            target.appendChild(href);
            String file = url.substring(url.lastIndexOf("/") + 1);
            href.appendChild(doc.createTextNode(file));
        }
        ByteArrayOutputStream collector = new ByteArrayOutputStream();
        Transformer transformer =
                TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        transformer.transform(new DOMSource(doc),
                new StreamResult(collector));
        return new ByteArrayRequestEntity(collector.toByteArray(),
                XML_CONTENT_TYPE);
    }

    private static RequestEntity createMarkReadEntity(
            List<ExchangeMessage> messages) throws Exception {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        dbf.setNamespaceAware(true);
        Document doc = dbf.newDocumentBuilder().newDocument();
        Element propertyUpdate = doc.createElementNS(DAV_NAMESPACE,
                "propertyupdate");
        doc.appendChild(propertyUpdate);
        Element target = doc.createElementNS(DAV_NAMESPACE, "target");
        propertyUpdate.appendChild(target);
        for (ExchangeMessage message : messages) {
            String url = message.getUrl();
            Element href = doc.createElementNS(DAV_NAMESPACE, "href");
            target.appendChild(href);
            String file = url.substring(url.lastIndexOf("/") + 1);
            href.appendChild(doc.createTextNode(file));
        }
        Element set = doc.createElementNS(DAV_NAMESPACE, "set");
        propertyUpdate.appendChild(set);
        Element prop = doc.createElementNS(DAV_NAMESPACE, "prop");
        set.appendChild(prop);
        Element read = doc.createElementNS(HTTPMAIL_NAMESPACE, "read");
        prop.appendChild(read);
        read.appendChild(doc.createTextNode("1"));
        ByteArrayOutputStream collector = new ByteArrayOutputStream();
        Transformer transformer =
                TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        transformer.transform(new DOMSource(doc),
                new StreamResult(collector));
        return new ByteArrayRequestEntity(collector.toByteArray(),
                XML_CONTENT_TYPE);
    }

    private static byte[] getResource(String resource) {
        if (resource == null) return null;
        synchronized (RESOURCES) {
            byte[] content = RESOURCES.get(resource);
            if (content != null) return content;
            try {
                InputStream input =
                        ExchangeConnection.class.getResourceAsStream(resource);
                ByteArrayOutputStream collector =
                        new ByteArrayOutputStream();
                byte[] buf = new byte[65536];
                int count;
                while ((count = input.read(buf, 0, 65536)) != -1) {
                    collector.write(buf, 0, count);
                }
                input.close();
                collector.flush();
                content = collector.toByteArray();
                RESOURCES.put(resource, content);
                return content;
            } catch (Exception ex) {
                throw new IllegalStateException(ex.getMessage(), ex);
            }
        }
    }

    private static String escape(String url) {
        StringBuilder collector = new StringBuilder(url);
        for (int i = collector.length() - 1; i >= 0; i--) {
            int value = (int) collector.charAt(i);
            if (value > 127 || !ALLOWED_CHARS[value]) {
                collector.deleteCharAt(i);
                collector.insert(i, HEXABET[value & 0x0f]);
                value >>>= 4;
                collector.insert(i, HEXABET[value & 0x0f]);
                value >>>= 4;
                collector.insert(i, '%');
                if (value > 0) {
                    collector.insert(i, HEXABET[value & 0x0f]);
                    value >>>= 4;
                    collector.insert(i, HEXABET[value & 0x0f]);
                    collector.insert(i, '%');
                }
            }
        }
        return collector.toString();
    }

}
