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

import java.io.InputStream;
import java.io.IOException;

import java.util.Enumeration;

import javax.mail.IllegalWriteException;
import javax.mail.MessagingException;

import javax.mail.internet.InternetHeaders;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.SharedInputStream;

class ExchangeMessage extends MimeMessage {

    private final String url;

    private final ExchangeConnection connection;

    public ExchangeMessage(ExchangeFolder folder, int messageNumber,
            String url, ExchangeConnection connection)
                    throws MessagingException {
        super(folder, messageNumber);
        this.url = url;
        this.connection = connection;
    }

    public String getUrl() {
        return url;
    }

    protected InputStream getContentStream() throws MessagingException {
        try {
            synchronized (this) {
                if (contentStream == null) {
                    InputStream stream = connection.getInputStream(this);
                    headers = new InternetHeaders(stream);
                    SharedInputStream shared = (SharedInputStream) stream;
                    contentStream = shared.newStream(shared.getPosition(), -1l);
                    stream = null;
                }
                return super.getContentStream();
            }
        } catch (MessagingException ex) {
            throw ex;
        } catch (Exception ex) {
            throw new MessagingException(ex.getMessage(), ex);
        }
    }

    public String[] getHeader(String name) throws MessagingException {
        synchronized (this) {
            if (headers == null) loadHeaders();
            return headers.getHeader(name);
        }
    }

    public String getHeader(String name, String delimiter)
            throws MessagingException {
        synchronized (this) {
            if (headers == null) loadHeaders();
            return headers.getHeader(name, delimiter);
        }
    }

    public void setHeader(String name, String value) throws MessagingException {
        throw new IllegalWriteException("Write not supported.");
    }

    public void addHeader(String name, String value) throws MessagingException {
        throw new IllegalWriteException("Write not supported.");
    }

    public void removeHeader(String name) throws MessagingException {
        throw new IllegalWriteException("Write not supported.");
    }

    public Enumeration getAllHeaders() throws MessagingException {
        synchronized (this) {
            if (headers == null) loadHeaders();
            return headers.getAllHeaders();
        }
    }

    public Enumeration getMatchingHeaders(String[] names)
                        throws MessagingException {
        synchronized (this) {
            if (headers == null) loadHeaders();
            return headers.getMatchingHeaders(names);
        }
    }

    public Enumeration getNonMatchingHeaders(String[] names)
                        throws MessagingException {
        synchronized (this) {
            if (headers == null) loadHeaders();
            return headers.getNonMatchingHeaders(names);
        }
    }

    public void addHeaderLine(String line) throws MessagingException {
        throw new IllegalWriteException("Write not supported.");
    }

    public Enumeration getAllHeaderLines() throws MessagingException {
        synchronized (this) {
            if (headers == null) loadHeaders();
            return headers.getAllHeaderLines();
        }
    }

    public Enumeration getMatchingHeaderLines(String[] names)
                                        throws MessagingException {
        synchronized (this) {
            if (headers == null) loadHeaders();
            return headers.getMatchingHeaderLines(names);
        }
    }

    public Enumeration getNonMatchingHeaderLines(String[] names)
                                        throws MessagingException {
        synchronized (this) {
            if (headers == null) loadHeaders();
            return headers.getNonMatchingHeaderLines(names);
        }
    }

    public void saveChanges() throws MessagingException {
        throw new IllegalWriteException("Write not supported.");
    }

    private void loadHeaders() throws MessagingException {
        try {
            if (headers == null) getContentStream().close();
        } catch (MessagingException ex) {
            throw ex;
        } catch (Exception ignore) { }
    }

}

