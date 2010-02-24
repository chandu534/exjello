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

import java.io.IOException;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.httpclient.HttpConnection;
import org.apache.commons.httpclient.HttpException;
import org.apache.commons.httpclient.HttpState;

import org.apache.commons.httpclient.methods.EntityEnclosingMethod;

class ExchangeMethod extends EntityEnclosingMethod {

    private final String methodName;

    private final Map<String, String> setHeaders =
            new HashMap<String, String>();

    private final Map<String, List<String>> addHeaders =
            new HashMap<String, List<String>>();

    public ExchangeMethod(String methodName, String path) {
        super(path);
        this.methodName = methodName;
    }

    public String getName() {
        return methodName;
    }

    public void addHeader(String name, String value) {
        if (name == null) return;
        if (value == null) value = "";
        synchronized (addHeaders) {
            List<String> values = addHeaders.get(name);
            if (values == null) {
                addHeaders.put(name, (values = new ArrayList<String>()));
            }
            values.add(value);
        }
    }

    public void setHeader(String name, String value) {
        if (name == null) return;
        if (value == null) value = "";
        synchronized (setHeaders) {
            setHeaders.put(name, value);
        }
    }

    public void addRequestHeaders(HttpState state, HttpConnection conn)
            throws IOException, HttpException {
        super.addRequestHeaders(state, conn);
        synchronized (setHeaders) {
            for (Map.Entry<String, String> setHeader : setHeaders.entrySet()) {
                setRequestHeader(setHeader.getKey(), setHeader.getValue());
            }
        }
        synchronized (addHeaders) {
            for (Map.Entry<String, List<String>> addHeader :
                    addHeaders.entrySet()) {
                String header = addHeader.getKey();
                for (String value : addHeader.getValue()) {
                    addRequestHeader(header, value);
                }
            }
        }
    }

}
