package com.qa.framework.library.httpclient;


import com.qa.framework.bean.Param;
import com.qa.framework.config.PropConfig;
import org.apache.http.HttpHost;
import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.methods.HttpPut;
import org.apache.http.message.BasicNameValuePair;
import org.apache.log4j.Logger;

import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by apple on 15/11/19.
 */
public class HttpMethod {
    protected final static Logger logger = Logger.getLogger(HttpMethod.class);
    private static boolean useProxy = PropConfig.isUseProxy();
    private static String localhost = PropConfig.getLocalhost();
    private static Integer localport = Integer.valueOf(PropConfig.getLocalport());
    private static Integer timeout = Integer.valueOf(PropConfig.getTimeout());

    public static String getUrl(String url, List<Param> params) {
        StringBuilder webPath = new StringBuilder();
        webPath.append(PropConfig.getWebPath());
        if (url.contains("/")) {
            webPath.append(url);
        } else {
            webPath.append(url).append("/");
        }
        if (params != null) {
            for (Param param : params) {
                if (param.isShow()) {
                    webPath.append(param.getName()).append("/").append(param.getValue(false)).append("/");
                }
            }
        }
        if (webPath.substring(webPath.length() - 1).equals("/")) {
            return webPath.substring(0, webPath.length() - 1);
        }

        return webPath.toString();
    }

    public static String postUrl(String url) {
        return PropConfig.getWebPath() + url;
    }

    public static String useGetMethod(String url, List<Param> params, boolean storeCookie, boolean useCookie) {
        String uri = getUrl(url, params);
        logger.info("拼接后的web地址为:" + uri);
        HttpGet get = new HttpGet(uri);
        if (useProxy) {
            HttpHost proxy = new HttpHost(localhost, localport, "http");
            RequestConfig requestConfig = RequestConfig.custom().setSocketTimeout(timeout).setConnectTimeout(timeout).setProxy(proxy).build();
            get.setConfig(requestConfig);
        }
        HttpConnectionImp imp = new HttpConnectionImp(get);
        String returnResult = imp.getResponseResult(storeCookie, useCookie);
        logger.info("actual result:" + returnResult);
        return returnResult;
    }

    public static String useGetMethod(String url, List<Param> params, boolean storeCookie, boolean useCookie, int trytimes) {
        String uri = getUrl(url, params);
        logger.info("拼接后的web地址为:" + uri);
        HttpGet get = new HttpGet(uri);
        if (useProxy) {
            HttpHost proxy = new HttpHost(localhost, localport, "http");
            RequestConfig requestConfig = RequestConfig.custom().setSocketTimeout(timeout).setConnectTimeout(timeout).setProxy(proxy).build();
            get.setConfig(requestConfig);
        }
        HttpConnectionImp imp = new HttpConnectionImp(get);
        String returnResult = imp.getResponseResult(storeCookie, useCookie);
        if (returnResult != null) {
            logger.info("actual result:" + returnResult);
            return returnResult;
        } else {
            int count = 0;
            while (count < trytimes && returnResult == null) {
                returnResult = imp.getResponseResult(storeCookie, useCookie);
                count++;
                try {
                    Thread.sleep(2000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }
        }
        logger.info("actual result:" + returnResult);
        return returnResult;
    }

    public static String usePostMethod(String url, List<Param> params, boolean storeCookie, boolean useCookie) {
        String uri = postUrl(url);
        logger.info("拼接后的web地址为:" + uri);
        HttpPost httpPost = new HttpPost(uri);
        RequestConfig requestConfig = null;
        if (useProxy) {
            HttpHost proxy = new HttpHost(localhost, localport, "http");
            requestConfig = RequestConfig.custom().setSocketTimeout(timeout).setConnectTimeout(timeout).setProxy(proxy).build();
        } else {
            requestConfig = RequestConfig.custom().setSocketTimeout(timeout).setConnectTimeout(timeout).build();
        }
        httpPost.setConfig(requestConfig);
        List<BasicNameValuePair> basicNameValuePairs = new ArrayList<BasicNameValuePair>();
        if (params != null) {
            for (Param param : params) {
                BasicNameValuePair basicNameValuePair = new BasicNameValuePair(param.getName(), param.getValue(false));
                basicNameValuePairs.add(basicNameValuePair);
            }
        }
        try {
            httpPost.setEntity(new UrlEncodedFormEntity(basicNameValuePairs, "UTF-8"));
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        HttpConnectionImp imp = new HttpConnectionImp(httpPost);
        String returnResult = imp.getResponseResult(storeCookie, useCookie);
        logger.debug("实际结果:" + returnResult);
        return returnResult;
    }

    public static String usePutMethod(String url, List<Param> params, boolean storeCookie, boolean useCookie) {
        String uri = postUrl(url);
        logger.info("拼接后的web地址为:" + uri);
        HttpPut httpPut = new HttpPut(uri);
//        HttpPost httpPost = new HttpPost(uri);
        RequestConfig requestConfig = null;
        if (useProxy) {
            HttpHost proxy = new HttpHost(localhost, localport, "http");
            requestConfig = RequestConfig.custom().setSocketTimeout(timeout).setConnectTimeout(timeout).setProxy(proxy).build();
        } else {
            requestConfig = RequestConfig.custom().setSocketTimeout(timeout).setConnectTimeout(timeout).build();
        }
        httpPut.setConfig(requestConfig);
        List<BasicNameValuePair> basicNameValuePairs = new ArrayList<BasicNameValuePair>();
        if (params != null) {
            for (Param param : params) {
                BasicNameValuePair basicNameValuePair = new BasicNameValuePair(param.getName(), param.getValue(false));
                basicNameValuePairs.add(basicNameValuePair);
            }
        }
        try {
            httpPut.setEntity(new UrlEncodedFormEntity(basicNameValuePairs, "UTF-8"));
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        HttpConnectionImp imp = new HttpConnectionImp(httpPut);
        String returnResult = imp.getResponseResult(storeCookie, useCookie);
        logger.debug("实际结果:" + returnResult);
        return returnResult;
    }
}
