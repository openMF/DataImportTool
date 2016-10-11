package org.openmf.mifos.dataimport.http;

import java.net.HttpURLConnection;
import java.security.KeyManagementException;
import java.security.KeyStoreException;
import java.security.NoSuchAlgorithmException;
import java.security.cert.CertificateException;
import java.security.cert.X509Certificate;

import javax.net.ssl.HostnameVerifier;
import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;
import javax.net.ssl.SSLSession;
import javax.net.ssl.SSLSocketFactory;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;

public class TrustModifier {
   private static final TrustingHostnameVerifier 
      TRUSTING_HOSTNAME_VERIFIER = new TrustingHostnameVerifier();
   private static SSLSocketFactory factory;

   /** Call this with any HttpURLConnection, and it will 
    modify the trust settings if it is an HTTPS connection. */
   public static String relaxHostChecking(HttpURLConnection conn) {
      try{
        if (conn instanceof HttpsURLConnection) {
           HttpsURLConnection httpsConnection = (HttpsURLConnection) conn;
           SSLSocketFactory factory = prepFactory(httpsConnection);
           httpsConnection.setSSLSocketFactory(factory);
           httpsConnection.setHostnameVerifier(TRUSTING_HOSTNAME_VERIFIER);
        }
        return "";
      } catch (KeyStoreException kse) {
    	  return kse.getMessage();
      } catch (KeyManagementException kme) {
    	  return kme.getMessage();
      } catch (NoSuchAlgorithmException nsae) {
    	  return nsae.getMessage();
      }
   }

   static synchronized SSLSocketFactory prepFactory(@SuppressWarnings("unused") HttpsURLConnection httpsConnection) 
            throws NoSuchAlgorithmException, KeyStoreException, KeyManagementException {

      if (factory == null) {
         SSLContext ctx = SSLContext.getInstance("TLS");
         ctx.init(null, new TrustManager[]{ new AlwaysTrustManager() }, null);
         factory = ctx.getSocketFactory();
      }
      return factory;
   }
   
   private static final class TrustingHostnameVerifier implements HostnameVerifier {
	   @Override
      public boolean verify(@SuppressWarnings("unused") String hostname, @SuppressWarnings("unused") SSLSession session) {
         return true;
      }
   }

   private static class AlwaysTrustManager implements X509TrustManager {
	   @Override
      public void checkClientTrusted(@SuppressWarnings("unused") X509Certificate[] arg0, @SuppressWarnings("unused") String arg1) throws CertificateException { }
	   @Override
      public void checkServerTrusted(@SuppressWarnings("unused") X509Certificate[] arg0, @SuppressWarnings("unused") String arg1) throws CertificateException { }
	   @Override
      public X509Certificate[] getAcceptedIssuers() { return new X509Certificate[0]; }      
   }
   
}