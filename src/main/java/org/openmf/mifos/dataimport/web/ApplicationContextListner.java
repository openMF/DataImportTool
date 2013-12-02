package org.openmf.mifos.dataimport.web;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import javax.servlet.ServletContextEvent;
import javax.servlet.ServletContextListener;
import javax.servlet.annotation.WebListener;

import org.openmf.mifos.dataimport.utils.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@WebListener
public class ApplicationContextListner implements ServletContextListener {


	private static final Logger logger = LoggerFactory.getLogger(ApplicationContextListner.class);

	@Override
	public void contextInitialized(@SuppressWarnings("unused") ServletContextEvent sce) {
		Properties prop = new Properties();
	    String homeDirectory = System.getProperty("user.home");
		FileInputStream fis = null;
		try {
			File file = new File(homeDirectory + "/dataimport.properties");
			fis = new FileInputStream(file);
			prop.load(fis);
			readAndSetAsSysProp("mifos.endpoint", "http://localhost:8080", prop);
			readAndSetAsSysProp("mifos.user.id", "mifos", prop);
			readAndSetAsSysProp("mifos.password", "testmifos", prop);
			readAndSetAsSysProp("mifos.tenant.id", "default", prop);
			fis.close();
		} catch (IOException e) {
			logger.error("Unable to find dataimport.properties", e);
			throw new IllegalStateException(e);
		}
	}


	private void readAndSetAsSysProp(String key, String defaultValue, Properties prop) {
		String value = prop.getProperty(key);
		if(!StringUtils.isBlank(value)) {
			System.setProperty(key, value);
		} else {
			System.setProperty(key, defaultValue);
		}
	}


	@Override
	public void contextDestroyed(@SuppressWarnings("unused") ServletContextEvent sce) {
		// TODO Auto-generated method stub
	} 
}
