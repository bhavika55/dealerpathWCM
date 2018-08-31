
/* 
 * Project    : DealerPath
 * Script     : LogFactory
 * Author     : Shrishail Baddi
 * Date       : April.14.2018
 * Last Modified On:
 * Modified By :
 */


package com.deere.Helpers;

import org.apache.log4j.Logger;

public class LogFactory {

	// Initialize Log4j logs

	private static Logger Log = Logger.getLogger(LogFactory.class.getName());//

	// This is to print log for the beginning of the test case, as we usually run so
	// many test cases as a test suite

	public static void beginTestCase(String strTestCaseName) {

		Log.info("****************************************************************************************");

		Log.info(strTestCaseName);

		Log.info("****************************************************************************************");

	}

	// This is to print log for the ending of the test case

	public static void endTestCase( String strTestCaseName) {

		Log.info("xxxxxxxxxxxxxxxxxxxxxxxx  " + strTestCaseName + "E---N---D-" + "  xxxxxxxxxxxxxxxxxxxxxxxxxx");

	}

	// Need to create these methods, so that they can be called

	public static void info(String strMessage) {

		if ( BaseClass.ENABLE_LOGS.equalsIgnoreCase("ON"))
		Log.info(strMessage.trim());

	}

	public static void warn(String strMessage) {
		
		if ( BaseClass.ENABLE_LOGS.equalsIgnoreCase("ON"))
		Log.warn(strMessage);

	}

	public static void error(String strMessage) {
		
		if ( BaseClass.ENABLE_LOGS.equalsIgnoreCase("ON"))
		Log.error(strMessage);

	}

	public static void fatal(String strMessage) {

		if ( BaseClass.ENABLE_LOGS.equalsIgnoreCase("ON"))
		Log.fatal(strMessage);

	}

	public static void debug(String strMessage) {
		if ( BaseClass.ENABLE_LOGS.equalsIgnoreCase("ON"))
		Log.debug(strMessage);

	}

}
