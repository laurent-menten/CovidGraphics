package lo.utils;

import java.net.ConnectException;

// Lo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, February 2015

/* A growing collection of utility functions to make Office
   easier to use. They are currently divided into the following
   groups:

     * interface object creation (uses generics)

     * office starting
     * office shutdown

     * document opening
     * document creation
     * document saving
     * document closing

     * initialization via Addon-supplied context
     * initialization via script context

     * dispatch
     * UNO cmds

     * use Inspectors extension

     * color methods
     * other utils

     * container manipulation
*/

import java.net.URLClassLoader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XIntrospection;
import com.sun.star.beans.XIntrospectionAccess;
import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XBridge;
import com.sun.star.bridge.XBridgeFactory;
import com.sun.star.comp.beans.OOoBean;
import com.sun.star.comp.beans.OfficeConnection;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.connection.XConnection;
import com.sun.star.connection.XConnector;
import com.sun.star.container.XChild;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNamed;
import com.sun.star.document.MacroExecMode;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.reflection.XIdlMethod;
import com.sun.star.script.provider.XScriptContext;
import com.sun.star.uno.Exception;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;

public class Lo
{
	// docType ints
	public static final int UNKNOWN = 0;
	public static final int WRITER = 1;
	public static final int BASE = 2;
	public static final int CALC = 3;
	public static final int DRAW = 4;
	public static final int IMPRESS = 5;
	public static final int MATH = 6;

	// docType strings
	public static final String UNKNOWN_STR = "unknown";
	public static final String WRITER_STR = "swriter";
	public static final String BASE_STR = "sbase";
	public static final String CALC_STR = "scalc";
	public static final String DRAW_STR = "sdraw";
	public static final String IMPRESS_STR = "simpress";
	public static final String MATH_STR = "smath";

	// docType service names
	public static final String UNKNOWN_SERVICE = "com.sun.frame.XModel";
	public static final String WRITER_SERVICE = "com.sun.star.text.TextDocument";
	public static final String BASE_SERVICE = "com.sun.star.sdb.OfficeDatabaseDocument";
	public static final String CALC_SERVICE = "com.sun.star.sheet.SpreadsheetDocument";
	public static final String DRAW_SERVICE = "com.sun.star.drawing.DrawingDocument";
	public static final String IMPRESS_SERVICE = "com.sun.star.presentation.PresentationDocument";
	public static final String MATH_SERVICE = "com.sun.star.formula.FormulaProperties";

	// connect to locally running Office via port 8100
	private static final int SOCKET_PORT = 8100;

	// CLSIDs for Office documents
	// defined in
	// <OFFICE>\officecfg\registry\data\org\openoffice\Office\Embedding.xcu
	public static final String WRITER_CLSID = "8BC6B165-B1B2-4EDD-aa47-dae2ee689dd6";
	public static final String CALC_CLSID = "47BBB4CB-CE4C-4E80-a591-42d9ae74950f";
	public static final String DRAW_CLSID = "4BAB8970-8A3B-45B3-991c-cbeeac6bd5e3";
	public static final String IMPRESS_CLSID = "9176E48A-637A-4D1F-803b-99d9bfac1047";
	public static final String MATH_CLSID = "078B7ABA-54FC-457F-8551-6147e776a997";
	public static final String CHART_CLSID = "12DCAE26-281F-416F-a234-c3086127382e";

	/*
	 * unsure about these: chart2 "80243D39-6741-46C5-926E-069164FF87BB" service:
	 * com.sun.star.chart2.ChartDocument
	 * 
	 * applet "970B1E81-CF2D-11CF-89CA-008029E4B0B1" service:
	 * com.sun.star.comp.sfx2.AppletObject
	 * 
	 * plug-in "4CAA7761-6B8B-11CF-89CA-008029E4B0B1" service:
	 * com.sun.star.comp.sfx2.PluginObject
	 * 
	 * frame "1A8A6701-DE58-11CF-89CA-008029E4B0B1" service:
	 * com.sun.star.comp.sfx2.IFrameObject
	 * 
	 * XML report chart "D7896D52-B7AF-4820-9DFE-D404D015960F" service:
	 * com.sun.star.report.ReportDefinition
	 * 
	 */

	public static final int TERMINATION_ATTEMPTS = 4;

	// remote component context
	private static XComponentContext xcc = null;

	// remote desktop UNO service
	private static XDesktop xDesktop = null;

	// remote service managers
	private static XMultiComponentFactory mcFactory = null;
	// has replaced XMultiServiceFactory
	private static XMultiServiceFactory msFactory = null;

	private static XComponent bridgeComponent = null;
	// this is only set if office is opened via a socket

	private static boolean isOfficeTerminated = false;

	// ========================================================================
	// = 
	// ========================================================================

	public static XComponentContext getContext()
	{
		return xcc;
	}

	public static XDesktop getDesktop()
	{
		return xDesktop;
	}

	public static XMultiComponentFactory getComponentFactory()
	{
		return mcFactory;
	}

	public static XMultiServiceFactory getServiceFactory()
	{
		return msFactory;
	}

	public static XComponent getBridge()
	{
		return bridgeComponent;
	}

	// ========================================================================
	// = 
	// ========================================================================

	// use OOoBean to initialize Lo globals

	@SuppressWarnings("deprecation")
	public static void setOOoBean( OOoBean oob )
	{
		try
		{
			OfficeConnection conn = oob.getOOoConnection();
			if( conn == null )
			{
				LOG.severe( "No office connection found in OOoBean" );
			}
			else
			{
				xcc = conn.getComponentContext();
				if( xcc == null )
				{
					LOG.severe( "No component context found in OOoBean" );
				}
				else
				{
					mcFactory = xcc.getServiceManager();
				}

				xDesktop = oob.getOOoDesktop();
				msFactory = oob.getMultiServiceFactory();
			}
		}
		catch( java.lang.Exception ex )
		{
			LOG.log( Level.SEVERE, "Couldn't initialize LO using OOoBean ", ex );
		}
	}

	// ========================================================================
	// = Interface object creation (uses generics) ============================
	// ========================================================================

	// the "Loki" function -- reduces typing
	public static <T> T qi( Class<T> aType, Object o )
	{
		return UnoRuntime.queryInterface(aType, o);
	}

	// retrieves the parent of the given object
	public static <T> T getParent( Object aComponent, Class<T> aType )
	{
		XChild xAsChild = Lo.qi( XChild.class, aComponent );

		return Lo.qi( aType, xAsChild.getParent() );
	}

	// ------------------------------------------------------------------------

	/*
	 * create an interface object of class aType from the named service; uses 'old'
	 * XMultiServiceFactory, so a document must have already been loaded/created
	 */
	public static <T> T createInstanceMSF( Class<T> aType, String serviceName )
	{
		if( msFactory == null )
		{
			LOG.warning( "No document found" );
			return null;
		}

		T interfaceObj = null;
		try
		{
			// uses bridge to obtain proxy to remote interface inside service;
			// implements casting across process boundaries

			Object o = msFactory.createInstance( serviceName );
			interfaceObj = Lo.qi( aType, o );
		}
		catch( Exception ex )
		{
			LOG.log( Level.SEVERE, "Couldn't create interface for \"" + serviceName + "\" ", ex );
		}

		return interfaceObj;
	}

	/*
	 * create an interface object of class aType from the named service; uses 'old'
	 * XMultiServiceFactory, so a document must have been already loaded/created
	 */
	public static <T> T createInstanceMSF( Class<T> aType, String serviceName, XMultiServiceFactory msf )
	{
		if( msf == null )
		{
			LOG.warning( "No document found" );
			return null;
		}

		T interfaceObj = null;
		try
		{
			// uses bridge to obtain proxy to remote interface inside service;
			// implements casting across process boundaries
	
			Object o = msf.createInstance( serviceName );
			interfaceObj = Lo.qi( aType, o );
		}
		catch( Exception ex )
		{
			LOG.log( Level.SEVERE, "Couldn't create interface for \"" + serviceName + "\" ", ex );
		}

		return interfaceObj;
	}

	/*
	 * create an interface object of class aType from the named service; uses
	 * XComponentContext and 'new' XMultiComponentFactory so only a bridge to office
	 * is needed
	 */
	public static <T> T createInstanceMCF( Class<T> aType, String serviceName )
	{
		if( (xcc == null) || (mcFactory == null) )
		{
			LOG.warning( "No office connection found" );
			return null;
		}

		T interfaceObj = null;
		try
		{
			// uses bridge to obtain proxy to remote interface inside service;
			// implements casting across process boundaries

			Object o = mcFactory.createInstanceWithContext( serviceName, xcc );
			interfaceObj = Lo.qi( aType, o );
		}
		catch( Exception ex )
		{
			LOG.log( Level.SEVERE, "Couldn't create interface for \"" + serviceName + "\" ", ex );
		}

		return interfaceObj;
	}

	/*
	 * create an interface object of class aType from the named service and
	 * arguments; uses XComponentContext and 'new' XMultiComponentFactory so only a
	 * bridge to office is needed
	 */
	public static <T> T createInstanceMCF( Class<T> aType, String serviceName, Object[] args )
	{
		if( (xcc == null) || (mcFactory == null) )
		{
			LOG.warning( "No office connection found" );
			return null;
		}

		T interfaceObj = null;
		try
		{
			// uses bridge to obtain proxy to remote interface inside service;
			// implements casting across process boundaries

			Object o = mcFactory.createInstanceWithArgumentsAndContext( serviceName, args, xcc );
			interfaceObj = Lo.qi(aType, o);
		}
		catch( Exception ex )
		{
			LOG.log( Level.SEVERE, "Couldn't create interface for \"" + serviceName + "\" ", ex );
		}

		return interfaceObj;
	}

	// ========================================================================
	// = Load libre office ==================================================== 
	// ========================================================================

	// default is to using office via pipes
	public static XComponentLoader loadOffice()
	{
		return loadOffice( false );
	}

	public static XComponentLoader loadSocketOffice()
	{
		return loadOffice( false );
	}

	public static XComponentLoader loadPipeOffice()
	{
		return loadOffice( true );
	}

	/*
	 * Creation sequence: remote component content (xcc) --> remote service manager
	 * (mcFactory) --> remote desktop (xDesktop) --> component loader
	 * (XComponentLoader) Once we have a component loader, we can load a document.
	 * xcc, mcFactory, and xDesktop are stored as static globals.
	 */
	private static XComponentLoader loadOffice( boolean usingPipes )
	{
		LOG.info( "Loading Office ..." );

		if( usingPipes )
		{
			LOG.fine( "Connect using pipe" );

			xcc = bootstrapContext();
		}
		else
		{
			LOG.fine( "Connect using socket" );

			xcc = socketContext();
		}

		if( xcc == null )
		{
			LOG.severe( "Office context could not be created" );
			return null;
		}

		// get the remote office service manager
		mcFactory = xcc.getServiceManager();
		if( mcFactory == null )
		{
			LOG.severe( "Office Service Manager is unavailable" );
			return null;
		}

		// desktop service handles application windows and documents
		xDesktop = createInstanceMCF( XDesktop.class, "com.sun.star.frame.Desktop" );
		if( xDesktop == null )
		{
			LOG.severe( "Could not create a desktop service" );
			return null;
		}

		XComponentLoader ldr = Lo.qi( XComponentLoader.class, xDesktop );

		// --------------------------------------------------------------------

		LOG.info( "Office name : " + Info.getConfig( "ooName" ) );
		LOG.info( "Office version : " + Info.getConfig( "ooSetupVersionAboutBox" ) );
		LOG.info( "Office directory : " + Info.getOfficeDir() );

		// --------------------------------------------------------------------

		return ldr;
	}

	// ========================================================================
	// = 
	// ========================================================================

	// connect pipes to office using the Bootstrap class
	// i.e. see code at
	// http://svn.apache.org/repos/asf/openoffice/symphony/trunk/main/
	// javaunohelper/com/sun/star/comp/helper/Bootstrap.java

	private static XComponentContext bootstrapContext()
	{
		XComponentContext xcc = null;

		try
		{
			xcc = Bootstrap.bootstrap();
		}
		catch( BootstrapException ex )
		{
			LOG.log( Level.SEVERE, "Unable to bootstrap Office", ex );
		}

		return xcc;
	}

	// use socket connection to Office
	// https://forum.openoffice.org/en/forum/viewtopic.php?f=44&t=1014
	// requires soffice to be in Windows PATH env var.

	private static XComponentContext socketContext()
	{
		XComponentContext xcc = null;

		try
		{
			String [] cmdArray = new String[3];
			if( System.getProperty( "os.name" ).startsWith( "Windows" ) )
			{
				cmdArray[0] = "soffice.exe";
			}
			else
			{
				cmdArray[0] = "soffice";
			}
			cmdArray[1] = "-headless";
			cmdArray[2] = "-accept=socket,host=localhost,port=" + SOCKET_PORT + ";urp;";
			
			Process p = Runtime.getRuntime().exec( cmdArray );
			if( p != null )
			{
				LOG.info( "Office process sucessfully created" );
			}

			// ----------------------------------------------------------------

			XComponentContext localContext
				= Bootstrap.createInitialComponentContext( null );

			XMultiComponentFactory localFactory
				= localContext.getServiceManager();

			XConnector connector
				= Lo.qi(
						XConnector.class,
						localFactory.createInstanceWithContext(
								"com.sun.star.connection.Connector",
								localContext
							)
					);

			XConnection connection = null;

			LOG.fine( "Waiting for listening mode..." );
			for( int i = 0 ; i <= 30 ; i++ )
			{
				LOG.finer( "Connection attemps #" + i );

				try
				{
					connection = connector.connect( "socket,host=localhost,port=" + SOCKET_PORT );
				}
				catch( com.sun.star.connection.NoConnectException ex )
				{
					if( i == 30 )
					{
						throw ex;
					}
				}

				if( connection != null )
				{
					break;
				}

				delay( 1* 1000 );
			}

			LOG.fine( "Connection to LibreOffice established" );

			// ----------------------------------------------------------------

			XBridgeFactory bridgeFactory
				= Lo.qi(
						XBridgeFactory.class,
						localFactory.createInstanceWithContext(
								"com.sun.star.bridge.BridgeFactory",
								localContext
							)
					);

			XBridge bridge
				= bridgeFactory.createBridge(
						"socketBridgeAD",
						"urp",
						connection,
						null
					);

			bridgeComponent = Lo.qi( XComponent.class, bridge );

			XMultiComponentFactory serviceManager
				= Lo.qi(
						XMultiComponentFactory.class,
						bridge.getInstance( "StarOffice.ServiceManager" )
					);

			XPropertySet props = Lo.qi( XPropertySet.class, serviceManager );
			
			Object defaultContext = props.getPropertyValue( "DefaultContext" );

			xcc = Lo.qi( XComponentContext.class, defaultContext );
		}
		catch( java.lang.Exception ex )
		{
			LOG.log( Level.SEVERE, "Unable to connect to LibreOffice socket", ex );
		}

		return xcc;
	}

	// ========================================================================
	// = Close Libre Office ===================================================
	// ========================================================================

	// tell office to terminate
	public static void closeOffice()
	{
		LOG.info( "Closing Office" );

		if( xDesktop == null )
		{
			LOG.warning( "No office connection found" );
			return;
		}

		if( isOfficeTerminated )
		{
			LOG.warning( "Office has already been requested to terminate" );
			return;
		}

		int numTries = 1;
		while( !isOfficeTerminated && (numTries < TERMINATION_ATTEMPTS) )
		{
			delay( 200 );

			isOfficeTerminated = tryToTerminate( numTries );
			numTries++;
		}

		if( numTries >= TERMINATION_ATTEMPTS )
		{
			LOG.warning( "Failed to terminate office after " + numTries + " attempts" );
		}
	}

	public static boolean tryToTerminate( int numTries )
	{
		try
		{
			boolean isDead = xDesktop.terminate();
			if( isDead )
			{
				if (numTries > 1)
				{
					LOG.fine( numTries + ". Office terminated" );
				}
				else
				{
					LOG.info( "Office terminated" );
				}
			}
			else
			{
				LOG.warning( numTries + ". Office failed to terminate");
			}

			return isDead;
		}
		catch( com.sun.star.lang.DisposedException ex )
		{
			LOG.log( Level.WARNING, "Office link disposed", ex );
			return true;
		}
		catch( java.lang.Exception ex )
		{
			LOG.log( Level.SEVERE, "Termination exception", ex );
			return false;
		}
	}

	// kill office processes using a batch file
	// or use JNAUtils.killOffice()

	public static void killOffice()
	{
		JNAUtils.killOffice();

//		try
//		{
//			Runtime.getRuntime().exec( "cmd /c lokill.bat" );
//			log.info( "Killed Office" );
//		}
//		catch( java.lang.Exception ex )
//		{
//			log.log( Level.SEVERE, "Unable to kill Office", ex );
//		}
	}
	
	// ======================== document opening ==============

	public static XComponent openFlatDoc( String fnm, String docType, XComponentLoader loader )
	{
		String nm = XML.getFlatFilterName(docType);
		System.out.println("Flat filter Name: " + nm);
		return openDoc(fnm, loader, Props.makeProps("FilterName", nm));
	}

	public static XComponent openDoc( String fnm, XComponentLoader loader )
	{
		return openDoc( fnm, loader, Props.makeProps( "Hidden", true ) );
	}

	public static XComponent openReadOnlyDoc( String fnm, XComponentLoader loader )
	{
		return openDoc( fnm, loader, Props.makeProps( "Hidden", true, "ReadOnly", true ) );
	}

	// open the specified document
	// the possibly props for a document are listed in the MediaDescriptor service
	public static XComponent openDoc( String fnm, XComponentLoader loader, PropertyValue[] props )
	{
		if( fnm == null )
		{
			LOG.warning( "filename is null" );
			return null;
		}

		String openFileURL = null;
		if( !FileIO.isOpenable( fnm ) )
		{
			if( isURL( fnm ) )
			{
				LOG.info( "Will treat filename as an URL: \"" + fnm + "\"" );
				openFileURL = fnm;
			}
			else
			{
				LOG.warning( "Filename is not an URL: \"" + fnm + "\"" );
				return null;
			}
		}
		else
		{
			LOG.fine( "Opening file \"" + fnm + "\"" );

			openFileURL = FileIO.fnmToURL( fnm );
			if( openFileURL == null )
			{
				return null;
			}
		}

		XComponent doc = null;
		try
		{
			doc = loader.loadComponentFromURL(openFileURL, "_blank", 0, props);
			msFactory = Lo.qi(XMultiServiceFactory.class, doc);
		}
		catch( Exception ex )
		{
			LOG.log( Level.SEVERE, "Unable to open the document \"" + fnm + "\"", ex);
		}

		return doc;
	}

	public static boolean isURL( String fnm )
	{
		try
		{
			java.net.URL u = new java.net.URL( fnm );
			u.toURI();

			return true;
		}
		catch( java.net.MalformedURLException ex )
		{
			return false;
		}
		catch( java.net.URISyntaxException ex )
		{
			return false;
		}
	} // end of isURL()

	// ======================== document creation ==============

	public static String ext2DocType(String ext) {
		switch (ext) {
		case "odt":
			return WRITER_STR;
		case "odp":
			return IMPRESS_STR;
		case "odg":
			return DRAW_STR;
		case "ods":
			return CALC_STR;
		case "odb":
			return BASE_STR;
		case "odf":
			return MATH_STR;
		default:
			System.out.println("Do not recognize extension \"" + ext + "\"; using writer");
			return WRITER_STR; // could use UNKNOWN_STR
		}
	} // end of ext2DocType()

	/*
	 * docType (without the ""private:factory/") is: "private:factory/swriter"
	 * Writer document "private:factory/simpress" Impress presentation document
	 * "private:factory/sdraw" Draw document "private:factory/scalc" Calc document
	 * "private:factory/sdatabase" Base document "private:factory/smath" Math
	 * formula document
	 * 
	 * These are not handled: "private:factory/schart" Chart
	 * "private:factory/swriter/web" Writer HTML Web document
	 * "private:factory/swriter/GlobalDocument" Master document
	 * 
	 * ".component:Bibliography/View1" Bibliography-Edit the bibliography entries
	 * 
	 * ".component:DB/QueryDesign" Database comp ".component:DB/TableDesign"
	 * ".component:DB/RelationDesign" ".component:DB/DataSourceBrowser"
	 * ".component:DB/FormGridView"
	 */

	public static String docTypeStr(int docTypeVal) {
		switch (docTypeVal) {
		case WRITER:
			return WRITER_STR;
		case IMPRESS:
			return IMPRESS_STR;
		case DRAW:
			return DRAW_STR;
		case CALC:
			return CALC_STR;
		case BASE:
			return BASE_STR;
		case MATH:
			return MATH_STR;
		default:
			System.out.println("Do not recognize extension \"" + docTypeVal + "\"; using writer");
			return WRITER_STR; // could use UNKNOWN_STR
		}
	} // end of docTypeStr()

	public static XComponent createDoc(String docType, XComponentLoader loader) {
		return createDoc(docType, loader, Props.makeProps("Hidden", true));
	}

	public static XComponent createMacroDoc(String docType, XComponentLoader loader) {
		return createDoc(docType, loader, Props.makeProps("Hidden", false,
				// "MacroExecutionMode", MacroExecMode.ALWAYS_EXECUTE) ); }
				"MacroExecutionMode", MacroExecMode.ALWAYS_EXECUTE_NO_WARN));
	}

	public static XComponent createDoc(String docType, XComponentLoader loader, PropertyValue[] props)
	// create a new document of the specified type
	{
		System.out.println("Creating Office document " + docType);
		// PropertyValue[] props = Props.makeProps("Hidden", true);
		// if Hidden == true, office will not terminate properly
		XComponent doc = null;
		try {
			doc = loader.loadComponentFromURL("private:factory/" + docType, "_blank", 0, props);
			msFactory = Lo.qi(XMultiServiceFactory.class, doc);
		} catch (Exception e) {
			System.out.println("Could not create a document");
		}
		return doc;
	} // end of createDoc()

	public static XComponent createDocFromTemplate(String templatePath, XComponentLoader loader)
	// create a new document using the specified template
	{
		if (!FileIO.isOpenable(templatePath))
			return null;
		System.out.println("Opening template " + templatePath);
		String templateURL = FileIO.fnmToURL(templatePath);
		if (templateURL == null)
			return null;

		PropertyValue[] props = Props.makeProps("Hidden", true, "AsTemplate", true);
		XComponent doc = null;
		try {
			doc = loader.loadComponentFromURL(templateURL, "_blank", 0, props);
			msFactory = Lo.qi(XMultiServiceFactory.class, doc);
		} catch (Exception e) {
			System.out.println("Could not create document from template: " + e);
		}
		return doc;
	} // end of createDocFromTemplate()

	// ======================== document saving ==============

	public static void save(Object odoc)
	// was XComponent
	{
		XStorable store = Lo.qi(XStorable.class, odoc);
		try {
			store.store();
			System.out.println("Saved the document by overwriting");
		} catch (IOException e) {
			System.out.println("Could not save the document");
		}
	} // end of save()

	public static void saveDoc(Object odoc, String fnm)
	// was XComponent
	{
		XStorable store = Lo.qi(XStorable.class, odoc);
		XComponent doc = Lo.qi(XComponent.class, odoc);
		int docType = Info.reportDocType(doc);
		storeDoc(store, docType, fnm, null); // no password
	}

	public static void saveDoc(Object odoc, String fnm, String password)
	// was XComponent
	{
		XStorable store = Lo.qi(XStorable.class, odoc);
		XComponent doc = Lo.qi(XComponent.class, odoc);
		int docType = Info.reportDocType(doc);
		storeDoc(store, docType, fnm, password);
	}

	public static void saveDoc(Object odoc, String fnm, String format, String password)
	// was XComponent
	{
		XStorable store = Lo.qi(XStorable.class, odoc);
		storeDocFormat(store, fnm, format, password);
	}

	// public static void storeDoc(XStorable store, int docType, String fnm)
	// { saveDoc(store, docType, fnm, null); } // no password

	public static void storeDoc(XStorable store, int docType, String fnm, String password)
	// Save the document using the file's extension as a guide.
	{
		String ext = Info.getExt(fnm);
		String format = "Text";
		if (ext == null)
			System.out.println("Assuming a text format");
		else
			format = ext2Format(docType, ext);
		storeDocFormat(store, fnm, format, password);
	} // end of storeDoc()

	public static String ext2Format(String ext) {
		return ext2Format(Lo.UNKNOWN, ext);
	}

	public static String ext2Format(int docType, String ext)
	/*
	 * convert the extension string into a suitable office format string. The
	 * formats were chosen based on the fact that they are being used to save (or
	 * export) a document.
	 * 
	 * The names were obtained from
	 * http://www.oooforum.org/forum/viewtopic.phtml?t=71294 and by running my
	 * DocInfo application on files.
	 * 
	 * I use the docType to distinguish between the various meanings of the PDF ext.
	 * 
	 * This could be a lot more extensive.
	 * 
	 * Use Info.getFilterNames() to get the filter names for your Office (the last
	 * time I tried it, I got back 246 names!)
	 */
	{
		switch (ext) {
		case "doc":
			return "MS Word 97";
		case "docx":
			return "Office Open XML Text"; // "MS Word 2007 XML"
		case "rtf":
			if (docType == Lo.CALC)
				return "Rich Text Format (StarCalc)";
			else
				return "Rich Text Format"; // assume writer

		case "odt":
			return "writer8";
		case "ott":
			return "writer8_template";

		case "pdf":
			if (docType == Lo.WRITER)
				return "writer_pdf_Export";
			else if (docType == Lo.IMPRESS)
				return "impress_pdf_Export";
			else if (docType == Lo.DRAW)
				return "draw_pdf_Export";
			else if (docType == Lo.CALC)
				return "calc_pdf_Export";
			else if (docType == Lo.MATH)
				return "math_pdf_Export";
			else
				return "writer_pdf_Export"; // assume we are saving a writer doc

		case "txt":
			return "Text";

		case "ppt":
			return "MS PowerPoint 97";
		case "pptx":
			return "Impress MS PowerPoint 2007 XML";
		case "odp":
			return "impress8";
		case "odg":
			return "draw8";

		case "jpg":
			if (docType == Lo.IMPRESS)
				return "impress_jpg_Export";
			else
				return "draw_jpg_Export"; // assume Draw doc

		case "png":
			if (docType == Lo.IMPRESS)
				return "impress_png_Export";
			else
				return "draw_png_Export"; // assume Draw doc

		case "xls":
			return "MS Excel 97";
		case "xlsx":
			return "Calc MS Excel 2007 XML";
		case "csv":
			return "Text - txt - csv (StarCalc)"; // "Text CSV";
		case "ods":
			return "calc8";
		case "odb":
			return "StarOffice XML (Base)";

		case "htm":
		case "html":
			if (docType == Lo.WRITER)
				return "HTML (StarWriter)"; // "writerglobal8_HTML";
			else if (docType == Lo.IMPRESS)
				return "impress_html_Export";
			else if (docType == Lo.DRAW)
				return "draw_html_Export";
			else if (docType == Lo.CALC)
				return "HTML (StarCalc)";
			else
				return "HTML";

		case "xhtml":
			if (docType == Lo.WRITER)
				return "XHTML Writer File";
			else if (docType == Lo.IMPRESS)
				return "XHTML Impress File";
			else if (docType == Lo.DRAW)
				return "XHTML Draw File";
			else if (docType == Lo.CALC)
				return "XHTML Calc File";
			else
				return "XHTML Writer File"; // assume we are saving a writer doc

		case "xml":
			if (docType == Lo.WRITER)
				return "OpenDocument Text Flat XML";
			else if (docType == Lo.IMPRESS)
				return "OpenDocument Presentation Flat XML";
			else if (docType == Lo.DRAW)
				return "OpenDocument Drawing Flat XML";
			else if (docType == Lo.CALC)
				return "OpenDocument Spreadsheet Flat XML";
			else
				return "OpenDocument Text Flat XML"; // assume we are saving a writer doc

		default: // assume user means text
			System.out.println("Do not recognize extension \"" + ext + "\"; using text");
			return "Text";
		}
	} // end of ext2Format()

	public static void storeDocFormat(XStorable store, String fnm, String format, String password)
	// save the document in the specified file using the supplied office format
	{
		System.out.println("Saving the document in " + fnm);
		System.out.println("Using format: " + format);
		try {
			String saveFileURL = FileIO.fnmToURL(fnm);
			if (saveFileURL == null)
				return;

			PropertyValue[] storeProps;
			if (password == null) // no password supplied
				storeProps = Props.makeProps("Overwrite", true, "FilterName", format);
			else {
				String[] nms = new String[] { "Overwrite", "FilterName", "Password" };
				Object[] vals = new Object[] { true, format, password };
				storeProps = Props.makeProps(nms, vals);
			}
			store.storeToURL(saveFileURL, storeProps);
		} catch (IOException e) {
			System.out.println("Could not save " + fnm + ": " + e);
		}
	} // end of storeDocFormat()

	// ======================== document closing ==============

	public static void closeDoc(Object doc)
	// was XComponent
	{
		try {
			XCloseable closeable = Lo.qi(XCloseable.class, doc);
			close(closeable);
		} catch (com.sun.star.lang.DisposedException e) {
			LOG.severe("Document close failed since Office link disposed");
		}
	}

	public static void close(XCloseable closeable) {
		if (closeable == null)
			return;
		LOG.info("Closing the document");
		try {
			closeable.close(false); // true to force a close
			// set modifiable to false to close a modified doc without complaint
			// setModified(False)
		} catch (CloseVetoException e) {
			LOG.warning("Close was vetoed");
		}
	} // end of close()

	// ================= initialization via Addon-supplied context
	// ====================

	public static XComponent addonInitialize(XComponentContext addonXcc) {
		xcc = addonXcc;
		if (xcc == null) {
			System.out.println("Could not access component context");
			return null;
		}

		mcFactory = xcc.getServiceManager();
		if (mcFactory == null) {
			System.out.println("Office Service Manager is unavailable");
			return null;
		}

		try {
			Object oDesktop = mcFactory.createInstanceWithContext("com.sun.star.frame.Desktop", xcc);
			xDesktop = Lo.qi(XDesktop.class, oDesktop);
		} catch (Exception e) {
			System.out.println("Could not access desktop");
			return null;
		}

		XComponent doc = xDesktop.getCurrentComponent();
		if (doc == null) {
			System.out.println("Could not access document");
			return null;
		}

		msFactory = Lo.qi(XMultiServiceFactory.class, doc);
		return doc;
	} // end of addonInitialize()

	// ============= initialization via script context ======================

	public static XComponent scriptInitialize(XScriptContext sc) {
		if (sc == null) {
			System.out.println("Script Context is null");
			return null;
		}

		xcc = sc.getComponentContext();
		if (xcc == null) {
			System.out.println("Could not access component context");
			return null;
		}
		mcFactory = xcc.getServiceManager();
		if (mcFactory == null) {
			System.out.println("Office Service Manager is unavailable");
			return null;
		}

		xDesktop = sc.getDesktop();
		if (xDesktop == null) {
			System.out.println("Could not access desktop");
			return null;
		}

		XComponent doc = xDesktop.getCurrentComponent();
		if (doc == null) {
			System.out.println("Could not access document");
			return null;
		}

		msFactory = Lo.qi(XMultiServiceFactory.class, doc);
		return doc;
	} // end of scriptInitialize()

	// ==================== dispatch ===============================
	// see https://wiki.documentfoundation.org/Development/DispatchCommands

	public static boolean dispatchCmd(String cmd) {
		return dispatchCmd(xDesktop.getCurrentFrame(), cmd, null);
	}

	public static boolean dispatchCmd(String cmd, PropertyValue[] props) {
		return dispatchCmd(xDesktop.getCurrentFrame(), cmd, props);
	}

	public static boolean dispatchCmd(XFrame frame, String cmd, PropertyValue[] props)
	// cmd does not include the ".uno:" substring; e.g. pass "Zoom" not ".uno:Zoom"
	{
		XDispatchHelper helper = createInstanceMCF(XDispatchHelper.class, "com.sun.star.frame.DispatchHelper");
		if (helper == null) {
			System.out.println("Could not create dispatch helper for command " + cmd);
			return false;
		}

		try {
			XDispatchProvider provider = Lo.qi(XDispatchProvider.class, frame);

			/*
			 * returns failure even when the event works (?), and an illegal value when the
			 * dispatch actually does fail
			 */
			/*
			 * DispatchResultEvent res = (DispatchResultEvent)
			 * helper.executeDispatch(provider, (".uno:" + cmd), "", 0, props); if
			 * (res.State == DispatchResultState.FAILURE)
			 * System.out.println("Dispatch failed for \"" + cmd + "\""); else if (res.State
			 * == DispatchResultState.DONTKNOW)
			 * System.out.println("Dispatch result unknown for \"" + cmd + "\"");
			 */
			helper.executeDispatch(provider, (".uno:" + cmd), "", 0, props);
			return true;
		} catch (java.lang.Exception e) {
			System.out.println("Could not dispatch \"" + cmd + "\":\n  " + e);
		}
		return false;
	} // end of dispatchCmd()

	// ================= Uno cmds =========================

	public static String makeUnoCmd(String itemName)
	// use a dummy Java class name, Foo
	{
		return "vnd.sun.star.script:Foo/Foo." + itemName + "?language=Java&location=share";
	}

	public static String extractItemName(String unoCmd)
	/*
	 * format is: "vnd.sun.star.script:Foo/Foo." + itemName +
	 * "?language=Java&location=share";
	 */
	{
		int fooPos = unoCmd.indexOf("Foo.");
		if (fooPos == -1) {
			System.out.println("Could not find Foo header in command: \"" + unoCmd + "\"");
			return null;
		}

		int langPos = unoCmd.indexOf("?language");
		if (langPos == -1) {
			System.out.println("Could not find language header in command: \"" + unoCmd + "\"");
			return null;
		}

		return unoCmd.substring(fooPos + 4, langPos);
	} // end of extractItemName()

	// ======================== use Inspector extensions ====================

	public static void inspect(Object obj)
	/*
	 * call XInspector.inspect() in the Inspector.oxt extension Available from
	 * https://wiki.openoffice.org/wiki/Object_Inspector
	 */
	{
		if ((xcc == null) || (mcFactory == null)) {
			System.out.println("No office connection found");
			return;
		}

		try {
			Type[] ts = Info.getInterfaceTypes(obj); // get class name for title
			String title = "Object";
			if ((ts != null) && (ts.length > 0))
				title = ts[0].getTypeName() + " " + title;

			Object inspector = mcFactory.createInstanceWithContext("org.openoffice.InstanceInspector", xcc);
			// hangs on second use
			if (inspector == null) {
				System.out.println("Inspector Service could not be instantiated");
				return;
			}

			System.out.println("Inspector Service instantiated");
			/*
			 * // report on inspector XServiceInfo si = Lo.qi(XServiceInfo.class,
			 * inspector); System.out.println("Implementation name: " +
			 * si.getImplementationName()); String[] serviceNames =
			 * si.getSupportedServiceNames(); for(String nm : serviceNames)
			 * System.out.println("Service name: " + nm);
			 */
			XIntrospection intro = createInstanceMCF(XIntrospection.class, "com.sun.star.beans.Introspection");
			XIntrospectionAccess introAcc = intro.inspect(inspector);
			XIdlMethod method = introAcc.getMethod("inspect", -1); // get ref to XInspector.inspect()
			/*
			 * // alternative, low-level way of getting the method Object coreReflect =
			 * mcFactory.createInstanceWithContext(
			 * "com.sun.star.reflection.CoreReflection", xcc); XIdlReflection idlReflect =
			 * Lo.qi(XIdlReflection.class, coreReflect); XIdlClass idlClass =
			 * idlReflect.forName("org.openoffice.XInstanceInspector"); XIdlMethod[] methods
			 * = idlClass.getMethods(); System.out.println("No of methods: " +
			 * methods.length); for(XIdlMethod m : methods) System.out.println("  " +
			 * m.getName());
			 * 
			 * XIdlMethod method = idlClass.getMethod("inspect");
			 */
			System.out.println("inspect() method was found: " + (method != null));

			Object[][] params = new Object[][] { new Object[] { obj, title } };
			method.invoke(inspector, params);
		} catch (Exception e) {
			System.out.println("Could not access Inspector: " + e);
		}
	} // end of accessInspector()

	public static void mriInspect(Object obj)
	/*
	 * call MRI's inspect() Available from
	 * http://extensions.libreoffice.org/extension-center/mri-uno-object-inspection-
	 * tool or http://extensions.services.openoffice.org/en/project/MRI Docs:
	 * https://github.com/hanya/MRI/wiki Forum tutorial:
	 * https://forum.openoffice.org/en/forum/viewtopic.php?f=74&t=49294
	 */
	{
		XIntrospection xi = createInstanceMCF(XIntrospection.class, "mytools.Mri");
		if (xi == null) {
			System.out.println("MRI Inspector Service could not be instantiated");
			return;
		}

		System.out.println("MRI Inspector Service instantiated");
		xi.inspect(obj);
	} // end of mriInspect()

	// ------------------ color methods ---------------------

	public static int getColorInt(java.awt.Color color)
	// return the color as an integer, ignoring the alpha channel
	{
		if (color == null) {
			System.out.println("No color supplied");
			return 0;
		} else
			return (color.getRGB() & 0xffffff);
	} // end of getColorInt()

	public static int hexString2ColorInt(String colStr)
	// e.g. "#FF0000", "0xFF0000"
	{
		java.awt.Color color = java.awt.Color.decode(colStr);
		return getColorInt(color);
	}

	public static String getColorHexString(java.awt.Color color) {
		if (color == null) {
			System.out.println("No color supplied");
			return "#000000";
		} else
			return int2HexString(color.getRGB() & 0xffffff);
	} // end of getColorHexString()

	public static String int2HexString(int val) {
		String hex = Integer.toHexString(val);
		if (hex.length() < 6)
			hex = "000000".substring(0, 6 - hex.length()) + hex;
		return "#" + hex;
	} // end of int2HexString()

	// ================== other utils =============================

	public static void wait(int ms) // I can never remember the name :)
	{
		delay(ms);
	}

	public static void delay(int ms) {
		try {
			Thread.sleep(ms);
		} catch (InterruptedException e) {
		}
	}

	public static boolean isNullOrEmpty(String s) {
		return ((s == null) || (s.length() == 0));
	}

	public static void waitEnter() {
		System.out.println("Press Enter to continue...");
		try {
			System.in.read();
		} catch (java.io.IOException e) {
		}
	} // end of waitEnter()

	public static String getTimeStamp() {
		java.text.SimpleDateFormat sdf = new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		return sdf.format(new java.util.Date());
	}

	public static void printNames(String[] names) {
		printNames(names, 4);
	}

	public static void printNames(String[] names, int numPerLine)
	// print a large array with <numPerLine> strings/line, indented by 2 spaces
	{
		if (names == null)
			System.out.println("  No names found");
		else {
			Arrays.sort(names, String.CASE_INSENSITIVE_ORDER);
			int nlCounter = 0;
			System.out.println("No. of names: " + names.length);
			for (String name : names) {
				System.out.print("  \"" + name + "\"");
				nlCounter++;
				if (nlCounter % numPerLine == 0) {
					System.out.println();
					nlCounter = 0;
				}
			}
			System.out.println("\n\n");
		}
	} // end of printNames()

	public static void printTable(String name, Object[][] table) {
		System.out.println("-- " + name + " ----------------");
		for (int i = 0; i < table.length; i++) {
			for (int j = 0; j < table[i].length; j++)
				System.out.print("  " + table[i][j]);
			System.out.println();
		}
		System.out.println("-----------------------------\n");
	} // end of printTable()

	public static String capitalize(String s) {
		if ((s == null) || (s.length() == 0))
			return null;
		else if (s.length() == 1)
			return s.toUpperCase();
		else
			return Character.toUpperCase(s.charAt(0)) + s.substring(1);
	} // end of capitalize()

	public static int parseInt(String s) {
		if (s == null)
			return 0;
		try {
			return Integer.parseInt(s);
		} catch (NumberFormatException ex) {
			System.out.println(s + " could not be parsed as an int; using 0");
			return 0;
		}
	} // end of parseInt()

	public static void addJar(String jarPath)
	// load this JAR into the classloader at run time
	// from
	// http://stackoverflow.com/questions/60764/how-should-i-load-jars-dynamically-at-runtime
	{
		try {
			URLClassLoader classLoader = (URLClassLoader) ClassLoader.getSystemClassLoader();
			java.lang.reflect.Method m = URLClassLoader.class.getDeclaredMethod("addURL", java.net.URL.class);
			m.setAccessible(true);
			m.invoke(classLoader, new java.net.URL(jarPath));
		} catch (java.lang.Exception e) {
			System.out.println(e);
		}
	} // end of addJar()

	// ------------------- container manipulation --------------------

	public static String[] getContainerNames(XIndexAccess con)
	// extract the names of the elements in the indexed container
	{
		if (con == null) {
			System.out.println("Container is null");
			return null;
		}

		int numElems = con.getCount();
		if (numElems == 0) {
			System.out.println("No elements in the container");
			return null;
		}

		ArrayList<String> namesList = new ArrayList<String>();
		for (int i = 0; i < numElems; i++) {
			try {
				XNamed named = Lo.qi(XNamed.class, con.getByIndex(i));
				namesList.add(named.getName());
			} catch (Exception e) {
				System.out.println("Could not access name of element " + i);
			}
		}

		int sz = namesList.size();
		if (sz == 0) {
			System.out.println("No element names found in the container");
			return null;
		}

		String[] names = new String[sz];
		for (int i = 0; i < sz; i++)
			names[i] = namesList.get(i);

		return names;
	} // end of getContainerNames()

	public static XPropertySet findContainerProps(XIndexAccess con, String nm) {
		if (con == null) {
			System.out.println("Container is null");
			return null;
		}

		for (int i = 0; i < con.getCount(); i++) {
			try {
				Object oElem = con.getByIndex(i);
				XNamed named = Lo.qi(XNamed.class, oElem);
				if (named.getName().equals(nm)) {
					return (XPropertySet) Lo.qi(XPropertySet.class, oElem);
				}
			} catch (Exception e) {
				System.out.println("Could not access element " + i);
			}
		}

		System.out.println("Could not find a \"" + nm + "\" property set in the container");
		return null;
	} // end of findContainerProps()

	// ========================================================================
	// = Logger ===============================================================
	// ========================================================================

	private static final Logger LOG
		= Logger.getLogger( Lo.class.getName() );
}
