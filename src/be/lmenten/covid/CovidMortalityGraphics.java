package be.lmenten.covid;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.Rectangle;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.Serializable;
import java.net.URL;
import java.nio.channels.Channels;
import java.nio.channels.ReadableByteChannel;
import java.security.DigestInputStream;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Map;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.imageio.ImageIO;
import javax.swing.JFrame;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartMouseEvent;
import org.jfree.chart.ChartMouseListener;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.DateAxis;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.plot.Crosshair;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.xy.XYDifferenceRenderer;
import org.jfree.chart.renderer.xy.XYLineAndShapeRenderer;
import org.jfree.chart.title.TextTitle;
import org.jfree.chart.ui.RectangleEdge;
import org.jfree.chart.ui.RectangleInsets;
import org.jfree.data.general.DatasetUtils;
import org.jfree.data.time.Day;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;
//import org.jfree.svg.SVGGraphics2D;
//import org.jfree.svg.SVGUtils;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.sheet.XCellRangesQuery;
import com.sun.star.sheet.XSheetCellRanges;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCellRange;

import be.lmenten.lo.utils.LoSheetHelper;
import be.lmenten.utils.logging.AnsiLogFormatter;
import lo.utils.Calc;
import lo.utils.Lo;

public class CovidMortalityGraphics
	extends JFrame
	implements ChartMouseListener
{
	private static final long serialVersionUID = 1L;

	private static final boolean forceRetrieveFiles = false;

	// ========================================================================
	// =
	// ========================================================================

	static class StatData
		implements Serializable
	{
		private static final long serialVersionUID = 1L;

		Integer min;
		Integer max;
		Double avg;
		Double std;
	}

	// ------------------------------------------------------------------------

	private static Map<String,Integer[]> mortality_all = new TreeMap<>();

	private static Map<String,Integer[]> mortality_unkown = new TreeMap<>();
	private static Map<String,Integer[]> mortality_0_24 = new TreeMap<>();
	private static Map<String,Integer[]> mortality_25_44 = new TreeMap<>();
	private static Map<String,Integer[]> mortality_45_64 = new TreeMap<>();
	private static Map<String,Integer[]> mortality_65_74 = new TreeMap<>();
	private static Map<String,Integer[]> mortality_75_84 = new TreeMap<>();
	private static Map<String,Integer[]> mortality_85p = new TreeMap<>();

	private static Map<String,StatData> mortality_stat_all = new TreeMap<>();

	private static Map<String,StatData> mortality_stat_unkown = new TreeMap<>();
	private static Map<String,StatData> mortality_stat_0_24 = new TreeMap<>();
	private static Map<String,StatData> mortality_stat_25_44 = new TreeMap<>();
	private static Map<String,StatData> mortality_stat_45_64 = new TreeMap<>();
	private static Map<String,StatData> mortality_stat_65_74 = new TreeMap<>();
	private static Map<String,StatData> mortality_stat_75_84 = new TreeMap<>();
	private static Map<String,StatData> mortality_stat_85p = new TreeMap<>();

	private static LocalDate maxMortalityDate = LocalDate.of( 1900, 1, 1 );

	// ------------------------------------------------------------------------

	private static Map<String,Integer[]> covid_all = new TreeMap<>();

	private static Map<String,Integer[]> covid_unkown = new TreeMap<>();
	private static Map<String,Integer[]> covid_0_24 = new TreeMap<>();
	private static Map<String,Integer[]> covid_25_44 = new TreeMap<>();
	private static Map<String,Integer[]> covid_45_64 = new TreeMap<>();
	private static Map<String,Integer[]> covid_65_74 = new TreeMap<>();
	private static Map<String,Integer[]> covid_75_84 = new TreeMap<>();
	private static Map<String,Integer[]> covid_85p = new TreeMap<>();

	private static LocalDate maxCovidDate = LocalDate.of( 1900, 1, 1 );

	// ------------------------------------------------------------------------

	private static XComponentLoader loader = null;

    private ChartPanel panel;
    
    private Crosshair xCrosshair;

    private Crosshair[] yCrosshairs;

	// ========================================================================
	// =
	// ========================================================================

	@SuppressWarnings("unchecked")
	public static void main( final String[] args )
	{
		AnsiLogFormatter.install();

		// --------------------------------------------------------------------

		File dataFile = new File( "data.bin" );
		if( ! dataFile.exists() || (dataFile.exists() && forceRetrieveFiles) )
		{
			loader = Lo.loadOffice();

			processMorbidity();
			processCovid();

			Lo.closeOffice();

			try( FileOutputStream fos = new FileOutputStream( dataFile ) )
			{
				try( ObjectOutputStream oos = new ObjectOutputStream( fos ) )
				{
					oos.writeObject( mortality_all );
					oos.writeObject( mortality_unkown );
					oos.writeObject( mortality_0_24 );
					oos.writeObject( mortality_25_44 );
					oos.writeObject( mortality_45_64 );
					oos.writeObject( mortality_65_74 );
					oos.writeObject( mortality_75_84 );
					oos.writeObject( mortality_85p );

					oos.writeObject( mortality_stat_all );
					oos.writeObject( mortality_stat_unkown );
					oos.writeObject( mortality_stat_0_24 );
					oos.writeObject( mortality_stat_25_44 );
					oos.writeObject( mortality_stat_45_64 );
					oos.writeObject( mortality_stat_65_74 );
					oos.writeObject( mortality_stat_75_84 );
					oos.writeObject( mortality_stat_85p );

					oos.writeObject( maxMortalityDate );

					oos.writeObject( covid_all );
					oos.writeObject( covid_unkown );
					oos.writeObject( covid_0_24 );
					oos.writeObject( covid_25_44 );
					oos.writeObject( covid_45_64 );
					oos.writeObject( covid_65_74 );
					oos.writeObject( covid_75_84 );
					oos.writeObject( covid_85p );

					oos.writeObject( maxCovidDate );
				}
			}
			catch( FileNotFoundException ex )
			{
				ex.printStackTrace();
			}
			catch( IOException ex )
			{
				ex.printStackTrace();
			}
		}
		else
		{
			try( FileInputStream fis = new FileInputStream( dataFile ) )
			{
				try( ObjectInputStream ois = new ObjectInputStream( fis ) )
				{
					mortality_all = (Map<String,Integer[]>) ois.readObject();
					mortality_unkown = (Map<String,Integer[]>) ois.readObject();
					mortality_0_24 = (Map<String,Integer[]>) ois.readObject();
					mortality_25_44 = (Map<String,Integer[]>) ois.readObject();
					mortality_45_64 = (Map<String,Integer[]>) ois.readObject();
					mortality_65_74 = (Map<String,Integer[]>) ois.readObject();
					mortality_75_84 = (Map<String,Integer[]>) ois.readObject();
					mortality_85p = (Map<String,Integer[]>) ois.readObject();

					mortality_stat_all = (Map<String,StatData>) ois.readObject();
					mortality_stat_unkown = (Map<String,StatData>) ois.readObject();
					mortality_stat_0_24 = (Map<String,StatData>) ois.readObject();
					mortality_stat_25_44 = (Map<String,StatData>) ois.readObject();
					mortality_stat_45_64 = (Map<String,StatData>) ois.readObject();
					mortality_stat_65_74 = (Map<String,StatData>) ois.readObject();
					mortality_stat_75_84 = (Map<String,StatData>) ois.readObject();
					mortality_stat_85p = (Map<String,StatData>) ois.readObject();

					maxMortalityDate = (LocalDate) ois.readObject();

					covid_all = (Map<String,Integer[]>) ois.readObject();
					covid_unkown = (Map<String,Integer[]>) ois.readObject();
					covid_0_24 = (Map<String,Integer[]>) ois.readObject();
					covid_25_44 = (Map<String,Integer[]>) ois.readObject();
					covid_45_64 = (Map<String,Integer[]>) ois.readObject();
					covid_65_74 = (Map<String,Integer[]>) ois.readObject();
					covid_75_84 = (Map<String,Integer[]>) ois.readObject();
					covid_85p = (Map<String,Integer[]>) ois.readObject();

					maxCovidDate = (LocalDate) ois.readObject();
				}
			}
			catch( ClassNotFoundException ex )
			{
				ex.printStackTrace();
				System.exit( -1 );
			}
			catch( FileNotFoundException ex )
			{
				ex.printStackTrace();
				System.exit( -1 );
			}
			catch( IOException ex )
			{
				ex.printStackTrace();
				System.exit( -1 );
			}
		}

		// --------------------------------------------------------------------

		for( AgeGroup25 group : AgeGroup25.values() )
		{
			if( group == AgeGroup25.UNKNOWN )
			{
				group = null;
			}

			CovidMortalityGraphics app = new CovidMortalityGraphics( "Mortalité COVID", group );
			app.pack();
			app.setVisible( true );
		}
	}

	// ========================================================================
	// =
	// ========================================================================

	private int w = 2048;
	private int h = 720;

	public CovidMortalityGraphics( final String title, AgeGroup25 ageGroup )
	{
		setTitle(title);

		String imageFileName = "deaths-"
			+ ((ageGroup != null) ? ageGroup.name() : "all" )
			;

		JFreeChart chart = createChart( ageGroup );

		// --------------------------------------------------------------------

//		try
//		{
//			SVGGraphics2D g2 = new SVGGraphics2D( w, h );
//			chart.draw( g2, new Rectangle( 0, 0, w, h ) );
//			SVGUtils.writeToSVG( new File( imageFileName + ".svg" ), g2.getSVGElement() );
//		}
//		catch( IOException ex )
//		{
//			ex.printStackTrace();
//		}

		// --------------------------------------------------------------------

		try
		{
			BufferedImage image = new BufferedImage( w, h, BufferedImage.TYPE_INT_ARGB);
		    Graphics2D g2 = image.createGraphics();

		    g2.setRenderingHint( JFreeChart.KEY_SUPPRESS_SHADOW_GENERATION, true );
		    Rectangle r = new Rectangle(0, 0, w, h );
		    chart.draw( g2, r );

		    File f = new File( imageFileName + ".png" );
		    BufferedImage chartImage = chart.createBufferedImage( w, h, null); 
		    ImageIO.write( chartImage, "png", f );
	    }
		catch( IOException ex )
		{
			ex.printStackTrace();
		}

		// --------------------------------------------------------------------

		panel = new ChartPanel( chart );
		panel.setPreferredSize( new Dimension( 1024, 360 ) );

//        panel.addChartMouseListener( this );

//        CrosshairOverlay crosshairOverlay = new CrosshairOverlay();
//        xCrosshair = new Crosshair( Double.NaN, Color.GRAY, new BasicStroke( 0.5f ) );
//        xCrosshair.setLabelVisible( true );
//        crosshairOverlay.addDomainCrosshair( xCrosshair );
//
//        yCrosshairs = new Crosshair[ dataset.getSeriesCount() ];
//        for( int i = 0; i <  dataset.getSeriesCount() ; i++ )
//        {
//            yCrosshairs[i] = new Crosshair(Double.NaN, Color.GRAY, new BasicStroke( 0.5f ) );
//            yCrosshairs[i].setLabelVisible( true );
//            if( i % 2 != 0 )
//            {
//                this.yCrosshairs[i].setLabelAnchor( RectangleAnchor.TOP_RIGHT );
//            }
//            crosshairOverlay.addRangeCrosshair( yCrosshairs[i] );
//        }
//        panel.addOverlay(crosshairOverlay);
		
		setLayout( new BorderLayout() );
		getContentPane().add(panel, BorderLayout.CENTER );
		setDefaultCloseOperation( EXIT_ON_CLOSE );
	}

	private JFreeChart createChart( AgeGroup25 ageGroup )
	{
		String title = "Mortalité au cours de la pandemie Covid-19";
		String subtitle = (ageGroup != null) ? ageGroup.getTitle() : "Tous ages confondus";
		String mortalityDate = maxMortalityDate.toString();
		String covidDate = maxCovidDate.toString();
		String subtitle2 = String.format( "Dernières dates: mortalité: %s, covid: %s.", mortalityDate, covidDate );
		
		JFreeChart chart = ChartFactory.createTimeSeriesChart(
				title,						// title
				"Date",						// x-axis label
				"Décès",					// y-axis label
				null,						// dataset
				true,						// create legend?
				true,						// generate tooltips?
				false						// generate URLs?
				);

		chart.addSubtitle( new TextTitle( subtitle ) );
		chart.addSubtitle( new TextTitle( subtitle2 ) );
		

		chart.setBackgroundPaint( Color.white );

		// --------------------------------------------------------------------

		XYPlot plot = (XYPlot) chart.getPlot();
		plot.setBackgroundPaint( Color.lightGray );
		plot.setDomainGridlinePaint( Color.white );
		plot.setRangeGridlinePaint( Color.white );
		plot.setAxisOffset(new RectangleInsets( 5.0, 5.0, 5.0, 5.0 ) );
		plot.setDomainCrosshairVisible( true );
		plot.setRangeCrosshairVisible( true );

		// --------------------------------------------------------------------

		Color minMaxAreaColor = new Color( 127, 127, 127, 127 );
		Color minMaxColor = new Color( 64, 64, 64, 159 );

        XYDifferenceRenderer minMaxRenderer = new XYDifferenceRenderer( minMaxAreaColor, minMaxAreaColor, false );
        minMaxRenderer.setSeriesPaint( 0, minMaxColor );
        minMaxRenderer.setSeriesPaint( 1, minMaxColor );
        minMaxRenderer.setRoundXCoordinates(true);

        plot.setDataset( 2, createMinMaxDataset( ageGroup ) );
        plot.setRenderer( 2, minMaxRenderer );
		
		// --------------------------------------------------------------------

		Color avgColor = new Color( 32, 32, 32, 191 );

        XYLineAndShapeRenderer normalRendererAvg = new XYLineAndShapeRenderer();
        normalRendererAvg.setSeriesShapesVisible( 0, false );
        normalRendererAvg.setSeriesPaint( 0, avgColor );

		plot.setDataset( 1, createAvgDataset( ageGroup ) );
        plot.setRenderer( 1, normalRendererAvg );

		// --------------------------------------------------------------------

        XYLineAndShapeRenderer normalRenderer = new XYLineAndShapeRenderer();
        normalRenderer.setSeriesShapesVisible( 0, false );
        normalRenderer.setSeriesPaint( 0, Color.BLUE );
        normalRenderer.setSeriesShapesVisible( 1, false );
        normalRenderer.setSeriesPaint( 1, Color.RED );

		plot.setDataset( 0, createDataset( ageGroup ) );
        plot.setRenderer( 0, normalRenderer );

		// --------------------------------------------------------------------

		DateAxis axis = (DateAxis) plot.getDomainAxis();
		axis.setDateFormatOverride( new SimpleDateFormat( "MMM-yyyy" ) );

		return chart;
	}

	// ========================================================================
	// =
	// ========================================================================

	private TimeSeriesCollection createDataset( AgeGroup25 ageGroup )
	{
		Map<String,Integer[]> dataMortality = getMortalityData( ageGroup );
		Map<String,Integer[]> dataCovid = getCovidData( ageGroup );
		
		// --------------------------------------------------------------------

		TimeSeries covidDeaths = new TimeSeries( "Mortalité covid" );		
		TimeSeries generalDeaths = new TimeSeries( "Mortalité générale" );

		LocalDate date = LocalDate.of( 2020, 1, 1 );
		for( date = LocalDate.of( 2020, 1, 1 ) ; date.isBefore( LocalDate.now() ) ; date = date.plusDays( 1 ) )
		{
			String key = String.format( "%02d.%02d", date.getMonthValue(), date.getDayOfMonth() );

			int year = date.getYear();
			Day day = new Day( date.getDayOfMonth(), date.getMonthValue(), date.getYear() );

			Integer [] value = dataCovid.get( key );
			if( value != null && value[year - 2020] != null )
			{
				covidDeaths.add( day, value[year - 2020] );
			}
			else
			{
				covidDeaths.add( day, 0 );
			}

			Integer [] deaths = dataMortality.get( key );
			if( deaths != null && deaths[year - 2009] != null )
			{
				generalDeaths.add( day, deaths[year - 2009] );
			}
//			else
//			{
//				generalDeaths.add( day, 0 );
//			}
		}

		TimeSeriesCollection dataset = new TimeSeriesCollection();
		dataset.addSeries( generalDeaths );
		dataset.addSeries( covidDeaths );

		return dataset;
	}

	// ========================================================================
	// =
	// ========================================================================

	private TimeSeriesCollection createMinMaxDataset( AgeGroup25 ageGroup )
	{
		Map<String,StatData> dataMortalityStat = getMortalityStatData( ageGroup );

		TimeSeries min = new TimeSeries( "Min" );
		TimeSeries max = new TimeSeries( "Max" );

		LocalDate date = LocalDate.of( 2020, 1, 1 );
		for( date = LocalDate.of( 2020, 1, 1 ) ; date.isBefore( LocalDate.now() ) ; date = date.plusDays( 1 ) )
		{
			String key = String.format( "%02d.%02d", date.getMonthValue(), date.getDayOfMonth() );
			StatData stat = dataMortalityStat.get( key );

			Day day = new Day( date.getDayOfMonth(), date.getMonthValue(), date.getYear() );
			min.add( day, stat.min );
			max.add( day, stat.max );
		}

		TimeSeriesCollection dataset = new TimeSeriesCollection();
		dataset.addSeries( min );
		dataset.addSeries( max );

		return dataset;
	}

	private TimeSeriesCollection createAvgDataset( AgeGroup25 ageGroup )
	{
		Map<String,StatData> dataMortalityStat = getMortalityStatData( ageGroup );

		TimeSeries avg = new TimeSeries( "Avg" );

		LocalDate date = LocalDate.of( 2020, 1, 1 );
		for( date = LocalDate.of( 2020, 1, 1 ) ; date.isBefore( LocalDate.now() ) ; date = date.plusDays( 1 ) )
		{
			String key = String.format( "%02d.%02d", date.getMonthValue(), date.getDayOfMonth() );
			StatData stat = dataMortalityStat.get( key );

			Day day = new Day( date.getDayOfMonth(), date.getMonthValue(), date.getYear() );
			avg.add( day, stat.avg );
		}

		TimeSeriesCollection dataset = new TimeSeriesCollection();
		dataset.addSeries( avg );

		return dataset;
	}

	// ========================================================================
	// =
	// ========================================================================

	private Map<String,Integer[]> getMortalityData( AgeGroup25 ageGroup )
	{
		if( ageGroup == null )
		{
			return mortality_all;
		}
		else switch( ageGroup )
		{
			case AGE_0_24:		return mortality_0_24;
			case AGE_25_44:		return mortality_25_44;
			case AGE_45_64:		return mortality_45_64;
			case AGE_65_74:		return mortality_65_74;
			case AGE_75_84:		return mortality_75_84;
			case AGE_85_PLUS:	return mortality_85p;
			case UNKNOWN:		return mortality_unkown;
		}

		throw new IllegalArgumentException();
	}

	private Map<String,StatData> getMortalityStatData( AgeGroup25 ageGroup )
	{
		if( ageGroup == null )
		{
			return mortality_stat_all;
		}
		else switch( ageGroup )
		{
			case AGE_0_24:		return mortality_stat_0_24;
			case AGE_25_44:		return mortality_stat_25_44;
			case AGE_45_64:		return mortality_stat_45_64;
			case AGE_65_74:		return mortality_stat_65_74;
			case AGE_75_84:		return mortality_stat_75_84;
			case AGE_85_PLUS:	return mortality_stat_85p;
			case UNKNOWN:		return mortality_stat_unkown;
		}

		throw new IllegalArgumentException();
	}

	private Map<String,Integer[]> getCovidData( AgeGroup25 ageGroup )
	{
		if( ageGroup == null )
		{
			return covid_all;
		}
		else switch( ageGroup )
		{
			case AGE_0_24:		return covid_0_24;
			case AGE_25_44:		return covid_25_44;
			case AGE_45_64:		return covid_45_64;
			case AGE_65_74:		return covid_65_74;
			case AGE_75_84:		return covid_75_84;
			case AGE_85_PLUS:	return covid_85p;
			case UNKNOWN:		return covid_unkown;
		}

		throw new IllegalArgumentException();
	}

	// ========================================================================
	// =
	// ========================================================================

    @Override
    public void chartMouseClicked( ChartMouseEvent event )
    {
    }

    @Override
    public void chartMouseMoved( ChartMouseEvent event )
    {
        Rectangle2D dataArea = panel.getScreenDataArea();
        JFreeChart chart = event.getChart();

        XYPlot plot = (XYPlot) chart.getPlot();
        ValueAxis xAxis = plot.getDomainAxis();
        double x = xAxis.java2DToValue( event.getTrigger().getX(), dataArea, RectangleEdge.BOTTOM );
        xCrosshair.setValue(x);

        for( int i = 0; i < yCrosshairs.length ; i++ )
        {
            double y = DatasetUtils.findYValue( plot.getDataset(), i, x );
            yCrosshairs[i].setValue( y );
        }
    }

	// ========================================================================
	// =
	// ========================================================================

	private static final String DEATH_REMOTE_URL = "https://statbel.fgov.be/sites/default/files/files/opendata/deathday/DEMO_DEATH_OPEN.xlsx";
	private static final String DEATH_LOCAL_URL = "DEMO_DEATH_OPEN.xlsx";
	private static final String DEATH_SHEET = "DEMO_DEATH_OPEN";

	private static void processMorbidity()
	{
		String logString;

		// --------------------------------------------------------------------

		if( CovidMortalityGraphics.forceRetrieveFiles )
		{
			retrieveFile( DEATH_LOCAL_URL, DEATH_REMOTE_URL );
		}

		// --------------------------------------------------------------------

		XSpreadsheetDocument doc = Calc.openDoc( DEATH_LOCAL_URL, loader );
	    XSpreadsheet sheet = Calc.getSheet( doc, DEATH_SHEET );

		XCellRange cellRange = Calc.findUsedRange( sheet );
		XCellRangesQuery query = Lo.qi( XCellRangesQuery.class, cellRange );
		XSheetCellRanges ranges = query.queryVisibleCells();
		CellRangeAddress [] addresses = ranges.getRangeAddresses();

		final int rowCount = (addresses[0].EndRow - addresses[0].StartRow);
		final int colCount = addresses[0].EndColumn - addresses[0].StartColumn + 1;
		final String colNames = LoSheetHelper.getColNamesList( sheet, addresses[0].StartRow, colCount );

		logString = String.format( "%s: Rows=%d, Columns=%d (%S)", DEATH_LOCAL_URL, rowCount, colCount, colNames );
		LOG.info( logString );

		long start = System.currentTimeMillis();
		for( int i = addresses[0].StartRow + 1 ; i <= addresses[0].EndRow ; i++ )
		{
			if( (i % (rowCount/200)) == 0 )
			{
				long duration = System.currentTimeMillis() - start;
				
				double p = ((i *100.d) / rowCount);
				long eta = (long) ((((double)duration / p) * (100.d-p)) / 1000.d);
				long m = eta / 60l;
				long s = eta % 60l;

				logString = String.format( "%2.2f %% (e.t.a. %d'%02d) ...", p, m, s );
				LOG.info( logString );
			}

			// ----------------------------------------------------------------

			// 0      1       2        3      4           5       6       7       8
			// CD_ARR,CD_PROV,CD_REGIO,CD_SEX,CD_AGEGROUP,DT_DATE,NR_YEAR,NR_WEEK,MS_NUM_DEATH

			String ageGroupRaw = (String) Calc.getString( sheet, 4, i );
			AgeGroup25 ageGroup = AgeGroup25.lookup( ageGroupRaw );

			Map<String,Integer[]> currentMap;
			switch( ageGroup )
			{
				case AGE_0_24:		currentMap = mortality_0_24;	break;
				case AGE_25_44:		currentMap = mortality_25_44;	break;
				case AGE_45_64:		currentMap = mortality_45_64;	break;
				case AGE_65_74:		currentMap = mortality_65_74;	break;
				case AGE_75_84:		currentMap = mortality_75_84;	break;
				case AGE_85_PLUS:	currentMap = mortality_85p;		break;
				default:			currentMap = mortality_unkown;	break;
			}

			// ----------------------------------------------------------------

			LocalDate date = LoSheetHelper.getDate( sheet, 5, i );
			if( date.isAfter( maxMortalityDate ) )
			{
				maxMortalityDate = date;
			}

			String day = String.format( "%02d.%02d", date.getMonthValue(), date.getDayOfMonth() );
	
			Integer [] currentDay = currentMap.get( day );
			if( currentDay == null )
			{
				currentDay = new Integer [ (2021 - 2009) + 1 ]; 
				currentMap.put( day, currentDay );
			}
	
			int yearIndex = date.getYear() - 2009;
			int deaths = LoSheetHelper.getInteger( sheet, 8, i );

			if( currentDay[yearIndex] == null )
			{
				currentDay[yearIndex] = deaths;
			}
			else
			{
				currentDay[yearIndex] += deaths;
			}

			// ----------------------------------------------------------------

			currentDay = mortality_all.get( day );
			if( currentDay == null )
			{
				currentDay = new Integer [ (2021 - 2009) + 1 ]; 
				mortality_all.put( day, currentDay );
			}

			if( currentDay[yearIndex] == null )
			{
				currentDay[yearIndex] = deaths;
			}
			else
			{
				currentDay[yearIndex] += deaths;
			}
		}

		// --------------------------------------------------------------------

		processMorbidity( mortality_all, mortality_stat_all ); 

		processMorbidity( mortality_0_24, mortality_stat_0_24 ); 
		processMorbidity( mortality_25_44, mortality_stat_25_44 ); 
		processMorbidity( mortality_45_64, mortality_stat_45_64 ); 
		processMorbidity( mortality_65_74, mortality_stat_65_74 ); 
		processMorbidity( mortality_75_84, mortality_stat_75_84 ); 
		processMorbidity( mortality_85p, mortality_stat_85p ); 
		processMorbidity( mortality_unkown, mortality_stat_unkown ); 
		
	}

	private static void processMorbidity( Map<String,Integer[]> in, Map<String,StatData> out )
	{
		in.forEach( (key,value) ->
		{
			int count = 0;

			// ----------------------------------------------------------------

			long sum = 0l;

			Integer min = null;
			Integer max = null;

			for( int i = 0 ; i < (2020-2009) ; i++ )
			{
				if( value[i] != null )
				{
					count++;

					sum += value[i];

					if( min == null )
					{
						min = Integer.valueOf( value[i] );
					}
					else if( value[i] < min )
					{
						min = value[i];
					}

					if( max == null )
					{
						max = Integer.valueOf( value[i] );
					}
					else if( value[i] > max )
					{
						max  = value[i];
					}
				}
				else if( ! key.equals( "02.29" ) )
				{
					count++;
				}
			}

			double avg = (double) sum / (double) count;

			// ----------------------------------------------------------------

			double sum2 = 0.d;

			for( int i = 0 ; i < value.length ; i++ )
			{
				if( value[i] != null )
				{
					sum2 += Math.pow( (double) value[i] - avg, 2 );
				}
				else if( ! key.equals( "02.29" ) )
				{
					sum2 += Math.pow( 0.d - avg, 2 );
				}
			}

			double std = Math.sqrt( sum2 / (count - 1) );

			// ----------------------------------------------------------------

			StatData data = out.get( key );
			if( data == null )
			{
				data = new StatData();
				out.put( key, data );
			}

			data.min = min;
			data.max = max;
			data.avg = avg;
			data.std = std;
		} );		
	}

	// ========================================================================
	// =
	// ========================================================================

	private static final String COVID_REMOTE_URL = "https://epistat.sciensano.be/Data/COVID19BE.xlsx";
	private static final String COVID_LOCAL_URL = "COVID19BE.xlsx";
	private static final String COVID_DEATH_SHEET = "MORT";

	private static void processCovid()
	{
		String logString;

		// --------------------------------------------------------------------

		if( CovidMortalityGraphics.forceRetrieveFiles )
		{
			retrieveFile( COVID_LOCAL_URL, COVID_REMOTE_URL );
		}

		// --------------------------------------------------------------------

		XSpreadsheetDocument doc = Calc.openDoc( COVID_LOCAL_URL, loader );
	    XSpreadsheet sheet = Calc.getSheet( doc, COVID_DEATH_SHEET );

		XCellRange cellRange = Calc.findUsedRange( sheet );
		XCellRangesQuery query = Lo.qi( XCellRangesQuery.class, cellRange );
		XSheetCellRanges ranges = query.queryVisibleCells();
		CellRangeAddress [] addresses = ranges.getRangeAddresses();

		final int rowCount = (addresses[0].EndRow - addresses[0].StartRow);
		final int colCount = addresses[0].EndColumn - addresses[0].StartColumn + 1;
		final String colNames = LoSheetHelper.getColNamesList( sheet, addresses[0].StartRow, colCount );

		logString = String.format( "%s: Rows=%d, Columns=%d (%S)", COVID_LOCAL_URL, rowCount, colCount, colNames );
		LOG.info( logString );

		long start = System.currentTimeMillis();
		for( int i = addresses[0].StartRow + 1 ; i <= addresses[0].EndRow ; i++ )
		{
			if( (i % (rowCount/200)) == 0 )
			{
				long duration = System.currentTimeMillis() - start;
				
				double p = ((i *100.d) / rowCount);
				long eta = (long) ((((double)duration / p) * (100.d-p)) / 1000.d);
				long m = eta / 60l;
				long s = eta % 60l;

				logString = String.format( "%2.2f %% (e.t.a. %d'%02d) ...", p, m, s );
				LOG.info( logString );
			}

			// ----------------------------------------------------------------

			// 0    1      2        3   4
		    // DATE REGION AGEGROUP SEX DEATHS

			String ageGroupRaw = (String) Calc.getString( sheet, 2, i );
			AgeGroup25 ageGroup = AgeGroup25.lookup( ageGroupRaw );

			Map<String,Integer[]> currentMap;
			switch( ageGroup )
			{
				case AGE_0_24:		currentMap = covid_0_24;	break;
				case AGE_25_44:		currentMap = covid_25_44;	break;
				case AGE_45_64:		currentMap = covid_45_64;	break;
				case AGE_65_74:		currentMap = covid_65_74;	break;
				case AGE_75_84:		currentMap = covid_75_84;	break;
				case AGE_85_PLUS:	currentMap = covid_85p;		break;
				default:			currentMap = covid_unkown;	break;
			}
			
			// ----------------------------------------------------------------

			LocalDate date = LoSheetHelper.getDate( sheet, 0, i );
			if( date.isAfter( maxCovidDate ) )
			{
				maxCovidDate = date;
			}

			String day = String.format( "%02d.%02d", date.getMonthValue(), date.getDayOfMonth() );

			Integer [] currentDay = currentMap.get( day );
			if( currentDay == null )
			{
				currentDay = new Integer [ (2021 - 2020) + 1 ]; 
				currentMap.put( day, currentDay );
			}
	
			int yearIndex = date.getYear() - 2020;
			int deaths = LoSheetHelper.getInteger( sheet, 4, i );

			if( currentDay[yearIndex] == null )
			{
				currentDay[yearIndex] = deaths;
			}
			else
			{
				currentDay[yearIndex] += deaths;
			}

			// ----------------------------------------------------------------

			currentDay = covid_all.get( day );
			if( currentDay == null )
			{
				currentDay = new Integer [ (2021 - 2020) + 1 ]; 
				covid_all.put( day, currentDay );
			}

			if( currentDay[yearIndex] == null )
			{
				currentDay[yearIndex] = deaths;
			}
			else
			{
				currentDay[yearIndex] += deaths;
			}			
		}
	}

	// ========================================================================
	// =
	// ========================================================================

	private static byte [] retrieveFile( String localUrl, String remoteUrl )
	{
		File localFile = new File( localUrl );

		byte [] md5 = null;

		try
		{
			LOG.info( "Downloading file '" + remoteUrl + "'." );
		
			MessageDigest md = MessageDigest.getInstance( "MD5" );
		
			URL url = new URL( remoteUrl );
			InputStream is = url.openStream();
			DigestInputStream dis = new DigestInputStream( is, md );
			ReadableByteChannel readableByteChannel = Channels.newChannel( dis );
		
			FileOutputStream fileOutputStream = new FileOutputStream( localFile );
		
			fileOutputStream.getChannel().transferFrom( readableByteChannel, 0, Long.MAX_VALUE );

			fileOutputStream.close();
			dis.close();

			md5 = md.digest();
		}
		catch( NoSuchAlgorithmException ex )
		{
			LOG.log( Level.SEVERE, "MD5 algorithm", ex );

			return null;
		}
		catch( IOException ex )
		{
			LOG.log( Level.WARNING, "Download error", ex );

			return null;
		}

		return md5;
	}
		
	private static final Logger LOG = Logger.getLogger( CovidMortalityGraphics.class.getName() );
}

