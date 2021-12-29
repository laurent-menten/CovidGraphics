package be.lmenten.covid;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.util.Map;
import java.util.TreeMap;

public class Analyse1b
{
	@SuppressWarnings("unchecked")
	public static void main( String[] args )
	{
		// --------------------------------------------------------------------
		// - Load data --------------------------------------------------------
		// --------------------------------------------------------------------

		Map<AgeGroup25,Map<String,Integer[]>> data = new TreeMap<>();

		try( FileInputStream fis = new FileInputStream( "data-mortalité.bin" ) )
		{
			try( ObjectInputStream ois = new ObjectInputStream( fis ) )
			{
				data = (Map<AgeGroup25,Map<String,Integer[]>>) ois.readObject();
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

		// --------------------------------------------------------------------
		// - Display data -----------------------------------------------------
		// --------------------------------------------------------------------

		data.forEach(
			(key, value) -> 
			{
				System.out.printf( "%s\n", key );
				System.out.printf( "Day, min, avg, max, std \n" );

				value.forEach(
					(key2, value2) ->
					{
						System.out.printf( "%s, ", key2 );

						int count = 0;
						long sum = 0l;
						Integer min = null;
						Integer max = null;

						for( int i = 0 ; i < value2.length ; i++ )
						{
							if( value2[i] != null )
							{
								count++;

								sum += value2[i];

								if( min == null )
								{
									min = Integer.valueOf( value2[i] );
								}
								else if( value2[i] < min )
								{
									min = value2[i];
								}

								if( max == null )
								{
									max = Integer.valueOf( value2[i] );
								}
								else if( value2[i] > max )
								{
									max  =value2[i];
								}
							}
							else if( ! key2.equals( "02.29" ) )
							{
								count++;
							}
						}

						double avg = (double)sum / (double)count;

						double sum2 = 0.d;

						for( int i = 0 ; i < value2.length ; i++ )
						{
							if( value2[i] != null )
							{
								sum2 += Math.pow( (double)value2[i] - avg, 2 );
							}
							else if( ! key2.equals( "02.29" ) )
							{
								sum2 += Math.pow( 0.d - avg, 2 );
							}
						}

						double std = Math.sqrt( sum2 / (count - 1) );
						
						System.out.printf( "%d, %f, %d, %f\n", min, avg, max, std );
					}
				);
			}
		);
	}
}
