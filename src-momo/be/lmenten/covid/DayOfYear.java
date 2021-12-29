package be.lmenten.covid;

import java.time.DateTimeException;

public class DayOfYear<T extends Number & Comparable<T>>
	implements Comparable<DayOfYear<T>>
{
	private final int day;
	private final int month;
	private final T [] data;

	private int count;

	private T min;
	private T max;
	private double avg;
	private double std;

	// ========================================================================
	// =
	// ========================================================================

	public DayOfYear( int day, int month, T [] data )
		throws DateTimeException
	{
		this.day = day;
		this.month = month;

		if( ! isValid() )
		{
			throw new DateTimeException( "Invalid parameters for DayOfYear" );
		}

		this.data = data;
	}

	// ========================================================================
	// =
	// ========================================================================

	/**
	 * 
	 * @return
	 */
	public int getDay()
	{
		return day;
	}

	/**
	 * 
	 * @return
	 */
	public int getMonth()
	{
		return month;
	}

	// ------------------------------------------------------------------------

	/**
	 * 
	 * @return
	 */
	public T getMin()
	{
		return min;
	}

	/**
	 * 
	 * @return
	 */
	public double getAverage()
	{
		return avg;
	}

	/**
	 * 
	 * @return
	 */
	public T getMax()
	{
		return max;
	}

	/**
	 * 
	 * @return
	 */
	public double getStandartDeviation()
	{
		return std;
	}

	// ------------------------------------------------------------------------

	/**
	 * 
	 * @param index
	 * @param dataItem
	 */
	public void setData( int index, T dataItem )
	{
		data[index] = dataItem;
	}

	/**
	 * 
	 * @param index
	 * @param dataItem
	 * @param op
	 */
	public void updateData( int index, T dataItem, Operation<T> op )
	{
		data[index] = op.operation( data[index], dataItem ); 
	}

	/**
	 * 
	 * @param index
	 * @return
	 */
	public T getData( int index )
	{
		return data[index];
	}

	// ------------------------------------------------------------------------

	/**
	 * 
	 */
	public void compute()
	{
		T sum = null;
		for( int index = 0 ; index < data.length ; index++ )
		{
			if( data[index] != null )
			{
				if( (min == null) || (data[index].compareTo(min) < 0) )
				{
					min = data[index];
				}

				if( (max == null) || (data[index].compareTo(max) > 0) )
				{
					max = data[index];
				}
				
				sum = (sum == null) ? data[index] : add( sum, data[index] );
				count++;
			}
		}

		avg = Double.valueOf( sum.doubleValue() / (double)count );

		double var = 0.d;
		for( int index = 0 ; index < data.length ; index++ )
		{
			if( data[index] != null )
			{
				var += Math.pow( data[index].doubleValue() - avg, 2 );
			}
		}

		std = Math.sqrt( var / (count - 1) );
	}

	@SuppressWarnings("unchecked")
	private final T add( T t1, T t2 )
	{
		if( t1 instanceof Double )
		{
			return (T) Double.valueOf((t1.doubleValue() + t2.doubleValue()));
		}
		else if( t1 instanceof Float )
		{
			return (T) Float.valueOf(((t1.floatValue() + t2.floatValue())));
		}

		else if (t1 instanceof Long)
		{
			return (T) Long.valueOf(((t1.longValue() + t2.longValue())));
		}
		else if (t1 instanceof Integer)
		{
			return (T) Integer.valueOf(((t1.intValue() + t2.intValue())));
		}

		throw new IllegalArgumentException();
    }
		
	// ========================================================================
	// =
	// ========================================================================

	public boolean isValid()
	{
		if( day < 1  )
		{
			return false;
		}

		switch( month )
		{
			case 1:		// January
			case 3:		// March
			case 5:		// May
			case 7:		// July
			case 8:		// Augusts
			case 10:	// October
			case 12:	// December
			{
				if( day > 31 )
				{
					return false;
				}

				break;
			}

			case 4:		// April
			case 6:		// June
			case 9:		// September
			case 11:	// November
			{
				if( day > 30 )
				{
					return false;
				}

				break;
			}

			case 2:		// February
			{
				if( day > 29 )
				{
					return false;
				}
				break;
			}
		}
		return true;
	}

	public boolean isValid( int year )
	{
		if( (month == 2) && (day == 29) )
		{
			return ((year % 4 == 0) && (year % 400 == 0));
		}

		return isValid();
	}

	// ========================================================================
	// =
	// ========================================================================

	@Override
	public int compareTo( DayOfYear<T> o )
	{
		int m = o.month - this.month;
		if( m != 0 )
		{
			return m;
		}

		return o.day - this.day;
	}
}
