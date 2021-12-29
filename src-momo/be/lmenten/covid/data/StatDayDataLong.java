package be.lmenten.covid.data;

import java.util.ArrayList;
import java.util.List;

public class StatDayDataLong
{
	private List<Long> averageValues = new ArrayList<>();

	private long min;
	private double avg;
	private long max;

	// ========================================================================
	// =
	// ========================================================================

	public StatDayDataLong()
	{
		this( 0, 0 );
	}

	public StatDayDataLong( long min, long max )
	{
		this.min = min;
		this.max = max;

		this.avg = (min + max) / 2.0d;
	}

	// ------------------------------------------------------------------------
	// - 
	// ------------------------------------------------------------------------

	public void setMin( long min )
	{
		if( min < this.min )
		{
			this.min = min;
		}
	}

	public void setMax( long max )
	{
		if( max > this.max )
		{
			this.max = max;
		}
	}

	// ------------------------------------------------------------------------

	public void addToAverageValues( long value )
	{
		averageValues.add( value );
	}

	public void resetAverageValues()
	{
		averageValues.clear();
	}

	public void computeAverage()
	{
		avg = averageValues.stream()
				.mapToDouble( val -> val )
					.average()
					.orElse(0.0)
				;
	}

	// ------------------------------------------------------------------------
	// - 
	// ------------------------------------------------------------------------

	public long getMin()
	{
		return min;
	}

	public double getAvg()
	{
		return avg;
	}

	public long getMax()
	{
		return max;
	}
}
